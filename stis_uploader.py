# stis_uploader.py
import argparse, os, re, sys, time, unicodedata, traceback, ctypes
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PwTimeout

# kde EXE skutečně leží
EXE_DIR = Path(sys.argv[0]).resolve().parent
# kam určitě umíme zapsat
TEMP_DIR = Path(os.environ.get("TEMP", str(EXE_DIR)))
# cesty pro „boot“ log (zapisujeme na obě místa)
BOOT_FILES = [
    TEMP_DIR / "stis_boot.log",
    EXE_DIR / "stis_boot.log",
]
# --- konfigurace mapování řádků v listu "zdroj" ---
ZDROJ_SHEET             = "zdroj"
ZDROJ_FIRST_SINGLE_ROW  = 7     # první řádek singlů (D/E = jména, I–M = sety)
SINGLES_COUNT           = 16    # kolik singlů se vyplňuje (2..17)


def _log(msg):
    # bezpecné logování do konzole i souboru (pokud máš vlastní log funkci, klidně ji zde použij)
    try:
        print(msg, flush=True)
    except Exception:
        pass

def _parse_hhmm(s: str) -> tuple[int, int]:
    """Vrátí (hod, min) z '19:00', '19.00', '19 00', '19', …  Minuty zaokrouhlí na 5."""
    if not s:
        return (0, 0)
    txt = str(s).strip()
    m = re.findall(r"\d+", txt)
    if not m:
        return (0, 0)
    if len(m) == 1:
        hh, mm = int(m[0]), 0
    else:
        hh, mm = int(m[0]), int(m[1])
    hh = max(0, min(23, hh))
    mm = max(0, min(59, mm))
    # STIS nabízí jen 00,05,10,…55 → zaokrouhlíme na nejbližších 5
    mm = int(round(mm / 5) * 5) % 60
    return (hh, mm)

def set_start_time(page, start_val, log, timeout_each=1500):
    """Vybere HH/MM z dropdownů i když mění jména/id (nejdřív známé selektory, pak heuristika)."""
    hh, mm = _parse_hhmm(start_val)

    tried = [
        ("select[name='zapis_zacatek_hod']", "select[name='zapis_zacatek_min']"),
        ("#zapis_zacatek_hod", "#zapis_zacatek_min"),
        ("select[name='zacatek_hod']", "select[name='zacatek_min']"),
    ]
    for hs, ms in tried:
        h = page.locator(hs); m = page.locator(ms)
        try:
            if h.count() and m.count():
                h.first.wait_for(state="visible", timeout=timeout_each)
                m.first.wait_for(state="visible", timeout=timeout_each)
                for tgt, val in ((h.first, hh), (m.first, mm)):
                    try: tgt.select_option(label=val)
                    except Exception: tgt.select_option(value=val)
                log(f"Začátek nastaven přes {hs}/{ms} → {hh}:{mm}")
                return
        except PwTimeout:
            pass

    # heuristika: najdi dvě <select> s ~24 a ~60 možnostmi
    cand_h, cand_m = [], []
    for i in range(page.locator("select").count()):
        s = page.locator("select").nth(i)
        try:
            n = s.evaluate("el => el.options ? el.options.length : 0")
            if 22 <= n <= 26: cand_h.append(s)
            elif 57 <= n <= 63: cand_m.append(s)
        except Exception:
            continue
    if cand_h and cand_m:
        h, m = cand_h[0], cand_m[0]
        h.wait_for(state="visible", timeout=timeout_each)
        m.wait_for(state="visible", timeout=timeout_each)
        for tgt, val in ((h, hh), (m, mm)):
            try: tgt.select_option(label=val)
            except Exception: tgt.select_option(value=val)
        log(f"Začátek nastaven heuristicky → {hh}:{mm}")
        return

    raise RuntimeError("Nepodařilo se nastavit 'Začátek utkání' – nenašel jsem časové <select>.")

def fill_start_form(page, herna: str, zacatek: str, ved_dom: str, ved_host: str, log):
    """Vyplní úvodní formulář (online_start.php) a přejde na online.php."""
    # Herna (textové pole)
    if herna:
        page.locator("input[name='zapis_herna']").fill(herna)
        log(f"Herna vyplněna: {herna}")

    # Začátek utkání – dvě <select> s name='zapis_zacatek_hodiny' / '..._minuty'
    hh, mm = _parse_hhmm(zacatek)
    page.select_option("select[name='zapis_zacatek_hodiny']", value=str(hh))
    page.select_option("select[name='zapis_zacatek_minuty']", value=str(mm))
    log(f"Začátek vyplněn: {hh:02d}:{mm:02d}")

    # Vedoucí družstev – použijeme „fallback“: doplníme text do viditelných inputů
    if ved_dom:
        page.locator("input[name='id_domaci_vedoucitext']").fill(ved_dom)
        log(f"Vedoucí 'Domácí:' → {ved_dom}")
    if ved_host:
        page.locator("input[name='id_hoste_vedoucitext']").fill(ved_host)
        log(f"Vedoucí 'Hosté:' → {ved_host}")

    # Uložit a pokračovat
    page.locator("input[name='odeslat']").click()

    # Po odeslání buď zůstaneme na online_start.php (chyba), nebo se přejde na online.php
    # Nejprve krátce počkáme na případnou chybovou hlášku:
    page.wait_for_load_state("networkidle")
    url = page.url
    if "online_start.php" in url:
        # Zkusíme přečíst případnou chybu
        err = page.locator(".exception").first
        if err.is_visible():
            log(f"Po odeslání zůstávám na startu – hláška: {err.inner_text()}")
        else:
            log("Po odeslání zůstávám na startu – bez hlášky.")
        # Abychom nepokračovali do části pro online.php:
        raise RuntimeError("Nepodařilo se přejít na online formulář (zkontroluj vyplnění času).")

    # Jsme na online.php – ještě počkáme, až se objeví blok s utkáním
    page.wait_for_selector("#zapis .event", timeout=25000)
def _fill_input_by_names_or_label_row(page, names, row_label, value, log):
    # 1) zkus známé name/id
    for css in names:
        loc = page.locator(css)
        if loc.count():
            loc.first.fill(str(value))
            log(f"Vedoucí '{row_label}' vyplněn přes {css}: {value}")
            return True

    # 2) fallback: najdi řádek s textem „Domácí:“/„Hosté:“ a vezmi první text input
    lab = page.locator(f"text={row_label}").first
    if lab.count():
        cont = lab.locator("xpath=ancestor::*[self::tr or self::div][1]")
        inp = cont.locator("input[type='text']").first
        if inp.count():
            inp.fill(str(value))
            log(f"Vedoucí '{row_label}' vyplněn fallbackem: {value}")
            return True
    return False

def set_team_leaders(page, home_leader, away_leader, log, set_home_only=True, set_away_only=True):
    """Vyplní vedoucí družstev. Pokusí se i zaškrtnout 'Jen z oddílu', pokud je checkbox poblíž."""
    if home_leader:
        _fill_input_by_names_or_label_row(
            page,
            ["input[name='zapis_domaci_vedouci']",
             "input[name='vedouci_domaci']",
             "#zapis_domaci_vedouci"],
            "Domácí:", home_leader, log
        )
        if set_home_only:
            # checkbox ve stejném řádku
            try:
                lab = page.locator("text=Domácí:").first
                cont = lab.locator("xpath=ancestor::*[self::tr or self::div][1]")
                chk = cont.locator("input[type='checkbox']").first
                if chk.count() and not chk.is_checked():
                    chk.check()
            except Exception:
                pass

    if away_leader:
        _fill_input_by_names_or_label_row(
            page,
            ["input[name='zapis_hoste_vedouci']",
             "input[name='vedouci_hoste']",
             "#zapis_hoste_vedouci"],
            "Hosté:", away_leader, log
        )
        if set_away_only:
            try:
                lab = page.locator("text=Hosté:").first
                cont = lab.locator("xpath=ancestor::*[self::tr or self::div][1]")
                chk = cont.locator("input[type='checkbox']").first
                if chk.count() and not chk.is_checked():
                    chk.check()
            except Exception:
                pass

def wait_online_ready(page, log, timeout=25000):
    # počkej na řádek zápasu a aspoň jeden vstup na sety
    page.wait_for_selector("#zapis .event", state="visible", timeout=timeout)
    page.wait_for_selector("#zapis .event .event-sety input.zapas-set", state="visible", timeout=timeout)
    log("Online editor ready.")

def _choose_player_loc(page, loc, name, log, timeout_each=2000):
    """
    Klikne do buňky hráče a vyplní jméno přes autocomplete.
    Zkouší několik vzorů a nakonec i :focus fallback. Ignoruje disabled/skryté inputy.
    """
    name = (name or "").strip()
    if not name:
        return

    # 1) klik do cílové buňky (span/div se jménem)
    loc.click()
    page.wait_for_timeout(80)

    # 2) kandidáti na autocomplete input
    candidates = [
        "input.ui-autocomplete-input",
        "input.ac_input",
        # obecný viditelný text input, ale NE sety ani skóre, NE disabled
        "input[type='text']:not(.zapas-set):not(.utkani-skore):not([name='body']):not([disabled])"
    ]

    for css in candidates:
        try:
            inp = page.locator(css).first
            inp.wait_for(state="visible", timeout=timeout_each)
            # bezpečně smaž a napiš
            inp.click()
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            inp.fill(name)
            page.keyboard.press("Enter")
            page.wait_for_timeout(120)
            log(f"  player ← {name} (via {css})")
            return
        except PwTimeout:
            # nenašli jsme viditelný input tohoto typu – zkus další
            pass
        except Exception as e:
            log(f"  candidate {css} failed: {repr(e)}")

    # 3) fallback: aktuálně fokusovaný element (většinou autocomplete input)
    try:
        foc = page.locator(":focus")
        if foc.count() > 0:
            tag  = foc.evaluate("el => el.tagName.toLowerCase()")
            typ  = foc.evaluate("el => el.type || ''")
            dis  = foc.evaluate("el => !!el.disabled")
            if tag == "input" and typ in ("text", "search") and not dis:
                foc.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                foc.fill(name)
                page.keyboard.press("Enter")
                page.wait_for_timeout(120)
                log(f"  player ← {name} (via :focus)")
                return
    except Exception as e:
        log(f"  :focus fallback failed: {repr(e)}")

    # 4) poslední nouze – napiš naslepo do aktivního prvku
    try:
        page.keyboard.type(name)
        page.keyboard.press("Enter")
        log(f"  player ← {name} (typed)")
        return
    except Exception:
        pass

    # když nic nevyšlo, ať je to vidět v logu i výjimkou
    raise RuntimeError(f"Autocomplete input pro hráče '{name}' se nenašel (všechny varianty selhaly).")

def _fill_sets_for_event(event_loc, sets, log):
    """Vyplní až 5 setů v daném řádku .event."""
    inputs = event_loc.locator(".event-sety input.zapas-set")
    for i in range(min(5, len(sets))):
        v = sets[i]
        if v in (None, ""):
            continue
        try:
            inputs.nth(i).fill(str(v))
            log(f"  set{i+1} ← {v}")
        except Exception as e:
            log(f"  set{i+1} fill failed: {repr(e)}")

def _map_wo(val):
    """WO 3:0 → 101, WO 0:3 → -101, jinak původní hodnota."""
    if not val:
        return ""
    s = str(val).strip().upper().replace(" ", "")
    if s in ("WO3:0", "3:0WO"):
        return "101"
    if s in ("WO0:3", "0:3WO"):
        return "-101"
    return val

def _dom_dump(page, xlsx_path, log):
    try:
        html_path = Path(xlsx_path).with_suffix(".online_dump.html")
        png_path  = Path(xlsx_path).with_suffix(".online_dump.png")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(page.content())
        page.screenshot(path=str(png_path), full_page=True)
        log(f"DOM dump → {html_path.name}, screenshot → {png_path.name}")
    except Exception as e:
        log(f"DOM dump failed: {repr(e)}")


def cnt(page, css):  # krátká pomůcka do logu
    try:
        return page.locator(css).count()
    except Exception:
        return -1
def map_wo(s):
    """mapování WO značek na STIS kódy, jinak vrací čistý text/číslo"""
    t = str(s or "").strip()
    if not t:
        return ""
    if t.upper() in ("WO 3:0", "WO 3:0", "WO3:0"):
        return "101"
    if t.upper() in ("WO 0:3", "WO 0:3", "WO0:3"):
        return "-101"
    return t
def _input_or_select(page, base, idx):
    """
    Vrátí locator na input/select pro dané pole. Bere:
      - exact: input[name='base_idx'] / select[name='base_idx']
      - tolerantní: input[name^='base'][name$='_{idx}'] / select[...]
    """
    exact = f"[name='{base}_{idx}']"
    loose = f"[name^='{base}'][name$='_{idx}']"
    sel = f"input{exact}, select{exact}, input{loose}, select{loose}"
    loc = page.locator(sel)
    if loc.count():
        return loc.first
    return None

def _fill_name(page, base, idx, value, log):
    if not value:
        return
    loc = _input_or_select(page, base, idx)
    if not loc:
        log(f"✗ nenalezeno pole {base}_{idx}")
        return
    try:
        tag = loc.evaluate("e => e.tagName.toLowerCase()")
        if tag == "select":
            loc.select_option(label=str(value))
        else:
            loc.fill(str(value))
        log(f"✓ {base}_{idx} ← {value}")
    except Exception as e:
        log(f"✗ {base}_{idx} fill failed:", repr(e))

def _fill_set(page, set_no, idx, value, log):
    if not value: 
        return
    base = f"set{set_no}"
    loc = _input_or_select(page, base, idx)
    if not loc:
        log(f"✗ nenalezeno pole {base}_{idx}")
        return
    try:
        loc.fill(str(value))
        log(f"✓ {base}_{idx} ← {value}")
    except Exception as e:
        log(f"✗ {base}_{idx} fill failed:", repr(e))

def read_zdroj_data(xlsx_path):
    """Načte čtyřhru a singly z listu 'zdroj' – podle výše uvedených konstant."""
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, data_only=True)
    if ZDROJ_SHEET not in wb.sheetnames:
        raise RuntimeError(f"V sešitu chybí list '{ZDROJ_SHEET}'")
    sh = wb[ZDROJ_SHEET]

    def cell(r, c):
        return str(sh.cell(r, c).value or "").strip()

    # ČTYŘHRA – jména (D2,D3,E2,E3) a sety (I3..M3)
    double = {
        "home1": cell(2, 4),  # D2
        "home2": cell(3, 4),  # D3
        "away1": cell(2, 5),  # E2
        "away2": cell(3, 5),  # E3
        "sets":  [ map_wo(cell(3, 9+i)) for i in range(5) ]  # I3..M3
    }

    # SINGLY – indexy 2..(2+SINGLES_COUNT-1) → řádky od ZDROJ_FIRST_SINGLE_ROW
    singles = []
    row = ZDROJ_FIRST_SINGLE_ROW
    for idx in range(2, 2 + SINGLES_COUNT):
        home = cell(row, 4)  # D
        away = cell(row, 5)  # E
        sets = [ map_wo(cell(row, 9+i)) for i in range(5) ]  # I..M
        singles.append({"idx": idx, "home": home, "away": away, "sets": sets})
        row += 1

    return {"double": double, "singles": singles}

def boot(msg: str):
    """Zapiš krátkou zprávu ještě před main() – přežije i selhání argparse."""
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n"
    for p in BOOT_FILES:
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            with open(p, "a", encoding="utf-8") as f:
                f.write(line)
        except Exception:
            pass

def fill_online_from_zdroj(page, data, log, xlsx_path=None):
    """
    data: dict z read_zdroj_data()
      - data['double']: {'home1','home2','away1','away2','sets':[...]}
      - data['singles']: list({'idx':2..17,'home','away','sets':[...]}).
    Vyplní online formulář STIS (DOM s #c0,#c1 a #d0..#d15).
    """
    try:
        wait_online_ready(page, log)
    except Exception:
        log("Inputs se neobjevily – dělám dump DOMu.")
        if xlsx_path:
            _dom_dump(page, xlsx_path, log)
        raise

    events = page.locator("#zapis .event")

    # ---- ČTYŘHRY (#c0, #c1) ----
    dbl = data.get("double", {})
    # #c0
    c0_dom = page.locator("#c0 .player.domaci .player-name")
    c0_hst = page.locator("#c0 .player.host .player-name")
    _choose_player_loc(page, c0_dom.nth(0), dbl.get("home1"), log)
    _choose_player_loc(page, c0_dom.nth(1), dbl.get("home2"), log)
    _choose_player_loc(page, c0_hst.nth(0), dbl.get("away1"), log)
    _choose_player_loc(page, c0_hst.nth(1), dbl.get("away2"), log)
    sets0 = [ _map_wo(x) for x in (dbl.get("sets") or []) ] + [""]*5
    _fill_sets_for_event(events.nth(0), sets0, log)

    # #c1 – pokud ve vstupu máš druhou čtyřhru, přidej ji do data['double2']
    dbl2 = data.get("double2")
    if dbl2:
        c1_dom = page.locator("#c1 .player.domaci .player-name")
        c1_hst = page.locator("#c1 .player.host .player-name")
        _choose_player_loc(page, c1_dom.nth(0), dbl2.get("home1"), log)
        _choose_player_loc(page, c1_dom.nth(1), dbl2.get("home2"), log)
        _choose_player_loc(page, c1_hst.nth(0), dbl2.get("away1"), log)
        _choose_player_loc(page, c1_hst.nth(1), dbl2.get("away2"), log)
        sets1 = [ _map_wo(x) for x in (dbl2.get("sets") or []) ] + [""]*5
        _fill_sets_for_event(events.nth(1), sets1, log)

    # ---- DVOUHRY (#d0 .. #d15) ----
    for m in (data.get("singles") or []):
        idx = int(m.get("idx", 0))
        if idx < 2 or idx > 17:
            continue
        d_id = f"#d{idx-2}"
        ev   = events.nth(idx)
        _choose_player_loc(page, page.locator(f"{d_id} .player.domaci .player-name"), m.get("home"), log)
        _choose_player_loc(page, page.locator(f"{d_id} .player.host .player-name"),   m.get("away"), log)
        setsS = [ _map_wo(x) for x in (m.get("sets") or []) ] + [""]*5
        _fill_sets_for_event(ev, setsS, log)

    log("Klikám 'Uložit změny'…")
    page.locator("input[name='ulozit']").last.click()
    page.wait_for_timeout(600)
    
def msgbox(text: str, title: str="stis-uploader"):
    try:
        ctypes.windll.user32.MessageBoxW(0, str(text), str(title), 0x40)  # MB_ICONINFORMATION
    except Exception:
        pass


BOOTLOG = Path(os.environ.get("TEMP", str(Path.cwd()))) / "stis_boot.log"

def boot(msg: str):
    try:
        with open(BOOTLOG, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
    except Exception:
        pass

def make_logger(xlsx_path: Path):
    """Vrátí (log_fn, file_handle, log_path) – loguje s časovou značkou."""
    log_path = xlsx_path.with_suffix(".stislog.txt")
    f = open(log_path, "a", encoding="utf-8")
    def log(*parts):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = " ".join(str(p) for p in parts)
        f.write(f"[{ts}] {line}\n")
        f.flush()
    return log, f, log_path

def ensure_pw_browsers(log=None):
    """Stáhne Chromium pro Playwright, pokud chybí (funguje i z PyInstaller EXE)."""
    default_store = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "ms-playwright"
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(default_store))

    chromium_ok = False
    if default_store.exists():
        for p in default_store.glob("chromium-*/*/chrome.exe"):
            if p.exists():
                chromium_ok = True
                break
    if chromium_ok:
        if log: log("Chromium already present in", default_store)
        return

    if log: log("Chromium not found – starting Playwright install (can take minutes)…")
    import playwright.__main__ as pw_cli
    old_argv = sys.argv[:]
    try:
        sys.argv = ["playwright", "install", "chromium"]
        pw_cli.main()   # CLI bere argumenty z sys.argv
        if log: log("Playwright install finished.")
    finally:
        sys.argv = old_argv


def _safe_fill(page, sel, value, log):
    if not value:
        return
    loc = page.locator(sel)
    if loc.count():
        try:
            loc.first.fill(str(value))
            log("fill", sel, "→", value)
        except Exception as e:
            log("fill FAILED", sel, repr(e))

def _fill_sets(page, base_idx, sets, log):
    for i, val in enumerate(sets, start=1):
        if not val: 
            continue
        sel = f"input[name='set{i}_{base_idx}']"
        _safe_fill(page, sel, val, log)


def norm(x) -> str:
    s = "" if x is None else str(x)
    s = s.strip().lower()
    # odstraň diakritiku
    s = "".join(ch for ch in unicodedata.normalize("NFD", s)
                if unicodedata.category(ch) != "Mn")
    # bez mezer a nealfanumerických znaků
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^0-9a-z_]", "", s)
    return s



def as_time_txt(v):
    if v is None: return None
    try:
        if isinstance(v, (int, float)):
            total = round(float(v) * 24 * 60)
            return f"{total//60:02d}:{total%60:02d}"
    except: pass
    s = str(v).strip()
    m = re.match(r"^(\d{1,2}):(\d{2})$", s)
    return f"{int(m.group(1)):02d}:{int(m.group(2)):02d}" if m else None

def get_setup_sheet(wb):
    """Najdi list 'setup' case-insensitive, jinak vrať první list."""
    for ws in wb.worksheets:
        if (ws.title or "").strip().lower() == "setup":
            return ws
    return wb.active

def find_login_pwd(ws):
    """Vrátí login/heslo. Nejprve z B1/B2, když chybí, zkusí popisky 'login' / 'heslo' v horní části listu."""
    login = str(ws["B1"].value or "").strip()
    pwd   = str(ws["B2"].value or "").strip()
    if login and pwd:
        return login, pwd

    max_r = min(30, ws.max_row or 0)
    max_c = min(20, ws.max_column or 0)
    found_login = found_pwd = ""
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = norm(ws.cell(r, c).value or "")
            if v == "login" and c + 1 <= (ws.max_column or 0):
                found_login = str(ws.cell(r, c + 1).value or "").strip()
            if v == "heslo" and c + 1 <= (ws.max_column or 0):
                found_pwd = str(ws.cell(r, c + 1).value or "").strip()
    return found_login, found_pwd


def find_teams_header_anywhere(wb):
    """
    Projdi všechny listy a najdi první řádek, kde je hlavička tabulky Teams:
      - obsahuje 'druzstvo' (ale ne 'id' / 'vedouci')
      - a zároveň některý JASNÝ marker ID ('druzstvoid', 'id_druzstva', 'iddruzstva') nebo PŘESNĚ 'id'
    Vrací (sheet, hdr_row) nebo (None, None).
    """
    strict_id_markers = ("druzstvoid", "id_druzstva", "iddruzstva")  # žádné volné "id" jako podřetězec!
    for ws in wb.worksheets:
        max_r = min(60, ws.max_row or 0)
        max_c = min(80, ws.max_column or 0)
        for r in range(1, max_r + 1):
            row_norm = [norm(ws.cell(r, c).value or "") for c in range(1, max_c + 1)]
            has_name = any(("druzstvo" in v and "id" not in v and "vedouci" not in v) for v in row_norm)
            has_id   = any(any(m in v for m in strict_id_markers) or (v == "id") for v in row_norm)
            if has_name and has_id:
                return ws, r
    return None, None

def open_match_form(page, log):
    """Na stránce družstva otevře formulář – preferuje 'vložit zápis', jinak 'upravit zápis'.
       Zkouší text i href (zapis_start.php / online.php). Vrací True/False.
    """
    try:
        page.wait_for_selector(
            "a:has-text('vložit zápis'), a:has-text('upravit zápis'), "
            "a[href*='zapis_start.php?u='], a[href*='online.php?u=']",
            timeout=15000
        )
    except Exception:
        log("Nenalezl jsem žádný z očekávaných odkazů do 15 s.")
        return False

    candidates = [
        page.get_by_role("link", name=re.compile(r"vložit\s*zápis", re.I)),
        page.get_by_role("link", name=re.compile(r"upravit\s*zápis", re.I)),
        page.locator("a[href*='zapis_start.php?u=']"),
        page.locator("a[href*='online.php?u=']"),
    ]

    for i, sel in enumerate(candidates, 1):
        try:
            if sel.count():
                log(f"Zkouším selector {i}")
                sel.first.click(timeout=5000)
                page.wait_for_load_state("domcontentloaded")

                if page.locator("text=/vkládání zápisu/i").count() \
                   or "online.php?u=" in page.url \
                   or "zapis_start.php?u=" in page.url:
                    log("Formulář otevřen na URL:", page.url)
                    return True

                if page.locator("text=/špatn.*url/i").count():
                    log("Server hlásí 'špatné URL' – zkusím jiný odkaz.")
                    page.go_back()
                    page.wait_for_load_state("domcontentloaded")
        except Exception as e:
            log(f"Selector {i} selhal:", repr(e))

    # poslední záchrana – proskenuj všechny <a> a skoč přímo na vyhovující href
    try:
        anchors = page.eval_on_selector_all(
            "a",
            "els => els.map(a => ({text: (a.textContent||'').trim(), href: a.href||''}))"
        )
        for a in anchors:
            if re.search(r"(zapis_start\.php|online\.php)\?u=\d+", a.get("href",""), re.I):
                log("Jdu přímo na", a["href"])
                page.goto(a["href"], wait_until="domcontentloaded")
                return True
    except Exception as e:
        log("Fallback scan anchorů selhal:", repr(e))

    return False

def read_excel_config(xlsx_path: Path, team_name: str):
    """
    Najde list s tabulkou Teams kdekoli v sešitu, přihlášení bere z B1/B2
    (nebo z popisků 'login'/'heslo'), namapuje sloupce a vrátí login, heslo a dict týmu.
    """
    wb = load_workbook(xlsx_path, data_only=True)

    # 1) Najdi list a řádek hlavičky Teams kdekoliv v sešitu
    setup, hdr_row = find_teams_header_anywhere(wb)
    if setup is None:
        # diagnostika – co jsme nahoře viděli
        try:
            with open(Path(xlsx_path).with_suffix(".log"), "w", encoding="utf-8") as f:
                f.write("Header not found. Top rows (normalized):\n")
                for ws in wb.worksheets:
                    f.write(f"[{ws.title}]\n")
                    max_r = min(15, ws.max_row or 0)
                    max_c = min(20, ws.max_column or 0)
                    for r in range(1, max_r + 1):
                        rn = [norm(ws.cell(r, c).value or "") for c in range(1, max_c + 1)]
                        f.write(f"R{r}: {rn}\n")
        except Exception:
            pass
        raise RuntimeError("V setup/Teams chybí sloupce 'Družstvo' a/nebo 'DruzstvoID'.")

    # 2) Login a heslo (z nalezeného listu)
    login, pwd = find_login_pwd(setup)
    if not login or not pwd:
        raise RuntimeError("Vyplň login/heslo (B1/B2, nebo vedle popisků 'login'/'heslo').")

    # 3) Namapuj sloupce podle hlavičky (ID sloupec detekuj STRIKTNĚ + vyluč 'vedouci*')
    max_c = min(80, setup.max_column or 0)
    idx = {}
    for c in range(1, max_c + 1):
        h = norm(setup.cell(hdr_row, c).value or "")

        # Družstvo (název)
        if ("druzstvo" in h) and ("id" not in h) and ("vedouci" not in h):
            idx["name"] = c

        # ID družstva – akceptuj jasné varianty a/nebo přesné 'id', nikdy ne cokoliv s 'vedouci'
        is_id_col = (("druzstvoid" in h) or ("id_druzstva" in h) or ("iddruzstva" in h) or (h == "id")) \
                    and ("vedouci" not in h)
        if is_id_col:
            idx["id"] = c

        # Vedoucí domácích / hostů
        if "vedoucidomacich" in h or ("vedouci" in h and "host" not in h):
            idx["ved_dom"] = c
        if "vedoucihostu" in h or ("vedouci" in h and "host" in h):
            idx["ved_host"] = c

        # Herna, začátek, konec
        if "herna" in h:
            idx["herna"] = c
        if "zacatekut" in h or "zacatek" in h:
            idx["zacatek"] = c
        if "konecutkani" in h or "konec" in h:
            idx["konec"] = c

    if "name" not in idx or "id" not in idx:
        raise RuntimeError("V Teams chybí sloupce 'Družstvo' a/nebo 'DruzstvoID'.")

    # 4) Najdi požadované družstvo pod hlavičkou
    team = None
    r = hdr_row + 1
    while r <= (setup.max_row or 0):
        nm = setup.cell(r, idx["name"]).value
        if nm is None or str(nm).strip() == "":
            break
        if str(nm).strip().lower() == team_name.strip().lower():
            def getcol(key):
                c = idx.get(key)
                return setup.cell(r, c).value if c else None
            team = {
                "name":    str(nm).strip(),
                "id":      str(getcol("id") or "").strip(),
                "ved_dom": getcol("ved_dom"),
                "ved_host":getcol("ved_host"),
                "herna":   getcol("herna"),
                "zacatek": as_time_txt(getcol("zacatek")),
                "konec":   as_time_txt(getcol("konec")),
            }
            break
        r += 1

    if not team:
        raise RuntimeError(f"Družstvo '{team_name}' nenalezeno v Teams.")
    if not team["id"]:
        raise RuntimeError("Prázdné DruzstvoID u zvoleného družstva.")

    # 5) Očisti ID na čisté číslo (zabráníme tomu, aby se do URL dostalo 'Kozel Petr' apod.)
    raw_id = str(team.get("id", "")).strip()
    m = re.search(r"\d+", raw_id)
    if not m:
        raise RuntimeError(f"Neplatné DruzstvoID: {raw_id!r}")
    team["id"] = m.group(0)

    return login, pwd, team


def click_save_and_continue(page):
    sels = [
        "input[type=submit][value*='pokrač']",
        "input[type=submit][value*='pokrac']",
        "button:has-text('pokrač')",
        "button:has-text('pokrac')",
        "input[name='odeslat'][value*='pokrač']",
    ]
    for sel in sels:
        if page.locator(sel).count():
            page.locator(sel).first.click()
            return True
    subs = page.locator("form input[type=submit]")
    if subs.count() >= 2:
        subs.nth(1).click()
        return True
    return False


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--xlsx", required=True, help="plná cesta k XLSX")
    p.add_argument("--team", required=True, help="název družstva (sloupec 'Družstvo')")
    g = p.add_mutually_exclusive_group()
    g.add_argument("--headed",  dest="headed",  action="store_true",  help="viditelný prohlížeč")
    g.add_argument("--headless", dest="headed", action="store_false", help="bez UI")
    p.set_defaults(headed=True)  # výchozí = viditelné okno
    return p.parse_args()



def main():
    args = parse_args()
    xlsx_path = Path(args.xlsx).resolve()
    if not xlsx_path.exists():
        raise RuntimeError(f"Soubor neexistuje: {xlsx_path}")

    log, log_file, log_path = make_logger(xlsx_path)
    log("==== stis_uploader start ====")
    log("XLSX:", xlsx_path)
    log("Team:", args.team)
    log("Headed:", getattr(args, "headed", True))

    try:
        login, pwd, team = read_excel_config(xlsx_path, args.team)
        log("Login OK; team:", team["name"], "ID:", team["id"])

        headed = bool(getattr(args, "headed", True))

        with sync_playwright() as p:
            ensure_pw_browsers(log)
            # --- spuštění prohlížeče
            log("Launching browser… headless =", not headed)
            try:
                browser = p.chromium.launch(headless=not headed)
                log("Launched: managed Chromium")
            except Exception as e1:
                log("Chromium failed:", repr(e1), "→ trying channel=chrome")
                try:
                    browser = p.chromium.launch(channel="chrome", headless=not headed)
                    log("Launched: channel=chrome")
                except Exception as e2:
                    log("Chrome failed:", repr(e2), "→ trying channel=msedge")
                    browser = p.chromium.launch(channel="msedge", headless=not headed)
                    log("Launched: channel=msedge")

            ctx = browser.new_context()
            page = ctx.new_page()

            # --- login
            log("Navigating to login…")
            page.goto("https://registr.ping-pong.cz/htm/auth/login.php", wait_until="domcontentloaded")
            page.fill("input[name='login']", login)
            page.fill("input[name='heslo']",  pwd)
            page.locator("[name='send']").click()
            page.wait_for_load_state("domcontentloaded")
            log("Logged in.")

            # --- stránka družstva
            team_url = f"https://registr.ping-pong.cz/htm/auth/klub/druzstva/vysledky/?druzstvo={team['id']}"
            log("Open team page:", team_url)
            page.goto(team_url, wait_until="domcontentloaded")

            # --- vstup do formuláře (vložit/upravit zápis)
            log("Hledám odkaz 'vložit/upravit zápis'…")
            if not open_match_form(page, log):
                raise RuntimeError("Na stránce družstva jsem nenašel odkaz do formuláře.")

            # --- pokud jsme na online_start.php → vyplň Herna/Začátek/Vedoucí a pokračuj
            if "online_start.php" in page.url:
                fill_start_form(
                    page,
                    herna   = str(team.get("herna")   or ""),
                    zacatek = str(team.get("zacatek") or ""),
                    ved_dom = str(team.get("ved_dom") or ""),
                    ved_host= str(team.get("ved_host")or ""),
                    log     = log
                )
            else:
                # (už rozpracovaný zápis skočil rovnou do online.php)
                log("Přeskočeno online_start – jsem rovnou na:", page.url)

            # --- teď už MUSÍME být na online.php
            log("Online formulář načten:", page.url)

            # --- načti data z listu 'zdroj' a vyplň
            data = read_zdroj_data(xlsx_path)
            fill_online_from_zdroj(page, data, log, xlsx_path)

            # volitelně uložit
            try:
                page.get_by_role("button", name=re.compile(r"uložit\s*změny", re.I)).first.click(timeout=4000)
                log("Kliknuto na 'Uložit změny'.")
            except Exception as e:
                log("'Uložit změny' nenašlo/nekliklo:", repr(e))

            # nech prohlížeč otevřený pro vizuální kontrolu
            if headed:
                log("Leaving browser open for manual finish.")
                while True:
                    time.sleep(1)
            else:
                ctx.close()
                browser.close()

    except Exception as e:
        log("ERROR:", repr(e))
        log(traceback.format_exc())
        try: os.startfile(str(log_path))
        except Exception: pass
        raise
    finally:
        try:
            log("==== stis_uploader end ====")
            log_file.close()
        except Exception:
            pass

if __name__ == "__main__":
    boot("=== EXE start ===")
    try:
        boot("argv: " + " ".join(sys.argv))
        main()  # tvoje stávající main() – neměň
        boot("main() finished OK")
    except SystemExit as e:
        boot(f"SystemExit (pravděpodobně argparse): code={getattr(e, 'code', None)}")
        msgbox("Spuštění skončilo hned na začátku (špatné/neúplné argumenty?).\n" +
               "Zkontroluj prosím volání z Excelu.\n" +
               "V TEMP nebo vedle EXE je stis_boot.log s detaily.")
        raise
    except Exception as e:
        boot("CRASH: " + repr(e))
        try:
            boot(traceback.format_exc())
        except Exception:
            pass
        msgbox(f"Nastala chyba: {e}\nPodrobnosti jsou ve stis_boot.log.")
        raise
    finally:
        boot("=== EXE end ===")
