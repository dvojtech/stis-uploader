# stis_uploader.py
import argparse, os, re, sys, time
import unicodedata
import traceback
import ctypes
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PwTimeout

# kde EXE skutečně leží
EXE_DIR = Path(sys.argv[0]).resolve().parent
# kam určitě umíme zapsat
TEMP_DIR = Path(os.environ.get("TEMP", str(EXE_DIR)))
# cesty pro "boot" log (zapisujeme na obě místa)
BOOT_FILES = [
    TEMP_DIR / "stis_boot.log",
    EXE_DIR / "stis_boot.log",
]
# --- konfigurace mapování řádků v listu "zdroj" ---
ZDROJ_SHEET             = "zdroj"
ZDROJ_FIRST_SINGLE_ROW  = 7     # první řádek singlů (D/E = jména, I—M = sety)
SINGLES_COUNT           = 16    # kolik singlů se vyplňuje (2..17)

def wait_online_ready(page, log, timeout=25000):
    """
    Počká na načtení online editoru.
    """
    try:
        # Počkej na základní strukturu
        page.wait_for_selector("#zapis .event", state="visible", timeout=timeout)
        page.wait_for_selector(".zapas-set", state="visible", timeout=timeout)
        log("Online editor ready.")
    except Exception as e:
        log(f"Online editor se nenačetl: {repr(e)}")
        raise

def _fill_player_by_click(page, selector, name, log):
    """
    Klikne na .player-name element a vyplní jméno přes autocomplete.
    """
    name = (name or "").strip()
    if not name:
        return
        
    try:
        # 1) Klik na cílový element
        player_elem = page.locator(selector)
        if not player_elem.count():
            log(f"  Nenalezen element: {selector}")
            return
            
        player_elem.click(timeout=3000)
        page.wait_for_timeout(300)  # Krátká pauza pro aktivaci autocomplete
        
        # 2) Najdi aktivní autocomplete input
        autocomplete_selectors = [
            "input.ui-autocomplete-input:focus",
            "input.ac_input:focus", 
            "input[type='text']:focus:not(.zapas-set):not([disabled])"
        ]
        
        input_found = False
        for sel in autocomplete_selectors:
            inp = page.locator(sel)
            if inp.count():
                inp.fill(name)
                page.keyboard.press("Tab")  # Nebo Enter pro potvrzení
                page.wait_for_timeout(200)
                log(f"  ✓ {name} → {selector}")
                input_found = True
                break
                
        if not input_found:
            # Fallback: zkus napsat do aktuálně fokusovaného elementu
            page.keyboard.type(name)
            page.keyboard.press("Tab")
            log(f"  ~ {name} → {selector} (fallback)")
            
    except Exception as e:
        log(f"  ✗ {name} → {selector} failed: {repr(e)}")

def _fill_sets_by_event_index(page, event_index, sets, log):
    """
    Vyplní sety pro daný event podle jeho pozice v seznamu.
    """
    if not sets:
        return
        
    try:
        # Najdi event podle indexu
        event = page.locator(".event").nth(event_index)
        if not event.count():
            log(f"  Event #{event_index} nenalezen")
            return
            
        # Vyplň až 5 setů
        for i, value in enumerate(sets[:5]):
            if not value:
                continue
                
            set_input = event.locator(f".zapas-set[data-set='{i+1}']")
            if set_input.count():
                set_input.fill(str(_map_wo(value)))
                log(f"  set{i+1} ← {value} (event #{event_index})")
            else:
                log(f"  Set {i+1} input nenalezen pro event #{event_index}")
                
    except Exception as e:
        log(f"  Sety pro event #{event_index} selhaly: {repr(e)}")

def _map_wo(val):
    """Mapuje WO značky na STIS kódy."""
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

def fill_online_from_zdroj(page, data, log, xlsx_path=None):
    """
    Vyplní online formulář STIS podle skutečné struktury DOM.
    """
    try:
        wait_online_ready(page, log)
    except Exception:
        log("Inputs se neobjevily – dělám dump DOMu.")
        if xlsx_path:
            _dom_dump(page, xlsx_path, log)
        raise

    # ---- ČTYŘHRA #1 (ID: c0) ----
    dbl = data.get("double", {})
    
    # Domácí hráči čtyřhry
    if dbl.get("home1"):
        _fill_player_by_click(page, "#c0 .cell-player:first-child .player.domaci .player-name", dbl["home1"], log)
    if dbl.get("home2"):
        _fill_player_by_click(page, "#c0 .cell-player:last-child .player.domaci .player-name", dbl["home2"], log)
    
    # Hostující hráči čtyřhry  
    if dbl.get("away1"):
        _fill_player_by_click(page, "#c0 .cell-player:first-child .player.host .player-name", dbl["away1"], log)
    if dbl.get("away2"):
        _fill_player_by_click(page, "#c0 .cell-player:last-child .player.host .player-name", dbl["away2"], log)
    
    # Sety pro čtyřhru #1 (první .event)
    if dbl.get("sets"):
        _fill_sets_by_event_index(page, 0, dbl["sets"], log)

    # ---- SINGLY (d0 až d15) ----
    # Vaše data mají indexy 2-17, ale DOM má d0-d15, takže index-2 = DOM_ID
    for match_data in data.get("singles", []):
        excel_idx = int(match_data.get("idx", 0))  # 2, 3, 4, ... 17
        if excel_idx < 2 or excel_idx > 17:
            continue
            
        dom_idx = excel_idx - 2  # 0, 1, 2, ... 15
        event_idx = excel_idx    # pozice v seznamu eventů (čtyřhry zabírají 0,1, singly začínají od 2)
        
        # Domácí hráč
        if match_data.get("home"):
            _fill_player_by_click(page, f"#d{dom_idx} .player.domaci .player-name", match_data["home"], log)
        
        # Hostující hráč    
        if match_data.get("away"):
            _fill_player_by_click(page, f"#d{dom_idx} .player.host .player-name", match_data["away"], log)
        
        # Sety
        if match_data.get("sets"):
            _fill_sets_by_event_index(page, event_idx, match_data["sets"], log)

    # Uložit změny
    log("Klikám 'Uložit změny'…")
    try:
        page.locator("input[name='ulozit']").click(timeout=5000)
        page.wait_for_timeout(1000)
        log("Změny uloženy.")
    except Exception as e:
        log(f"Uložení selhalo: {repr(e)}")

def boot(msg: str):
    """Zapíš krátkou zprávu ještě před main() – přežije i selhání argparse."""
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n"
    for p in BOOT_FILES:
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            with open(p, "a", encoding="utf-8") as f:
                f.write(line)
        except Exception:
            pass

def msgbox(text: str, title: str="stis-uploader"):
    try:
        ctypes.windll.user32.MessageBoxW(0, str(text), str(title), 0x40)  # MB_ICONINFORMATION
    except Exception:
        pass

BOOTLOG = Path(os.environ.get("TEMP", str(Path.cwd()))) / "stis_boot.log"

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

    # 1) Najdi list a řádek hlavičky Teams kdekoli v sešitu
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

    # logger vedle XLSX
    log, log_file, log_path = make_logger(xlsx_path)
    log("==== stis_uploader start ====")
    log("XLSX:", xlsx_path)
    log("Team:", args.team)
    log("Headed:", getattr(args, "headed", True))

    try:
        # 1) načti přihlášení + tým
        user_login, user_pwd, team = read_excel_config(xlsx_path, args.team)
        log("Login OK; team:", team["name"], "ID:", team["id"])

        headed = bool(getattr(args, "headed", True))
        headless = not headed

        with sync_playwright() as p:
            ensure_pw_browsers(log)

            # 2) spuštění prohlížeče (Chromium → Chrome → Edge)
            log("Launching browser… headless =", headless)
            browser = None
            try:
                browser = p.chromium.launch(headless=headless)
                log("Launched: managed Chromium")
            except Exception as e1:
                log("Chromium failed:", repr(e1), "→ trying channel=chrome")
                try:
                    browser = p.chromium.launch(channel="chrome", headless=headless)
                    log("Launched: channel=chrome")
                except Exception as e2:
                    log("Chrome failed:", repr(e2), "→ trying channel=msedge")
                    browser = p.chromium.launch(channel="msedge", headless=headless)
                    log("Launched: channel=msedge")

            context = browser.new_context()
            page = context.new_page()

            # 3) login
            log("Navigating to login…")
            page.goto("https://registr.ping-pong.cz/htm/auth/login.php",
                      wait_until="domcontentloaded")
            page.fill("input[name='login']", user_login)
            page.fill("input[name='heslo']",  user_pwd)
            page.locator("[name='send']").click()
            page.wait_for_load_state("domcontentloaded")
            log("Logged in.")

            # 4) stránka družstva
            team_url = f"https://registr.ping-pong.cz/htm/auth/klub/druzstva/vysledky/?druzstvo={team['id']}"
            log("Open team page:", team_url)
            page.goto(team_url, wait_until="domcontentloaded")

            # 5) najdi vstup do formuláře (vložit/upravit)
            log("Hledám odkaz 'vložit/upravit zápis'…")
            if not open_match_form(page, log):
                raise RuntimeError("Na stránce družstva jsem nenašel odkaz do formuláře.")

            # 6) vyplň úvodní údaje (herna / začátek / vedoucí)
            if team.get("herna") and page.locator("input[name='zapis_herna']").count():
                page.fill("input[name='zapis_herna']", str(team["herna"]))
                log("Herna vyplněna:", team["herna"])

            # OPRAVENÉ vyplnění času pomocí dvou select elementů
            if team.get("zacatek"):
                try:
                    hh, mm = team["zacatek"].split(":")
                    # Použij správné name atributy z HTML dumpu
                    if page.locator("select[name='zapis_zacatek_hodiny']").count():
                        page.select_option("select[name='zapis_zacatek_hodiny']", value=str(int(hh)))
                        log(f"Hodina nastavena: {hh}")
                    if page.locator("select[name='zapis_zacatek_minuty']").count():
                        page.select_option("select[name='zapis_zacatek_minuty']", value=str(int(mm)))
                        log(f"Minuta nastavena: {mm}")
                    log("Začátek nastaven:", team["zacatek"])
                except Exception as e:
                    log("Set start time failed:", repr(e))

            # OPRAVENÉ vyplnění vedoucích pomocí autocomplete inputů
            if team.get("ved_dom"):
                try:
                    ved_input = page.locator("input[name='id_domaci_vedoucitext']")
                    if ved_input.count():
                        ved_input.fill(str(team["ved_dom"]))
                        page.keyboard.press("Tab")  # Aktivace autocomplete
                        page.wait_for_timeout(500)
                        log("Vedoucí domácích:", team["ved_dom"])
                except Exception as e:
                    log("Vedoucí domácích selhal:", repr(e))

            if team.get("ved_host"):
                try:
                    ved_input = page.locator("input[name='id_hoste_vedoucitext']")
                    if ved_input.count():
                        ved_input.fill(str(team["ved_host"]))
                        page.keyboard.press("Tab")  # Aktivace autocomplete
                        page.wait_for_timeout(500)
                        log("Vedoucí hostů:", team["ved_host"])
                except Exception as e:
                    log("Vedoucí hostů selhal:", repr(e))

            # 7) Uložit a pokračovat (robustně)
            log("Click 'Uložit a pokračovat'…")
            clicked = False
            
            # Zkus najít tlačítko "Uložit a pokračovat" podle name atributu z HTML dumpu
            try:
                btn = page.locator("input[name='odeslat']")  # "Uložit a pokračovat" má name="odeslat"
                if btn.count():
                    btn.click(timeout=5000)
                    clicked = True
                    log("Kliknuto na 'Uložit a pokračovat' (name=odeslat)")
            except Exception as e:
                log("Button(name=odeslat) click failed:", repr(e))
            
            # Fallback podle value atributu
            if not clicked:
                try:
                    btn = page.locator("input[value*='pokračovat']")
                    if btn.count():
                        btn.first.click(timeout=5000)
                        clicked = True
                        log("Kliknuto na 'Uložit a pokračovat' (value)")
                except Exception as e:
                    log("Button(value) click failed:", repr(e))

            if not clicked:
                raise RuntimeError("Nenašel jsem tlačítko/odkaz 'Uložit a pokračovat'.")

            # 8) Čekej na online formulář
            try:
                page.wait_for_url(re.compile(r"/online\.php\?u=\d+"), timeout=20000)
            except Exception:
                # fallback: aspoň počkat na nějaké textové inputy
                page.wait_for_selector("input[type='text']", timeout=10000)
            log("Online formulář načten:", page.url)
            
            # 9) vyplň data
            data = read_zdroj_data(xlsx_path)
            fill_online_from_zdroj(page, data, log, xlsx_path)
            
            # 10) vizuální dokončení
            if headed:
                log("Leaving browser open for manual finish.")
                print("✅ Online formulář načten – dokonči ručně. Okno nechávám otevřené.")
                while True:
                    time.sleep(1)
            else:
                context.close()
                browser.close()

    except Exception as e:
        log("ERROR:", repr(e))
        log(traceback.format_exc())
        try:
            os.startfile(str(log_path))  # otevřít log v Notepadu
        except Exception:
            pass
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
        main()
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
