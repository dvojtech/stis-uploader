# stis_uploader.py
import argparse, os, re, sys, time
import unicodedata
import traceback
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

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
      - a zároveň některý marker ID ('druzstvoid', 'druzstvoid', 'iddruzstva', 'id')
    Vrací (sheet, hdr_row) nebo (None, None).
    """
    id_markers = ("druzstvoid", "druzstvoid", "iddruzstva", "id")
    for ws in wb.worksheets:
        max_r = min(60, ws.max_row or 0)
        max_c = min(80, ws.max_column or 0)
        for r in range(1, max_r + 1):
            row_norm = [norm(ws.cell(r, c).value or "") for c in range(1, max_c + 1)]
            has_name = any(("druzstvo" in v and "id" not in v and "vedouci" not in v) for v in row_norm)
            has_id   = any(any(m in v for m in id_markers) for v in row_norm)
            if has_name and has_id:
                return ws, r
    return None, None

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

    # 3) Namapuj sloupce podle hlavičky
    id_markers = ("druzstvoid", "druzstvoid", "iddruzstva", "id")
    max_c = min(80, setup.max_column or 0)
    idx = {}
    for c in range(1, max_c + 1):
        h = norm(setup.cell(hdr_row, c).value or "")

        if ("druzstvo" in h) and ("id" not in h) and ("vedouci" not in h):
            idx["name"] = c
        if any(m in h for m in id_markers):
            idx["id"] = c
        if "vedoucidomacich" in h or ("vedouci" in h and "host" not in h):
            idx["ved_dom"] = c
        if "vedoucihostu" in h or ("vedouci" in h and "host" in h):
            idx["ved_host"] = c
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

    # zřídíme logger
    log, log_file, log_path = make_logger(xlsx_path)
    log("==== stis_uploader start ====")
    log("XLSX:", xlsx_path)
    log("Team:", args.team)
    log("Headed:", getattr(args, "headed", True))

    try:
        if not xlsx_path.exists():
            raise RuntimeError(f"Soubor neexistuje: {xlsx_path}")

        # --- načti přihlašovací údaje a družstvo z Excelu ---
        user_login, user_pwd, team = read_excel_config(xlsx_path, args.team)
        log("Login OK, team:", team["name"], "ID:", team["id"])

        # --- režim prohlížeče ---
        headed = bool(getattr(args, "headed", True))
        headless = not headed

        with sync_playwright() as p:
            ensure_pw_browsers(log)

            # robustní launch: managed Chromium → Chrome → Edge
            log("Launching browser… headless =", headless)
            try:
                browser = p.chromium.launch(headless=headless)
            except Exception as e1:
                log("Chromium launch failed:", repr(e1), " – trying channel=chrome")
                try:
                    browser = p.chromium.launch(channel="chrome", headless=headless)
                except Exception as e2:
                    log("Chrome launch failed:", repr(e2), " – trying channel=msedge")
                    browser = p.chromium.launch(channel="msedge", headless=headless)

            context = browser.new_context()
            page = context.new_page()

            # 1) login
            log("Navigating to login…")
            page.goto("https://registr.ping-pong.cz/htm/auth/login.php",
                      wait_until="domcontentloaded")
            page.fill("input[name='login']", user_login)
            page.fill("input[name='heslo']",  user_pwd)
            page.locator("[name='send']").click()
            page.wait_for_load_state("domcontentloaded")
            log("Logged in.")

            # 2) stránka družstva podle ID
            url_team = f"https://registr.ping-pong.cz/htm/auth/klub/druzstva/vysledky/?druzstvo={team['id']}"
            log("Open team page:", url_team)
            page.goto(url_team, wait_until="domcontentloaded")

            # 3) vložit/upravit zápis
            log("Click 'vložit zápis' / 'upravit zápis'…")
            l = page.locator("a:has-text('vložit zápis')")
            if not l.count(): l = page.locator("a:has-text('upravit zápis')")
            l.first.click()
            page.wait_for_load_state("domcontentloaded")

            # 4) úvodní část – herna/začátek/vedoucí
            if team.get("herna") and page.locator("input[name='zapis_herna']").count():
                page.fill("input[name='zapis_herna']", str(team["herna"]))
            if team.get("zacatek"):
                if page.locator("input[name='zapis_zacatek']").count():
                    page.fill("input[name='zapis_zacatek']", team["zacatek"])
                else:
                    try:
                        hh, mm = team["zacatek"].split(":")
                        sels = page.locator("select")
                        if sels.count() >= 2:
                            sels.nth(0).select_option(value=hh)
                            sels.nth(1).select_option(value=mm)
                    except Exception as e:
                        log("Set start time via selects failed:", repr(e))

            if team.get("ved_dom") and page.locator("select[name='id_domaci_vedouci']").count():
                page.select_option("select[name='id_domaci_vedouci']", label=str(team["ved_dom"]))
            if team.get("ved_host") and page.locator("select[name='id_hoste_vedouci']").count():
                page.select_option("select[name='id_hoste_vedouci']", label=str(team["ved_host"]))

            # 5) Uložit a pokračovat → online formulář
            log("Click 'Uložit a pokračovat'…")
            if not click_save_and_continue(page):
                raise RuntimeError("Nenašel jsem tlačítko 'Uložit a pokračovat'.")
            page.wait_for_url(re.compile(r".*/online\.php\?u=\d+"), timeout=20000)
            page.wait_for_selector("input[type='text']", timeout=10000)
            log("Online formulář načten – připraveno k ručnímu dopsání výsledků.")

            if headed:
                log("Leaving browser open for manual finish.")
                print("✅ Online formulář načten – dokonči ručně. Okno nechávám otevřené.")
                while True:
                    time.sleep(1)
            else:
                context.close()
                browser.close()

    except Exception as e:
        # zapiš chybový stack a otevři log v Notepadu
        log("ERROR:", repr(e))
        log(traceback.format_exc())
        try:
            os.startfile(str(log_path))  # otevřít log v Notepadu (Windows)
        except Exception:
            pass
        raise
    finally:
        try:
            log("==== stis_uploader end ====")
            log_file.close()
        except Exception:
            pass

