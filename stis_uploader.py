# stis_uploader.py
import argparse, os, re, sys, time
import unicodedata
from pathlib import Path
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

def ensure_pw_browsers():
    # když běží z EXE (onefile), data jsou v sys._MEIPASS
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).parent
    browsers = base / "ms-playwright"
    if browsers.is_dir():
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers)

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

def main():
    ensure_pw_browsers()
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    ap.add_argument("--team", required=True)
    ap.add_argument("--headed", action="store_true")  # okno viditelné
    args = ap.parse_args()

    login, pwd, team = read_excel_config(Path(args.excel), args.team)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.headed)
        ctx = browser.new_context()
        page = ctx.new_page()

        # 1) login
        page.goto("https://registr.ping-pong.cz/htm/auth/login.php", wait_until="domcontentloaded")
        page.fill("input[name=login]", login)
        page.fill("input[name=heslo]", pwd)
        page.locator("[name=send]").click()
        page.wait_for_load_state("networkidle")

        # 2) družstvo
        page.goto(f"https://registr.ping-pong.cz/htm/auth/klub/druzstva/vysledky/?druzstvo={team['id']}",
                  wait_until="domcontentloaded")

        # 3) vložit/upravit zápis
        l = page.locator("a:has-text('vložit zápis')")
        if not l.count(): l = page.locator("a:has-text('upravit zápis')")
        l.first.click()
        page.wait_for_load_state("domcontentloaded")

        # 4) úvodní část
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
                except: pass

        if team.get("ved_dom") and page.locator("select[name='id_domaci_vedouci']").count():
            page.select_option("select[name='id_domaci_vedouci']", label=str(team["ved_dom"]))
        if team.get("ved_host") and page.locator("select[name='id_hoste_vedouci']").count():
            page.select_option("select[name='id_hoste_vedouci']", label=str(team["ved_host"]))

        # 5) Uložit a pokračovat -> čekej na online formulář
        if not click_save_and_continue(page):
            raise RuntimeError("Nenašel jsem tlačítko 'Uložit a pokračovat'.")
        page.wait_for_url(re.compile(r".*/online\.php\?u=\d+"), timeout=20000)
        page.wait_for_selector("input[type='text']", timeout=10000)

        # nech okno otevřené pro vizuální dokončení
        if args.headed:
            print("✅ Online formulář načten – dokonči ručně. Okno nechávám otevřené.")
            while True: time.sleep(1)
        else:
            browser.close()

if __name__ == "__main__":
    main()
