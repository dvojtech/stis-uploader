# stis_uploader.py
import argparse, os, re, sys, time
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


def norm(s: str) -> str:
    s = (s or "").strip().lower()
    repl = str.maketrans("áčďéěíňóřšťúůýž", "acdeeinorstuuyz")
    return re.sub(r"[^\w]", "", s.translate(repl))

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
    # najdi list 'setup' case-insensitive, jinak první list
    for ws in wb.worksheets:
        if (ws.title or "").strip().lower() == "setup":
            return ws
    return wb.active

def read_excel_config(xlsx_path: Path, team_name: str):
    wb = load_workbook(xlsx_path, data_only=True)
    setup = get_setup_sheet(wb)

    login = str(setup["B1"].value or "").strip()
    pwd   = str(setup["B2"].value or "")
    if not login or not pwd:
        raise RuntimeError("Vyplň login/heslo v setup!B1:B2.")

    # --- vyhledej řádek hlavičky kdekoli v prvních 30 řádcích ---
    def row_norm(r, max_c):
        return [norm(setup.cell(r, c).value or "") for c in range(1, max_c + 1)]

    max_r = min(30, setup.max_row)
    max_c = min(50, setup.max_column)
    hdr_row = None
    for r in range(1, max_r + 1):
        rn = row_norm(r, max_c)
        has_name = any("druzstvo" in v for v in rn)
        # akceptuj DruzstvoID i Druzstvoid (překlepy)
        has_id   = any(("druzstvoid" in v) or ("druzstvoid" in v) or v == "id" for v in rn)
        if has_name and has_id:
            hdr_row = r
            break
    if hdr_row is None:
        raise RuntimeError("V setup!Teams chybí sloupce 'Družstvo' a/nebo 'DruzstvoID'.")

    # --- namapuj sloupce podle nalezené hlavičky (synonyma povolena) ---
    idx = {}
    for c in range(1, max_c + 1):
        h = norm(setup.cell(hdr_row, c).value or "")
        if ("druzstvo" in h) and ("id" not in h) and ("vedouci" not in h):
            idx["name"] = c
        if ("druzstvoid" in h) or ("druzstvoid" in h) or h == "id" or "iddruzstva" in h:
            idx["id"] = c
        if "vedoucidomacich" in h or (("vedouci" in h) and ("host" not in h)):
            idx["ved_dom"] = c
        if "vedoucihostu" in h or (("vedouci" in h) and ("host" in h)):
            idx["ved_host"] = c
        if "herna" in h:
            idx["herna"] = c
        if "zacatekut" in h or "zacatek" in h:
            idx["zacatek"] = c
        if "konecutkani" in h or "konec" in h:
            idx["konec"] = c

    if "name" not in idx or "id" not in idx:
        raise RuntimeError("V setup!Teams chybí sloupce 'Družstvo' a/nebo 'DruzstvoID'.")

    # --- najdi řádek družstva pod hlavičkou ---
    team = None
    r = hdr_row + 1
    while r <= setup.max_row:
        nm = setup.cell(r, idx["name"]).value
        if nm is None or str(nm).strip() == "":
            break  # konec tabulky
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
        raise RuntimeError(f"Družstvo '{team_name}' nenalezeno v setup!Teams.")
    if not team["id"]:
        raise RuntimeError("Prázdné DruzstvoID.")

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
