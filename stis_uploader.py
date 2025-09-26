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

EXE_DIR = Path(sys.argv[0]).resolve().parent
TEMP_DIR = Path(os.environ.get("TEMP", str(EXE_DIR)))
BOOT_FILES = [
    TEMP_DIR / "stis_boot.log",
    EXE_DIR / "stis_boot.log",
]

ZDROJ_SHEET = "zdroj"
ZDROJ_FIRST_SINGLE_ROW = 7
SINGLES_COUNT = 16

def wait_online_ready(page, log, timeout=25000):
    try:
        page.wait_for_selector("#zapis .event", state="visible", timeout=timeout)
        page.wait_for_selector(".zapas-set", state="visible", timeout=timeout)
        log("Online editor ready.")
    except Exception as e:
        log(f"Online editor se nenačetl: {repr(e)}")
        raise

def _fill_player_by_click(page, selector, name, log):
    name = (name or "").strip()
    if not name:
        return
    try:
        player_elem = page.locator(selector)
        if not player_elem.count():
            log(f" Nenalezen element: {selector}")
            return
        player_elem.click(timeout=3000)
        page.wait_for_timeout(300)
        for sel in [
            "input.ui-autocomplete-input:focus",
            "input.ac_input:focus",
            "input[type='text']:focus:not(.zapas-set):not([disabled])"
        ]:
            inp = page.locator(sel)
            if inp.count():
                inp.fill(name)
                page.keyboard.press("Tab")
                page.wait_for_timeout(200)
                log(f" ✓ {name} → {selector}")
                break
        else:
            page.keyboard.type(name)
            page.keyboard.press("Tab")
            log(f" ~ {name} → {selector} (fallback)")
    except Exception as e:
        log(f" ✗ {name} → {selector} failed: {repr(e)}")

def _fill_sets_by_event_index(page, event_index, sets, log):
    if not sets:
        return
    try:
        event = page.locator(".event").nth(event_index)
        if not event.count():
            log(f" Event #{event_index} nenalezen")
            return
        for i, value in enumerate(sets[:5]):
            if not value:
                continue
            set_input = event.locator(f".zapas-set[data-set='{i+1}']")
            if set_input.count():
                set_input.fill(str(_map_wo(value)))
                log(f" set{i+1} ← {value} (event #{event_index})")
            else:
                log(f" Set {i+1} input nenalezen pro event #{event_index}")
    except Exception as e:
        log(f" Sety pro event #{event_index} selhaly: {repr(e)}")

def _map_wo(val):
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
        png_path = Path(xlsx_path).with_suffix(".online_dump.png")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(page.content())
        page.screenshot(path=str(png_path), full_page=True)
        log(f"DOM dump → {html_path.name}, screenshot → {png_path.name}")
    except Exception as e:
        log(f"DOM dump failed: {repr(e)}")

def cnt(page, css):
    try:
        return page.locator(css).count()
    except Exception:
        return -1

def read_zdroj_data(xlsx_path):
    wb = load_workbook(xlsx_path, data_only=True)
    if ZDROJ_SHEET not in wb.sheetnames:
        raise RuntimeError(f"V sešitu chybí list '{ZDROJ_SHEET}'")
    sh = wb[ZDROJ_SHEET]
    def cell(r, c):
        return str(sh.cell(r, c).value or "").strip()
    double = {
        "home1": cell(2, 4),
        "home2": cell(3, 4),
        "away1": cell(2, 5),
        "away2": cell(3, 5),
        "sets": [ _map_wo(cell(3, 9+i)) for i in range(5) ]
    }
    singles = []
    row = ZDROJ_FIRST_SINGLE_ROW
    for idx in range(2, 2 + SINGLES_COUNT):
        home = cell(row, 4)
        away = cell(row, 5)
        sets = [_map_wo(cell(row, 9+i)) for i in range(5)]
        singles.append({"idx": idx, "home": home, "away": away, "sets": sets})
        row += 1
    return {"double": double, "singles": singles}

def fill_online_from_zdroj(page, data, log, xlsx_path=None):
    try:
        wait_online_ready(page, log)
    except Exception:
        log("Inputs se neobjevily – dělám dump DOMu.")
        if xlsx_path:
            _dom_dump(page, xlsx_path, log)
        raise
    dbl = data.get("double", {})
    if dbl.get("home1"):
        _fill_player_by_click(page, "#c0 .cell-player:first-child .player.domaci .player-name", dbl["home1"], log)
    if dbl.get("home2"):
        _fill_player_by_click(page, "#c0 .cell-player:last-child .player.domaci .player-name", dbl["home2"], log)
    if dbl.get("away1"):
        _fill_player_by_click(page, "#c0 .cell-player:first-child .player.host .player-name", dbl["away1"], log)
    if dbl.get("away2"):
        _fill_player_by_click(page, "#c0 .cell-player:last-child .player.host .player-name", dbl["away2"], log)
    if dbl.get("sets"):
        _fill_sets_by_event_index(page, 0, dbl["sets"], log)
    for match_data in data.get("singles", []):
        excel_idx = int(match_data.get("idx", 0))
        if excel_idx < 2 or excel_idx > 17: continue
        dom_idx = excel_idx - 2
        event_idx = excel_idx
        if match_data.get("home"):
            _fill_player_by_click(page, f"#d{dom_idx} .player.domaci .player-name", match_data["home"], log)
        if match_data.get("away"):
            _fill_player_by_click(page, f"#d{dom_idx} .player.host .player-name", match_data["away"], log)
        if match_data.get("sets"):
            _fill_sets_by_event_index(page, event_idx, match_data["sets"], log)
    log("Klikám 'Uložit změny'…")
    try:
        page.locator("input[name='ulozit']").click(timeout=5000)
        page.wait_for_timeout(1000)
        log("Změny uloženy.")
    except Exception as e:
        log(f"Uložení selhalo: {repr(e)}")

# ... (Další pomocné funkce: boot, msgbox, logger, norm, as_time_txt, get_setup_sheet, find_login_pwd, find_teams_header_anywhere, open_match_form...)

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
        user_login, user_pwd, team = read_excel_config(xlsx_path, args.team)
        log("Login OK; team:", team["name"], "ID:", team["id"])
        headed = bool(getattr(args, "headed", True))
        headless = not headed
        with sync_playwright() as p:
            ensure_pw_browsers(log)
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
            log("Navigating to login…")
            page.goto("https://registr.ping-pong.cz/htm/auth/login.php", wait_until="domcontentloaded")
            page.fill("input[name='login']", user_login)
            page.fill("input[name='heslo']", user_pwd)
            page.locator("[name='send']").click()
            page.wait_for_load_state("domcontentloaded")
            log("Logged in.")
            team_url = f"https://registr.ping-pong.cz/htm/auth/klub/druzstva/vysledky/?druzstvo={team['id']}"
            log("Open team page:", team_url)
            page.goto(team_url, wait_until="domcontentloaded")
            log("Hledám odkaz 'vložit/upravit zápis'…")
            if not open_match_form(page, log):
                raise RuntimeError("Na stránce družstva jsem nenašel odkaz do formuláře.")
            # 6) Vyplň úvodní údaje včetně začátku zápasu
            if team.get("herna") and page.locator("input[name='zapis_herna']").count():
                page.fill("input[name='zapis_herna']", str(team["herna"]))
                log("Herna vyplněna:", team["herna"])
            if team.get("zacatek"):
                try:
                    hh, mm = team["zacatek"].split(":")
                    if page.locator("select[name='zapis_zacatek_hodiny']").count():
                        page.select_option("select[name='zapis_zacatek_hodiny']", value=str(int(hh)))
                        log(f"Hodina nastavena: {hh}")
                    if page.locator("select[name='zapis_zacatek_minuty']").count():
                        page.select_option("select[name='zapis_zacatek_minuty']", value=str(int(mm)))
                        log(f"Minuta nastavena: {mm}")
                    log("Začátek nastaven:", team["zacatek"])
                except Exception as e:
                    log("Set start time failed:", repr(e))
            if team.get("ved_dom"):
                try:
                    ved_input = page.locator("input[name='id_domaci_vedoucitext']")
                    if ved_input.count():
                        ved_input.fill(str(team["ved_dom"]))
                        page.keyboard.press("Tab")
                        page.wait_for_timeout(500)
                        log("Vedoucí domácích:", team["ved_dom"])
                except Exception as e:
                    log("Vedoucí domácích selhal:", repr(e))
            if team.get("ved_host"):
                try:
                    ved_input = page.locator("input[name='id_hoste_vedoucitext']")
                    if ved_input.count():
                        ved_input.fill(str(team["ved_host"]))
                        page.keyboard.press("Tab")
                        page.wait_for_timeout(500)
                        log("Vedoucí hostů:", team["ved_host"])
                except Exception as e:
                    log("Vedoucí hostů selhal:", repr(e))
            # 7) Uložit a pokračovat
            log("Click 'Uložit a pokračovat'…")
            clicked = False
            try:
                btn = page.locator("input[name='odeslat']")
                if btn.count():
                    btn.click(timeout=5000)
                    clicked = True
                    log("Kliknuto na 'Uložit a pokračovat' (name=odeslat)")
            except Exception as e:
                log("Button(name=odeslat) click failed:", repr(e))
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
            try:
                page.wait_for_url(re.compile(r"/online\\.php\\?u=\\d+"), timeout=20000)
            except Exception:
                page.wait_for_selector("input[type='text']", timeout=10000)
                log("Online formulář načten:", page.url)
            data = read_zdroj_data(xlsx_path)
            fill_online_from_zdroj(page, data, log, xlsx_path)
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
                os.startfile(str(log_path))
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
        msgbox("Spuštění skončilo hned na začátku (špatné/neúplné argumenty?).\nZkontroluj prosím volání z Excelu.\nV TEMP nebo vedle EXE je stis_boot.log s detaily.")
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
