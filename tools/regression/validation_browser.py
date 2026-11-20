"""Regrese životního cyklu investičních widgetů ve skutečném Chromiu.

Vývojový nástroj (Playwright), bez změny dat. Spusťte testovací Streamlit server
v režimu jediného testovacího hesla a:
  DE_TEST_PASSWORD=… python tools/regression/validation_browser.py URL VYSTUP
Heslo patří pouze do prostředí. DE_CHROMIUM volitelně určuje binárku prohlížeče,
DE_WIDTHS má výchozí hodnotu 1440,1024,390,320. Ověří i DOCX proti původní referenci.
"""
import hashlib
import json
import os
from pathlib import Path
import shutil
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
import navigace

EXPECTED = ['Kolový bagr', 'Silniční válec', 'Optimistický', 60000, 25, 2000, 5000000]
KEYS = ['inv_machine_1', 'inv_machine_2', 'inv_scenario', 'inv_fixed_cost',
        'inv_expected_days', 'inv_n', 'inv_budget']


def idle(page):
    page.wait_for_timeout(1300)
    page.wait_for_function("document.querySelector('[data-testid=stApp]')?.getAttribute('data-test-script-state') !== 'running'", timeout=120000)
    assert not page.get_by_test_id('stException').count()


def values(page):
    return [page.locator(f'.st-key-{k} [role=combobox]').input_value() for k in KEYS[:3]] + [
        int(page.locator(f'.st-key-{k} input').input_value()) for k in KEYS[3:]]


def select(page, key, value):
    field = page.locator(f'.st-key-{key} [role=combobox]')
    field.evaluate("e=>e.scrollIntoView({block:'center',behavior:'instant'})")
    page.wait_for_timeout(250)
    field.click()
    page.get_by_role('option', name=value, exact=True).click()
    idle(page)


def number(page, key, value):
    field = page.locator(f'.st-key-{key} input')
    field.fill(str(value)); field.press('Enter'); idle(page)


def slider(page, key, target):
    field = page.locator(f'.st-key-{key} input[type=range]')
    if int(field.input_value()) == target:
        return
    field.evaluate("e=>e.parentElement.parentElement.scrollIntoView({block:'center',behavior:'instant'})")
    page.wait_for_timeout(250)
    track = field.evaluate('e=>e.parentElement.parentElement.parentElement.getBoundingClientRect().toJSON()')
    thumb = field.evaluate('e=>e.parentElement.parentElement.getBoundingClientRect().toJSON()')
    lo, hi = float(field.get_attribute('min')), float(field.get_attribute('max'))
    page.mouse.click(track['x']+(target-lo)/(hi-lo)*track['width'], thumb['y']+thumb['height']/2)
    idle(page)
    for _ in range(2):
        diff = target-int(field.input_value())
        assert abs(diff) < 40, (key, target, field.input_value())
        for _ in range(abs(diff)):
            field.press('ArrowRight' if diff > 0 else 'ArrowLeft')
        idle(page)
        if int(field.input_value()) == target:
            return
    raise AssertionError(f'Unable to set slider {key}')


def view(page, slug):
    expand = page.get_by_test_id('stExpandSidebarButton')
    if expand.count() and expand.is_visible():
        expand.click(); page.wait_for_timeout(400)
    side = page.get_by_test_id('stSidebar')
    details = side.get_by_test_id('stExpander').locator('details')
    if details.get_attribute('open') is None:
        details.locator('summary').click(); page.wait_for_timeout(400)
    side.locator(f'.st-key-nav_inv_{slug} button:visible').click()
    idle(page)
    close = page.get_by_test_id('stSidebarCollapseButton')
    if page.viewport_size['width'] < 500 and close.is_visible():
        close.click(); page.wait_for_timeout(400)


def run(url, output):
    from playwright.sync_api import sync_playwright, expect
    password = os.environ.get('DE_TEST_PASSWORD')
    if not password:
        raise SystemExit('Set DE_TEST_PASSWORD using a test account only.')
    output.mkdir(parents=True, exist_ok=True)
    widths = [int(v) for v in os.environ.get('DE_WIDTHS', '1440,1024,390,320').split(',')]
    reference = json.loads((ROOT/'tests/fixtures/investment_navigation_reference.json').read_text())['optimisticky']['docx']
    with sync_playwright() as pw:
        executable = os.environ.get('DE_CHROMIUM') or shutil.which('chromium')
        browser = pw.chromium.launch(**({'executable_path': executable} if executable else {}), args=['--no-sandbox'])
        try:
            for width in widths:
                page = browser.new_page(viewport={'width': width, 'height': 1200}, accept_downloads=True)
                page.goto(url.rstrip('/')+'/investice?pohled=doporuceni')
                page.locator('input[type=password]').fill(password); page.keyboard.press('Enter')
                page.locator('.st-key-de_inv_inputs').wait_for(timeout=120000); idle(page)
                close = page.get_by_test_id('stSidebarCollapseButton')
                if width < 500 and close.is_visible(): close.click()
                for key, value in zip(KEYS[:3], EXPECTED[:3]): select(page, key, value)
                number(page, 'inv_budget', EXPECTED[6])
                details = page.locator('.st-key-de_inv_advanced details')
                if details.get_attribute('open') is None: details.locator('summary').click()
                number(page, 'inv_fixed_cost', EXPECTED[3])
                slider(page, 'inv_expected_days', EXPECTED[4]); slider(page, 'inv_n', EXPECTED[5])
                assert values(page) == EXPECTED
                page.locator('.st-key-de_inv_inputs').screenshot(path=str(output/f'{width}-valid.png'))
                if width in (1024, 390):
                    details.locator('summary').click(); idle(page)
                select(page, 'inv_machine_2', EXPECTED[0])
                assert all(page.locator(f'.st-key-{k}').count() == 1 for k in KEYS), 'All seven widgets must remain mounted during validation'
                assert values(page) == [EXPECTED[0], EXPECTED[0], *EXPECTED[2:]], values(page)
                expect(page.get_by_test_id('stAlert')).to_contain_text('různé')
                assert not page.locator('.de-hero').count()
                assert not page.get_by_test_id('stPlotlyChart').count()
                assert not page.get_by_role('button', name='Stáhnout report (DOCX)', exact=True).count()
                page.locator('.st-key-de_inv_inputs').screenshot(path=str(output/f'{width}-invalid.png'))
                # I změna pohledu během validace musí zachovat všechny vstupy.
                view(page, 'riziko')
                assert values(page) == [EXPECTED[0], EXPECTED[0], *EXPECTED[2:]]
                select(page, 'inv_machine_2', EXPECTED[1])
                assert values(page) == EXPECTED, values(page)
                page.locator('.st-key-de_inv_inputs').screenshot(path=str(output/f'{width}-repaired.png'))
                for slug in navigace.INVESTICE_VIEWS:
                    view(page, slug)
                    assert values(page) == EXPECTED, (slug, values(page))
                    expect(page.locator('.de-inv-summary')).to_contain_text('optimistický')
                view(page, 'doporuceni')
                with page.expect_download() as info:
                    page.get_by_role('button', name='Stáhnout report (DOCX)', exact=True).click()
                report = output/f'{width}-report.docx'; info.value.save_as(report)
                with zipfile.ZipFile(report) as archive:
                    got = {n: hashlib.sha256(archive.read(n)).hexdigest() for n in sorted(archive.namelist())}
                assert got == reference, 'DOCX changed after validation/repair'
                idle(page); assert values(page) == EXPECTED
                print(f'PASS {width}: seven inputs, invalid/repair, all views, exact DOCX', flush=True)
                page.close()
        finally:
            browser.close()


if __name__ == '__main__':
    if len(sys.argv) != 3:
        raise SystemExit('Usage: validation_browser.py URL OUTPUT_DIRECTORY')
    run(sys.argv[1], Path(sys.argv[2]))
