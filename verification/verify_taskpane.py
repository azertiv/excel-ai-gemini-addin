
from playwright.sync_api import sync_playwright
import os

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        # Load the taskpane HTML file directly
        # We need the absolute path
        cwd = os.getcwd()
        filepath = f'file://{cwd}/src/taskpane/taskpane.html'

        print(f'Navigating to {filepath}')
        page.goto(filepath)

        # Check if the Base URL input exists and is visible
        base_url_input = page.locator('#baseUrlInput')
        if base_url_input.is_visible():
            print('Base URL input is visible')
            # Type something into it
            base_url_input.fill('https://api.gptbots.ai/v1')
        else:
            print('Base URL input NOT visible')

        # Take a screenshot
        page.screenshot(path='verification/taskpane_screenshot.png')
        browser.close()

if __name__ == '__main__':
    run()
