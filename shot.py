# Macht einen Screenshot vom Nachrichtenartikel (nutzt das installierte Chrome)
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    b = p.chromium.launch(executable_path="/usr/bin/google-chrome-stable")
    pg = b.new_page(viewport={"width": 1280, "height": 900}, device_scale_factor=2, locale="de-DE")
    pg.goto("https://skill-sprinters.de/blog/news/amodei-tempo-ki-entwicklung-externe-pruefer-2026/", wait_until="networkidle")
    pg.screenshot(path="artikel.png", clip={"x": 240, "y": 170, "width": 800, "height": 270})
    b.close()
