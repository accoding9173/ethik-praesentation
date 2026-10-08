#!/usr/bin/env python3
"""Nimmt ein Preview-Video aller Folien auf (mit Animationen)."""
import os, time
from playwright.sync_api import sync_playwright

HIER = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HIER, 'preview.webm')

with sync_playwright() as p:
    browser = p.chromium.launch(executable_path='/usr/bin/google-chrome-stable')
    ctx = browser.new_context(
        viewport={'width': 1920, 'height': 1080},
        record_video_dir='/tmp/preview_vid',
        record_video_size={'width': 1920, 'height': 1080},
    )
    page = ctx.new_page()
    page.goto('file://' + os.path.join(HIER, 'index.html'))

    anzahl = page.locator('.slide').count()
    print(f'{anzahl} Folien gefunden, starte Aufnahme...')

    # Erste Folie etwas länger zeigen
    page.wait_for_timeout(2500)

    for i in range(1, anzahl):
        page.keyboard.press('ArrowRight')
        # Video-Folie (Folie 2) länger zeigen
        if i == 1:
            page.wait_for_timeout(5000)
        else:
            page.wait_for_timeout(2500)

    # Letzte Folie kurz stehen lassen
    page.wait_for_timeout(1500)

    # Video-Pfad holen bevor Context geschlossen wird
    video_path = page.video.path()
    ctx.close()
    browser.close()

    # Kopieren
    import shutil
    shutil.move(video_path, OUT)
    print(f'Video gespeichert: {OUT}')
