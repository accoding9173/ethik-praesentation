# Macht Screenshots von jeder Folie der index.html + von den Stichpunkten.
import pathlib
from playwright.sync_api import sync_playwright

base = pathlib.Path(__file__).parent
out = base / "screenshots"
out.mkdir(exist_ok=True)

with sync_playwright() as p:
    b = p.chromium.launch(executable_path="/usr/bin/google-chrome-stable")

    # --- Praesentation: jede Folie einzeln ---
    pg = b.new_page(viewport={"width": 1280, "height": 720}, device_scale_factor=2, locale="de-DE")
    pg.goto((base / "index.html").as_uri(), wait_until="networkidle")
    n = pg.evaluate("document.querySelectorAll('.slide').length")
    print(f"{n} Folien gefunden")
    for i in range(n):
        pg.evaluate("i => show(i)", i)
        pg.wait_for_timeout(900)  # Opacity-Transition + ggf. Video-Frame
        path = out / f"folie-{i+1:02d}.png"
        pg.screenshot(path=str(path))
        print("->", path.name)

    # --- Stichpunkte (A4-Dokument, ganze Seite) ---
    pg2 = b.new_page(viewport={"width": 1000, "height": 1400}, device_scale_factor=2, locale="de-DE")
    pg2.goto((base / "stichpunkte.html").as_uri(), wait_until="networkidle")
    pg2.screenshot(path=str(out / "stichpunkte.png"), full_page=True)
    print("-> stichpunkte.png")

    b.close()
print("Fertig.")
