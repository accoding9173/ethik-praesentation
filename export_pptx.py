#!/usr/bin/python3.14
"""Export der Ethik-Praesentation als .pptx"""

import sys, glob
for p in glob.glob('/home/js/.local/lib/python3.*/site-packages'):
    sys.path.insert(0, p)

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

Q = '„'  # opening german quote
Qc = '“'  # closing german quote
Qs = '‚'  # single low-9
Qsc = '‘'  # single closing

BG = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
GREY = RGBColor(170, 170, 170)
LIGHT = RGBColor(204, 204, 204)
DARK_GREY = RGBColor(136, 136, 136)

W = Inches(13.333)
H = Inches(7.5)


def set_bg(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BG


def add_text(slide, text, left, top, width, height, size=18, color=WHITE,
             bold=False, italic=False, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].alignment = align
    r = tf.paragraphs[0].add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = 'Helvetica Neue'
    return tb


def add_label(slide, text, top=Inches(1.2)):
    add_text(slide, text.upper(), Inches(1.5), top, Inches(10), Inches(0.5),
             size=14, color=DARK_GREY)


def add_title(slide, text, top=Inches(1.8)):
    add_text(slide, text, Inches(1.5), top, Inches(10.3), Inches(1),
             size=32, color=WHITE, bold=True)


def add_bullets(slide, items, top=Inches(3.0)):
    tb = slide.shapes.add_textbox(Inches(1.5), top, Inches(10.3), Inches(4))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(8)
        parts = item.split('**')
        for j, part in enumerate(parts):
            if not part:
                continue
            r = p.add_run()
            r.text = part
            r.font.size = Pt(18)
            r.font.name = 'Helvetica Neue'
            r.font.bold = (j % 2 == 1)
            r.font.color.rgb = WHITE if (j % 2 == 1) else LIGHT


def add_quote(slide, quote, author, top=Inches(5.8)):
    tb = slide.shapes.add_textbox(Inches(1.8), top, Inches(9.5), Inches(1.2))
    tf = tb.text_frame
    tf.word_wrap = True
    r = tf.paragraphs[0].add_run()
    r.text = Q + quote + Qc
    r.font.size = Pt(16)
    r.font.color.rgb = WHITE
    r.font.italic = True
    r.font.name = 'Helvetica Neue'
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = author
    r2.font.size = Pt(11)
    r2.font.color.rgb = DARK_GREY
    r2.font.name = 'Helvetica Neue'


def add_source(slide, text):
    add_text(slide, text, Inches(0.5), Inches(6.8), Inches(12), Inches(0.5),
             size=9, color=RGBColor(102, 102, 102), align=PP_ALIGN.CENTER)


# ======= Build =======

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
blank = prs.slide_layouts[6]

# -- 1: Titel --
s = prs.slides.add_slide(blank); set_bg(s)
add_text(s, 'Vortrag über ein ethisches Thema', Inches(1), Inches(2.5),
         Inches(11), Inches(1.5), size=48, bold=True, align=PP_ALIGN.CENTER)
add_text(s, 'RELIGION', Inches(1), Inches(4.3), Inches(11), Inches(0.6),
         size=22, color=GREY, align=PP_ALIGN.CENTER)

# -- 2: Spaghetti-Video --
s = prs.slides.add_slide(blank); set_bg(s)
add_text(s, '\U0001F3AC  Video: Will Smith isst Spaghetti (KI-generiert)\n\n'
         'Links: 2023  —  Rechts: 2025\n\n→ spaghetti.mp4 abspielen',
         Inches(1.5), Inches(2), Inches(10), Inches(3.5), size=28, color=GREY, align=PP_ALIGN.CENTER)
add_source(s, 'Quelle: ' + Q + 'Comparison AI of Will Smith eating spaghetti' + Qc + ', r/GenAI4all')

# -- 3: Artikel-Screenshot --
s = prs.slides.add_slide(blank); set_bg(s)
img = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'artikel.png')
if os.path.exists(img):
    s.shapes.add_picture(img, Inches(1.5), Inches(0.8), Inches(10.3), Inches(5.5))
else:
    add_text(s, '\U0001F4F0  artikel.png hier einfügen', Inches(1.5), Inches(2.5),
             Inches(10), Inches(2), size=28, color=GREY, align=PP_ALIGN.CENTER)
add_source(s, 'Quelle: skill-sprinters.de, 13.09.2026')

# -- 4: Aktuelle Relevanz --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Aktuelle Relevanz')
add_title(s, Q + 'We Must Pace the Frontier' + Qc)
add_bullets(s, [
    '**12.09.2026:** Anthropic-Chef Dario Amodei fordert, die KI-Entwicklung zu bremsen',
    'Kein Stopp, aber die Fähigkeiten der stärksten Modelle sollen **langsamer** wachsen',
    'Plan: **externe Prüfer** in den Firmen, **Gesetze** für alle KI-Firmen, **weltweite Absprachen** (auch mit China)',
    'OpenAI, xAI (Musk) und Microsoft stimmen zu',
    '**Aber:** Nur 10 Tage später bringen Anthropic und OpenAI neue Top-Modelle heraus',
])

# -- 5: Grundfrage / Waage --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Die Grundfrage')
add_text(s, 'Sollte die Entwicklung immer stärkerer KI gebremst werden?',
         Inches(1.5), Inches(1.7), Inches(10.3), Inches(1.2), size=34, bold=True)
add_text(s, 'Vollgas', Inches(1.5), Inches(3.3), Inches(4.5), Inches(0.5), size=22, bold=True)
add_text(s, '• Fortschritt (Medizin, Forschung)\n• Wohlstand und Wettbewerb\n• ' + Q + 'Sonst baut es China' + Qc,
         Inches(1.5), Inches(3.9), Inches(4.5), Inches(2), size=18, color=LIGHT)
add_text(s, 'Bremsen', Inches(7.3), Inches(3.3), Inches(4.5), Inches(0.5), size=22, bold=True)
add_text(s, '• Sicherheit und Kontrolle\n• Verantwortung für die Zukunft\n• Folgen sind nicht umkehrbar',
         Inches(7.3), Inches(3.9), Inches(5), Inches(2), size=18, color=LIGHT)
add_text(s, 'Problem: Wenige Firmen entscheiden über ein Risiko, das alle betrifft.',
         Inches(1.5), Inches(6.0), Inches(10.3), Inches(0.6), size=18, color=RGBColor(221,221,221))

# -- 6: Konkrete Risiken --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Warum überhaupt bremsen?')
add_title(s, 'Was auf dem Spiel steht')
add_bullets(s, [
    '**Deepfakes & Desinformation** — KI-generierte Videos/Stimmen kaum noch erkennbar',
    '**Autonome Waffen** — Maschinen, die ohne Menschen über Leben entscheiden',
    '**Kontrollverlust** — Systeme, die klüger sind als ihre Entwickler',
    '**Jobs** — Viele Berufe werden sich verschieben, aber manche werden einfach ersetzt',
    '**Werte:** Das sind keine Theorien, das passiert bereits jetzt',
])

# -- 7: Position 1: Wirtschaft --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Position 1: Wirtschaft')
add_title(s, Q + 'Nicht bremsen — das kann sich niemand leisten' + Qc)
add_bullets(s, [
    'Hunderte Milliarden Dollar stecken in KI — basierend auf der Erwartung, dass es **immer schneller** weitergeht',
    'Bremsen sendet ein Signal: **' + Q + 'Das Wachstum verlangsamt sich' + Qc + '** → Investoren ziehen sich zurück',
    'Ohne Risikokapital **platzt die Blase** → Startups sterben, Bewertungen kollabieren',
    'Folge: eine Wirtschaftskrise wie beim **Dot-com-Crash 2000**',
    '**Werte:** Wohlstand, Stabilität, Wachstum',
])
add_quote(s, 'Die KI-Branche hat ein System gebaut, das Vollgas braucht, um nicht zu crashen.',
          'Eigene Formulierung')

# -- 8: Geldkreislauf-Dreieck --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Der Geldkreislauf')
add_title(s, 'Warum niemand freiwillig bremst')
add_text(s, 'NVIDIA — Jensen Huang\n(Chips / GPUs)', Inches(5), Inches(3.0),
         Inches(3.3), Inches(1.2), size=16, color=RGBColor(118,185,0), bold=True, align=PP_ALIGN.CENTER)
add_text(s, '↙', Inches(3.8), Inches(4.2), Inches(1), Inches(0.6), size=28, color=GREY, align=PP_ALIGN.CENTER)
add_text(s, '↘', Inches(8.5), Inches(4.2), Inches(1), Inches(0.6), size=28, color=GREY, align=PP_ALIGN.CENTER)
add_text(s, 'Oracle — Larry Ellison\n(Rechenzentren)', Inches(1.5), Inches(5.2),
         Inches(3.3), Inches(1.2), size=16, color=RGBColor(248,0,0), bold=True, align=PP_ALIGN.CENTER)
add_text(s, '→→→→→', Inches(5), Inches(5.5), Inches(3.3), Inches(0.5),
         size=16, color=GREY, align=PP_ALIGN.CENTER)
add_text(s, 'OpenAI — Sam Altman\n(KI-Modelle)', Inches(8.5), Inches(5.2),
         Inches(3.3), Inches(1.2), size=16, bold=True, align=PP_ALIGN.CENTER)
add_text(s, '↑', Inches(2.8), Inches(4.5), Inches(0.5), Inches(0.5), size=20, color=GREY, align=PP_ALIGN.CENTER)

# -- 9: Der Widerspruch --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Der Widerspruch')
add_title(s, 'Bremsen fordern, Vollgas geben')
add_bullets(s, [
    '**12. September 2026:** Anthropic-Chef Amodei veröffentlicht ' + Q + 'We Must Pace the Frontier' + Qc,
    'OpenAI, Microsoft und xAI stimmen öffentlich zu',
    '**22. September 2026:** Anthropic bringt ein neues Top-Modell heraus',
    '**23. September 2026:** OpenAI zieht nach',
    'Beide Firmen planen gleichzeitig ihren **Börsengang** — Bremsen passt nicht zur Wachstumsstory',
])
add_quote(s, 'Wenn wir von ' + Qs + 'Pacing' + Qsc + ' sprechen, meinen wir nicht ' + Qs + 'Stoppen' + Qsc + '.',
          'Sam Altman, CEO von OpenAI (übersetzt)')

# -- 10: Position 2: Europa --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Position 2: Europa')
add_title(s, Q + 'Erst aufholen, dann bremsen' + Qc)
add_bullets(s, [
    'Europa hat **keine echten Frontier-Modelle**, nur Firmen wie Mistral (FR) und Black Forest Labs (DE)',
    'Ziel: erst einmal **gute eigene Modelle** bauen',
    'Eine Bremse jetzt würde den **Vorsprung von USA und China festschreiben**',
    '**Werte:** Gerechtigkeit, Chancengleichheit, Unabhängigkeit',
])
add_quote(s, 'Einige etablierte Firmen nutzen diesen Moment, um ihre Marktposition zu festigen.',
          'Mistral AI (übersetzt)', top=Inches(5.2))
add_text(s, 'KI-Risikokapital:  USA ████████████████████ 77 %   Europa ███ 11 %',
         Inches(1.8), Inches(6.3), Inches(10), Inches(0.4), size=12, color=GREY)

# -- 11: Position 3: USA --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Position 3: USA')
add_title(s, Q + 'Keine Gesetze, das ist nur Taktik' + Qc)
add_bullets(s, [
    '**Trump** lehnt Regulierung ab und greift Amodei öffentlich an',
    'KI-Berater **David Sacks:** Die Firmen sollen selbst bremsen, statt nach Gesetzen zu rufen',
    'Vorwurf **' + Q + 'Regulatory Capture' + Qc + '**: Große Firmen wollen Regeln, die nur sie erfüllen können',
    'China würde sowieso nicht mitmachen',
    '**Werte:** Freiheit, Wettbewerb, nationale Stärke',
])
add_quote(s, 'Hört auf so zu tun, als wäre der Wunsch zu bremsen rein selbstlos.',
          'David Sacks, KI-Berater des Weißen Hauses (übersetzt)')

# -- 12: Position 4: China --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Position 4: China')
add_title(s, Q + 'Wir bremsen nicht – schon gar nicht auf Zuruf' + Qc)
add_bullets(s, [
    'China drosselt **nicht** freiwillig – KI gilt als strategische, nationale Priorität',
    'Sieht die US-Rufe nach ' + Q + 'Pacing' + Qc + ' als Versuch, den **eigenen Vorsprung einzufrieren**',
    'Schlägt aber selbst eine **' + Q + 'Global AI Governance Initiative' + Qc + '** vor: Zusammenarbeit – aber zu eigenen Bedingungen',
    'Das widerlegt teils das US-Argument ' + Q + 'China macht eh nicht mit' + Qc,
    '**Werte:** nationale Stärke, Souveränität, technologische Unabhängigkeit',
])

# -- 13: Unsere Meinung --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Unsere Meinung')
add_title(s, Q + 'Bremsen an der Spitze, nicht bei den Aufholern' + Qc)
add_bullets(s, [
    '**Werte:** Sicherheit und Unabhängigkeit',
    '**Norm:** Die stärksten KI-Modelle dürfen erst veröffentlicht werden, wenn **unabhängige, internationale Prüfer** sie getestet haben',
    'Diese Regel gilt für **alle** Spitzenmodelle, egal aus welchem Land',
    'Wer aufholt (z. B. Europa), wird dadurch **nicht** ausgebremst',
])

# -- 14: Urteilsfindung --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Wie wir zu unserem Urteil kamen')
add_title(s, 'Sicherheit und Unabhängigkeit abwägen')
add_bullets(s, [
    'Das Spaghetti-Video zeigt: KI wird **extrem schnell** besser, deshalb sind wir eher fürs Bremsen',
    'Aber: Eine Bremse für alle würde den **Vorsprung von USA und China festschreiben**',
    'Wenige Firmen würden dann allein entscheiden, Europa wäre abhängig',
    'Unsere Lösung: **Nur das Tempo an der Spitze kontrollieren**, damit beide Werte geschützt sind',
    'Gegenargument USA: ' + Q + 'China macht nicht mit' + Qc + ', deshalb müssen die Prüfer **international** sein',
])

# -- 15: Diskussion --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Diskussion')
add_text(s, 'Würdet ihr die Entwicklung\nimmer stärkerer KI bremsen?',
         Inches(1), Inches(2.5), Inches(11), Inches(2), size=38, bold=True, align=PP_ALIGN.CENTER)

# -- 16: Schluss --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Schluss')
add_text(s, 'Bremsen an der Spitze,\nnicht bei den Aufholern.',
         Inches(1), Inches(2.2), Inches(11), Inches(2), size=38, bold=True, align=PP_ALIGN.CENTER)
add_text(s, 'DANKE FÜRS ZUHÖREN', Inches(1), Inches(4.8), Inches(11), Inches(0.6),
         size=22, color=GREY, align=PP_ALIGN.CENTER)

# -- 17: Quellen --
s = prs.slides.add_slide(blank); set_bg(s)
add_label(s, 'Quellen')
quellen = [
    'Reddit, r/GenAI4all: ' + Q + 'Comparison AI of Will Smith eating spaghetti' + Qc,
    'skill-sprinters.de: ' + Q + 'Anthropic-Chef Amodei will das Tempo der KI-Entwicklung drosseln …' + Qc + ', 13.09.2026',
    'Dario Amodei: ' + Q + 'We Must Pace the Frontier' + Qc + ', 12.09.2026',
    'The Register: ' + Q + 'Frontier AI keeps racing despite calls to slow down' + Qc + ', 23.09.2026',
    'MarkTechPost: ' + Q + 'Anthropic’s 3-Step Pace the Frontier Plan …' + Qc + ', 13.09.2026',
    'CNBC: ' + Q + 'Trump rejects AI regulation calls, slams Anthropic CEO …' + Qc + ', 14.09.2026',
    'Dealroom: ' + Q + 'David Sacks to Altman and Amodei: pace the frontier yourselves' + Qc,
    'Reuters: ' + Q + 'Europe’s AI firms, playing catch-up, challenge US calls for slowdown' + Qc,
    'Cryptopolitan: ' + Q + 'Europe’s AI firms warn a frontier slowdown would lock in the US lead' + Qc,
]
tb = s.shapes.add_textbox(Inches(1.5), Inches(2.0), Inches(10.3), Inches(5))
tf = tb.text_frame
tf.word_wrap = True
for i, q in enumerate(quellen):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    r = p.add_run()
    r.text = str(i + 1) + '. ' + q
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(187, 187, 187)
    r.font.name = 'Helvetica Neue'
    p.space_after = Pt(4)

# === Save ===
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ethik-praesentation.pptx')
prs.save(out)
print('Gespeichert: ' + out)
