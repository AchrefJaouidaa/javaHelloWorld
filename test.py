from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
import os

OUT = os.path.join(os.getcwd(), "presentations", "Data_Product_Generative_AI_Strategy_FR.pptx")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

BLUE = RGBColor(12, 52, 90)
LIGHT = RGBColor(238, 244, 250)
ACCENT = RGBColor(58, 130, 246)
GREEN = RGBColor(26, 128, 74)
ORANGE = RGBColor(227, 120, 56)
RED = RGBColor(187, 52, 52)
TEXT = RGBColor(35, 35, 35)
WHITE = RGBColor(255, 255, 255)
GRAY = RGBColor(90, 90, 90)


def set_bg(slide, color=WHITE):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_title(slide, title, subtitle=None, title_color=BLUE):
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.0), Inches(0.7))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.alignment = PP_ALIGN.LEFT
    p.runs[0].font.size = Pt(24)
    p.runs[0].font.bold = True
    p.runs[0].font.color.rgb = title_color
    if subtitle:
        sub = slide.shapes.add_textbox(Inches(0.7), Inches(1.0), Inches(11.0), Inches(0.45))
        stf = sub.text_frame
        sp = stf.paragraphs[0]
        sp.text = subtitle
        sp.alignment = PP_ALIGN.LEFT
        sp.runs[0].font.size = Pt(11)
        sp.runs[0].font.color.rgb = GRAY


def add_bullets(slide, x, y, w, h, bullets, font_size=18, color=TEXT):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.05)
    for i, item in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.bullet = True
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(8)
        p.runs[0].font.size = Pt(font_size)
        p.runs[0].font.color.rgb = color
    return box


def add_box(slide, left, top, width, height, text, fill_color=LIGHT, line_color=BLUE, font_size=18, title=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.color.rgb = line_color
    shape.line.width = Pt(1.2)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.12)
    tf.margin_right = Inches(0.12)
    tf.margin_top = Inches(0.08)
    if title:
        p = tf.paragraphs[0]
        p.text = title
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = BLUE
        p = tf.add_paragraph()
        p.text = text
        p.font.size = Pt(font_size)
        p.font.color.rgb = TEXT
    else:
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(font_size)
        p.font.color.rgb = TEXT
    return shape


def add_chart(slide, left, top, width, height, chart_type, data, title=None):
    chart_data = CategoryChartData()
    chart_data.categories = data['categories']
    for series_name, vals in data['series'].items():
        chart_data.add_series(series_name, vals)
    chart = slide.shapes.add_chart(chart_type, Inches(left), Inches(top), Inches(width), Inches(height), chart_data).chart
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
    return chart


def add_footer(slide):
    footer = slide.shapes.add_textbox(Inches(0.5), Inches(7.1), Inches(12.0), Inches(0.2))
    p = footer.text_frame.paragraphs[0]
    p.text = "Data Product Engineering & AI Automation | Stratégie d’industrialisation"
    p.alignment = PP_ALIGN.RIGHT
    p.runs[0].font.size = Pt(8)
    p.runs[0].font.color.rgb = GRAY


# Slide 1
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.25))
bar.fill.solid(); bar.fill.fore_color.rgb = BLUE; bar.line.fill.background()
logo = slide.shapes.add_textbox(Inches(0.7), Inches(0.4), Inches(2.8), Inches(0.7))
p = logo.text_frame.paragraphs[0]
p.text = "AI Data Product Factory"
p.runs[0].font.size = Pt(18); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = WHITE
main = slide.shapes.add_textbox(Inches(0.7), Inches(1.3), Inches(9.0), Inches(1.8))
tf = main.text_frame
for i, line in enumerate(["Usine de génération", "de data products", "avec IA et gouvernance"]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.text = line
    p.alignment = PP_ALIGN.LEFT
    p.runs[0].font.size = Pt(26 if i == 1 else 18)
    p.runs[0].font.bold = i == 1
    p.runs[0].font.color.rgb = BLUE
sub = slide.shapes.add_textbox(Inches(0.8), Inches(3.1), Inches(8.5), Inches(1.2))
subtf = sub.text_frame
for i, line in enumerate([
    "Présentation orientée managers + équipes de développement",
    "Une architecture qui convertit un besoin métier en spécification, puis en artefacts techniques, sans perdre la traçabilité ni la gouvernance.",
    "Scénarios : interface métier, agents IA, skills, Git, CI/CD et multi-cibles de génération."
]):
    p = subtf.paragraphs[0] if i == 0 else subtf.add_paragraph()
    p.text = line
    p.font.size = Pt(16 if i == 0 else 14)
    p.font.bold = i == 0
    p.font.color.rgb = TEXT
for i, (label, val, color) in enumerate([
    ("Productivité", "+50%", ACCENT),
    ("Qualité", "+35%", GREEN),
    ("Temps de spec", "-60%", BLUE),
    ("Réutilisation", "+40%", ORANGE),
]):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.3 + i * 1.0), Inches(2.6), Inches(0.9), Inches(1.8))
    box.fill.solid(); box.fill.fore_color.rgb = LIGHT; box.line.color.rgb = color
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = val
    p.alignment = PP_ALIGN.CENTER
    p.runs[0].font.size = Pt(18); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = color
    p = tf.add_paragraph()
    p.text = label
    p.alignment = PP_ALIGN.CENTER
    p.runs[0].font.size = Pt(8); p.runs[0].font.color.rgb = GRAY
add_footer(slide)

# Slide 2
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "1. Contexte : pourquoi ce sujet prend de l’ampleur", "Le besoin n’est plus seulement d’écrire du SQL ou de décrire une table, mais de livrer un produit de données réutilisable et gouverné")
add_bullets(slide, Inches(0.8), Inches(1.45), Inches(5.8), Inches(4.7), [
    "Les organisations manipulent des volumes de données plus importants, des catalogues plus complexes et des demandes métier plus rapides à transformer.",
    "Les data products doivent exprimer plus que des schémas : ils portent un contrat de qualité, une propriété, une documentation et un mode d’exposition.",
    "La création de nouveaux artefacts est encore trop dépendante de parcours manuels, d’Excel, de tickets et de fragments de connaissance dispersés.",
    "L’IA doit accélérer la transformation sans diluer la qualité, la sécurité ni la responsabilité métier.",
    "Le vrai enjeu n’est pas seulement la génération, c’est la standardisation du contrat produit."
])
add_box(slide, Inches(7.1), Inches(1.8), Inches(4.9), Inches(2.0), "Problème courant\n\n• spécifications éclatées\n• ownership flou\n• scripts redondants\n• revue lente\n• déploiement non reproductible", fill_color=LIGHT, line_color=RED, font_size=16, title="Pain points")
add_box(slide, Inches(7.1), Inches(4.3), Inches(4.9), Inches(1.8), "État ciblé\n\nUne seule source de vérité, des règles de validation, des générateurs modulaires et une revue humaine explicite.", fill_color=LIGHT, line_color=GREEN, font_size=16, title="Vision")
add_footer(slide)

# Slide 3
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "2. Le point de rupture : la chaîne actuelle est trop longue", "Le problème est moins technique que méthodologique")
add_bullets(slide, Inches(0.8), Inches(1.5), Inches(5.7), Inches(4.5), [
    "Le besoin métier est souvent exprimé dans un document, un mail, un ticket, ou un fichier Excel avec un niveau d’ambiguïté élevé.",
    "La traduction vers la couche technique est faite manuellement par plusieurs intervenants : analyste, data engineer, architecte, développeur.",
    "A chaque étape, l’intention s’érode, les hypothèses se multiplient et les écarts de qualité augmentent.",
    "La mise en production dépend d’une série d’étapes de validation peu automatisées et difficilement traçables.",
    "Résultat : retard, erreurs de compréhension, redondance d’effort et faible capitalisation."
])
add_box(slide, Inches(7.1), Inches(1.8), Inches(4.8), Inches(2.1), "Besoin métier\n→ spéc docs\n→ discussions manuelles\n→ schéma technique\n→ SQL / jobs / tests\n→ revue\n→ prod", fill_color=LIGHT, line_color=ORANGE, font_size=16, title="Chaîne actuelle")
add_box(slide, Inches(7.1), Inches(4.3), Inches(4.8), Inches(1.6), "Conséquence\n\nLe temps de valeur augmente, la qualité dépend du héros du projet et la gouvernance reste a posteriori.", fill_color=LIGHT, line_color=RED, font_size=16)
add_footer(slide)

# Slide 4
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "3. L’idée centrale : un produit de données doit être décrit avant d’être généré", "La spécification est la source de vérité, pas le code HTML ou les fragments dispersés")
blocks = [
    (Inches(0.8), Inches(2.1), Inches(2.0), Inches(1.4), "Besoin métier\n+ contraintes", BLUE),
    (Inches(3.1), Inches(2.1), Inches(2.0), Inches(1.4), "Interface\nformulaires / YAML", ACCENT),
    (Inches(5.8), Inches(2.1), Inches(1.9), Inches(1.4), "Agent IA\nvalidation & plan", GREEN),
    (Inches(8.1), Inches(2.1), Inches(2.0), Inches(1.4), "Skills\n+ générateurs", ORANGE),
    (Inches(10.7), Inches(2.1), Inches(1.8), Inches(1.4), "Livraison\ncode + docs + tests", RED),
]
for left, top, width, height, text, color in blocks:
    b = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    b.fill.solid(); b.fill.fore_color.rgb = LIGHT; b.line.color.rgb = color; b.line.width = Pt(1.5)
    tf = b.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text; p.alignment = PP_ALIGN.CENTER; p.runs[0].font.size = Pt(14); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = color
for x1, y1, x2, y2 in [
    (Inches(2.8), Inches(2.8), Inches(3.1), Inches(2.8)),
    (Inches(5.1), Inches(2.8), Inches(5.8), Inches(2.8)),
    (Inches(7.7), Inches(2.8), Inches(8.1), Inches(2.8)),
    (Inches(10.1), Inches(2.8), Inches(10.7), Inches(2.8)),
]:
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    line.line.color.rgb = BLUE; line.line.width = Pt(2)
add_box(slide, Inches(1.1), Inches(4.3), Inches(3.8), Inches(1.6), "Entrées utiles\n• business rules\n• ownership\n• règles d’accès\n• qualité & contrat", fill_color=LIGHT, line_color=BLUE, font_size=15, title="Source de vérité")
add_box(slide, Inches(5.4), Inches(4.3), Inches(3.2), Inches(1.6), "Machine / IA\n• validation\n• complétion\n• génération\n• revues guidées", fill_color=LIGHT, line_color=GREEN, font_size=15, title="Orchestration")
add_box(slide, Inches(9.3), Inches(4.3), Inches(2.8), Inches(1.6), "Sorties\n• code\n• SQL / API\n• docs\n• tests\n• catalog", fill_color=LIGHT, line_color=ORANGE, font_size=15, title="Livrables")
add_footer(slide)

# Slide 5
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "4. Interfaces métiers et techniques : l’architecture doit servir les deux mondes", "Le système doit être utilisable par un métier guidé et par un développeur avancé")
add_box(slide, Inches(0.9), Inches(1.8), Inches(5.2), Inches(3.9), "Interface métier / product owner\n\n• formulaire de description du produit\n• définition du périmètre, des flux et des règles\n• ownership, classification, publication\n• validation par règles de gouvernance\n• génération du YAML / contrat source\n\nValeur : démocratiser la création sans exposer l’implémentation technique.", fill_color=LIGHT, line_color=BLUE, font_size=15, title="Interface 1 : parcours guidé")
add_box(slide, Inches(6.5), Inches(1.8), Inches(5.9), Inches(3.9), "Interface développeur / Git / IDE\n\n• prompts, commandes ou workflows de génération\n• orchestration via skills et agents\n• génération SQL / dbt / API / tests\n• intégration CI/CD et validation\n• revue, correction et merge dans le dépôt\n\nValeur : accélérer la production avec un contrôle explicite et reproductible.", fill_color=LIGHT, line_color=ACCENT, font_size=15, title="Interface 2 : développeur + skills")
add_footer(slide)

# Slide 6
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "5. Architecture technique générique", "Une solution modulaire, ouverte et capable de servir plusieurs cibles")
layers = [
    (Inches(0.7), Inches(1.8), Inches(2.1), Inches(1.1), "Interface\nUI / portail / IDE"),
    (Inches(3.2), Inches(1.8), Inches(2.1), Inches(1.1), "Spécification\nYAML / contrat"),
    (Inches(5.8), Inches(1.8), Inches(2.1), Inches(1.1), "Agent IA\nplanification"),
    (Inches(8.4), Inches(1.8), Inches(2.1), Inches(1.1), "Skills\n+ générateurs"),
    (Inches(10.9), Inches(1.8), Inches(1.8), Inches(1.1), "Runtime\ncible"),
]
for left, top, w, h, label in layers:
    b = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, w, h)
    b.fill.solid(); b.fill.fore_color.rgb = LIGHT; b.line.color.rgb = BLUE; b.line.width = Pt(1.5)
    tf = b.text_frame; p = tf.paragraphs[0]; p.text = label; p.alignment = PP_ALIGN.CENTER; p.runs[0].font.size = Pt(12); p.runs[0].font.color.rgb = BLUE
for i in range(4):
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(2.8 + i*2.55), Inches(2.35), Inches(3.2 + i*2.55), Inches(2.35))
    line.line.color.rgb = BLUE; line.line.width = Pt(2)
add_box(slide, Inches(0.9), Inches(3.8), Inches(5.6), Inches(2.4), "Principes d’architecture\n• métadonnées comme source de vérité\n• validation avant livraison\n• gouvernance intégrée au modèle\n• générateurs spécialisés et réutilisables\n• compatibilité multi-plateformes et multi-environnements", fill_color=LIGHT, line_color=GREEN, font_size=16, title="Contrat technique")
add_box(slide, Inches(7.1), Inches(3.8), Inches(5.3), Inches(2.4), "Exemples de cibles\n• SQL / dbt / transformations\n• API / services\n• vues sémantiques\n• docs, glossaires, règles de qualité\n• CI/CD et pipelines d’orchestration", fill_color=LIGHT, line_color=ORANGE, font_size=15, title="Sorties")
add_footer(slide)

# Slide 6 bis
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "5. Architecture technique détaillée du système d’orchestration IA", "Les agents ne sont pas des fonctions isolées : ils forment une chaîne de contrôle, de génération et de validation")

boxes = [
    (Inches(0.6), Inches(1.8), Inches(1.8), Inches(1.2), "1. Intake\nUI / portal / IDE", BLUE),
    (Inches(2.7), Inches(1.8), Inches(1.9), Inches(1.2), "2. Spec Agent\nnormalisation\nYAML", ACCENT),
    (Inches(5.0), Inches(1.8), Inches(1.9), Inches(1.2), "3. Score Agent\nqualité + complétude", GREEN),
    (Inches(7.4), Inches(1.8), Inches(2.0), Inches(1.2), "4. Validator Agent\ncontrats, schéma, règles", ORANGE),
    (Inches(9.9), Inches(1.8), Inches(1.8), Inches(1.2), "5. Generator Agent\nSQL / dbt / API", RED),
    (Inches(11.9), Inches(1.8), Inches(1.0), Inches(1.2), "6. Review Agent", BLUE),
    (Inches(1.6), Inches(3.8), Inches(2.3), Inches(1.5), "7. Compliance Agent\nownership / sécurité / publication", BLUE),
    (Inches(4.4), Inches(3.8), Inches(2.1), Inches(1.5), "8. Test Agent\nvalidation / lineage / qualité", GREEN),
    (Inches(7.0), Inches(3.8), Inches(2.1), Inches(1.5), "9. Diff Agent\nreview / patch / corrigendum", ORANGE),
    (Inches(9.6), Inches(3.8), Inches(2.2), Inches(1.5), "10. Human Approval\nvalidateur final", RED),
    (Inches(6.4), Inches(5.8), Inches(2.8), Inches(1.3), "11. CI/CD\nmerge, build, deploy, traceability", BLUE),
]
for left, top, w, h, text, color in boxes:
    b = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, w, h)
    b.fill.solid(); b.fill.fore_color.rgb = LIGHT; b.line.color.rgb = color; b.line.width = Pt(1.4)
    tf = b.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text; p.alignment = PP_ALIGN.CENTER; p.runs[0].font.size = Pt(11); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = color

for x1, y1, x2, y2 in [
    (Inches(2.4), Inches(2.4), Inches(2.7), Inches(2.4)),
    (Inches(4.6), Inches(2.4), Inches(5.0), Inches(2.4)),
    (Inches(6.9), Inches(2.4), Inches(7.4), Inches(2.4)),
    (Inches(9.4), Inches(2.4), Inches(9.9), Inches(2.4)),
    (Inches(11.7), Inches(2.4), Inches(11.9), Inches(2.4)),
    (Inches(2.3), Inches(3.2), Inches(2.3), Inches(3.8)),
    (Inches(3.9), Inches(3.2), Inches(4.4), Inches(3.8)),
    (Inches(6.5), Inches(3.2), Inches(7.0), Inches(3.8)),
    (Inches(9.2), Inches(3.2), Inches(9.6), Inches(3.8)),
    (Inches(7.5), Inches(5.3), Inches(7.5), Inches(5.8)),
    (Inches(7.8), Inches(6.1), Inches(8.9), Inches(6.1)),
]:
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    line.line.color.rgb = BLUE; line.line.width = Pt(2)

add_box(slide, Inches(0.7), Inches(6.0), Inches(11.6), Inches(0.9), "Rôle des agents\nSpec Agent = transforme le besoin en contrat technique ; Score Agent = détecte les lacunes ; Validator = vérifie la conformité ; Generator = produit les artefacts ; Review = critique, corrige et sécurise ; Human Approval = décision finale.", fill_color=LIGHT, line_color=GREEN, font_size=14)
add_footer(slide)

# Slide 7
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "6. Les agents IA nécessaires au cycle de production", "Chaque agent joue une fonction claire dans la chaîne de valeur du data product")
add_bullets(slide, Inches(0.8), Inches(1.5), Inches(5.8), Inches(4.8), [
    "Spec Agent : normalise les inputs métier, crée ou corrige le YAML de référence et aide à structurer le produit, les entités, les sources et les contraintes.",
    "Score Agent : évalue la complétude, la cohérence, la qualité métier et la maturité du produit. Il identifie les « gaps » avant génération.",
    "Validator Agent : contrôle les règles de schéma, les dépendances, les obligations gouvernance, les droits d’accès et la cohérence des objets de données.",
    "Generator Agent : produit le code technique cible : SQL, dbt, modèles, API, jobs, tests, métriques, scripts et documentation.",
    "Review Agent : lit le résultat généré, critique la logique, repère les anomalies et propose des correctifs de qualité ou de sécurité.",
    "Compliance / Security Agent : vérifie l’ownership, la classification, les règles éditoriales, les accès, et les restrictions de publication.",
    "Diff Agent : compare la version générée avec la version précédente, identifie les écarts et vérifie qu’il n’y a pas de rupture non intentionnelle.",
    "Human Approval Agent / Reviewer : ne produit pas le code lui-même, mais valide les écarts critiques et autorise la publication ou le merge."
])
add_box(slide, Inches(7.0), Inches(2.1), Inches(4.7), Inches(1.2), "Validation\ncontrat + qualité + schéma + access", fill_color=LIGHT, line_color=GREEN, font_size=16)
add_box(slide, Inches(7.0), Inches(3.6), Inches(4.7), Inches(1.2), "Scoring\ncomplétude + maturité + risques", fill_color=LIGHT, line_color=ACCENT, font_size=16)
add_box(slide, Inches(7.0), Inches(5.1), Inches(4.7), Inches(1.2), "Review\nqualité + sécurité + revue humaine", fill_color=LIGHT, line_color=ORANGE, font_size=16)
add_footer(slide)

# Slide 8
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "7. Les skills et agents travaillent en chaîne, pas en silo", "La valeur vient du couplage entre fonctions, validation et contrôle humain")
add_bullets(slide, Inches(0.8), Inches(1.5), Inches(5.5), Inches(4.4), [
    "Le flux ne doit pas être un simple prompt unique. Il doit être un pipeline de décision : capture → classification → score → validation → génération → revue → approbation.",
    "Chaque agent a un rôle et une responsabilité explicite. Il ne remplace pas le métier, il alimente la qualité du système en réduisant les erreurs de compréhension et d’implémentation.",
    "Les skills encapsulent les générateurs tandis que les agents orchestrent le séquençage des tâches, les conditions de validation et les points d’arrêt métiers.",
    "Dans un environnement mature, cette architecture permet d’industrialiser la production de data products sans dépendre d’un seul individu ou d’un seul outil."
])
add_box(slide, Inches(7.0), Inches(2.0), Inches(4.7), Inches(1.2), "Capture\nspécification métier", fill_color=LIGHT, line_color=BLUE, font_size=16)
add_box(slide, Inches(7.0), Inches(3.5), Inches(4.7), Inches(1.2), "Validation\nscore + conformité", fill_color=LIGHT, line_color=GREEN, font_size=16)
add_box(slide, Inches(7.0), Inches(5.0), Inches(4.7), Inches(1.2), "Approbation\nrevue + merge + publication", fill_color=LIGHT, line_color=RED, font_size=16)
for pos in [2.6, 4.1, 5.6]:
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(7.1), Inches(pos + 0.4), Inches(11.7), Inches(pos + 0.4))
    line.line.color.rgb = BLUE; line.line.width = Pt(2)
add_footer(slide)

# Renumbering subsequent slides from 8 onward to keep order coherent
# This keeps the original content while adding the detailed architecture before the usage scenarios.

# Slide 9
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "8. Gains attendus et KPI de pilotage", "Le bénéfice doit être exprimé en valeur opérationnelle, pas seulement en quantité de code généré")
chart_data = CategoryChartData()
chart_data.categories = ["Temps de spec", "Révision manuelle", "Qualité logicielle", "Réutilisation", "Mise en prod"]
chart_data.add_series("Approche actuelle", [75, 70, 55, 40, 80])
chart_data.add_series("Flux assisté par IA", [30, 35, 85, 78, 40])
chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.8), Inches(1.6), Inches(8.2), Inches(4.2), chart_data).chart
chart.chart_title.text_frame.text = "Indice d’effort relatif (plus faible est mieux)"
chart.has_legend = True
chart.legend.include_in_layout = False
add_box(slide, Inches(9.5), Inches(1.8), Inches(2.7), Inches(1.1), "-60%\nTemps de spec", fill_color=LIGHT, line_color=ACCENT, font_size=16)
add_box(slide, Inches(9.5), Inches(3.2), Inches(2.7), Inches(1.1), "+35%\nQualité", fill_color=LIGHT, line_color=GREEN, font_size=16)
add_box(slide, Inches(9.5), Inches(4.6), Inches(2.7), Inches(1.0), "+40%\nRéutilisation", fill_color=LIGHT, line_color=ORANGE, font_size=16)
add_footer(slide)

# Slide 10
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "9. Premier backlog priorisé", "Le bon ordre d’attaque est le facteur de succès le plus important")
add_bullets(slide, Inches(0.8), Inches(1.5), Inches(5.7), Inches(4.8), [
    "P0 – Définir le modèle métier des métadonnées et les règles de validation minimales (ownership, classification, qualité, publication).",
    "P0 – Construire un éditeur minimal pour capturer l’intention métier et générer le YAML de référence, sans dépendre d’outils complexes.",
    "P0 – Créer un workflow de génération IA piloté par un agent / skill pour une cible technique unique (ex. SQL / dbt / structure métier).",
    "P1 – Ajouter des contrôles sémantiques, des tests automatiques, des checks de conformité et des garde-fous de revue avant merge ou production.",
    "P1 – Rendre le système générique pour plusieurs cibles : SQL, API, docs, glossaires, vues sémantiques, configuration d’orchestration.",
    "P1 – Intégrer les bouches de publication dans la CI/CD : validation, approbation, publication et traçabilité d’artefacts.",
    "P2 – Mesurer le ROI sur plusieurs équipes et domaines : temps, coûts, qualité, réutilisation, incident rate, satisfaction utilisateurs.",
    "P2 – Développer les templates et la maturité d’urbanisation : catalogues, standards de design, communautés d’usage, documentation de référence."
])
add_box(slide, Inches(7.2), Inches(2.0), Inches(4.7), Inches(2.0), "Ce qu’on construit d’abord\n\nUne source de vérité robuste, un flow de validation, un générateur fiable, puis une industrialisation progressive.", fill_color=LIGHT, line_color=BLUE, font_size=16)
add_footer(slide)

# Slide 11
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "10. Risques, défis et gouvernance", "L’IA accélère la génération, mais la qualité dépend du cadre de contrôle")
add_bullets(slide, Inches(0.8), Inches(1.5), Inches(5.7), Inches(4.4), [
    "L’IA peut produire des artefacts plausibles mais incorrects si le modèle de métadonnées est incomplet ou ambigu.",
    "La gouvernance doit être intégrée dans la spécification elle-même, et non seulement au moment du déploiement.",
    "Les droits d’accès, les niveaux d’exposition et les approbations doivent être explicites et testables.",
    "Les contrôles de qualité doivent être présents tout au long du workflow, pas seulement en fin de pipeline.",
    "La vraie valeur économique vient du couplage entre capitalisation de connaissances, génération assistée et revue humaine structurée."
])
add_box(slide, Inches(7.4), Inches(2.0), Inches(4.1), Inches(1.8), "Contrôles à inclure\n• ownership\n• classification\n• qualité\n• conformité\n• approbation\n• traçabilité", fill_color=LIGHT, line_color=RED, font_size=16, title="Gouvernance")
add_box(slide, Inches(7.4), Inches(4.3), Inches(4.1), Inches(1.5), "Succès attendu\n\nL’IA accélère la livraison sans supprimer la responsabilité, ni la confiance métier.", fill_color=LIGHT, line_color=GREEN, font_size=16)
add_footer(slide)

# Slide 12
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "11. Questions ouvertes pour la discussion stratégique", "Ce qui doit être tranché avant industrialisation")
add_bullets(slide, Inches(0.8), Inches(1.5), Inches(6.0), Inches(4.6), [
    "Quels domaines sont les meilleurs premiers candidats pour un modèle commun de data product ?",
    "Quelle part de gouvernance doit être intégrée dans l’interface et quelle part dans la CI/CD ?",
    "Quel niveau de revue humaine est adapté selon la criticité du flux et du générateur ?",
    "Comment gérer les sorties spécifiques à une plateforme sans casser le modèle commun de définition ?",
    "Comment mesurer la valeur : réduction de cycle, qualité, réduction de coût, ou délai de mise sur le marché ?",
    "Comment construire la confiance dans l’organisation sans ralentir les équipes de produit ?"
])
add_footer(slide)

# Slide 13
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "12. Conclusion", "Le vrai gain vient de l’industrialisation du produit de données, pas d’un simple prompt")
summary = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(3.3))
summary_tf = summary.text_frame
for i, text in enumerate([
    "La transformation ne consiste pas à remplacer l’ingénierie par l’IA, mais à standardiser le contrat produit et à automatiser la traduction vers la cible technique.",
    "Un bon système associe spécification, agent, skills, générateurs, validation de qualité et gouvernance explicite.",
    "La valeur vient de la reproductibilité, de la traçabilité, et de la capacité à scaler un modèle de production sur plusieurs équipes et plusieurs cibles."
]):
    p = summary_tf.paragraphs[0] if i == 0 else summary_tf.add_paragraph()
    p.text = text
    p.alignment = PP_ALIGN.LEFT
    p.runs[0].font.size = Pt(22)
    p.runs[0].font.color.rgb = BLUE
    p.runs[0].font.bold = True
add_footer(slide)

# Slide 8 bis
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "8. Exemple d’interface de saisie : remplir un data product", "Une interface simple qui structure le besoin de manière exploitable par l’IA et la gouvernance")

# Form mockup
form = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.7), Inches(6.4), Inches(4.7))
form.fill.solid(); form.fill.fore_color.rgb = LIGHT; form.line.color.rgb = BLUE; form.line.width = Pt(1.5)
form_tf = form.text_frame; form_tf.word_wrap = True
for i, line in enumerate([
    "Produit de données",
    "Nom : Contrats_client",
    "Owner : Product Owner",
    "Description : Synthèse des contrats et des montants",
    "Classification : Interne / Partagée / Publique",
    "Sources : CRM, ERP, BI",
    "Règles qualité : non null, audit trail, publication contrôlée",
    "Domaine métier : relation client / finance",
    "Accès : équipe CRM, finance, data",
    "Statut : draft / validated / ready for generation"
]):
    p = form_tf.paragraphs[0] if i == 0 else form_tf.add_paragraph()
    p.text = line
    p.font.size = Pt(15 if i == 0 else 13)
    p.font.bold = i == 0
    p.font.color.rgb = BLUE if i == 0 else TEXT

# Right side panel
panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.6), Inches(1.8), Inches(4.8), Inches(4.3))
panel.fill.solid(); panel.fill.fore_color.rgb = WHITE; panel.line.color.rgb = GREEN; panel.line.width = Pt(1.5)
panel_tf = panel.text_frame
for i, line in enumerate([
    "Auto-completion",
    "- owner détecté",
    "- domaines métiers suggérés",
    "- relations de sources proposées",
    "Validation",
    "- schéma complet",
    "- droits d’accès cohérents",
    "- classification conforme",
    "Score de maturité",
    "82/100",
    "Prêt pour génération : Oui"
]):
    p = panel_tf.paragraphs[0] if i == 0 else panel_tf.add_paragraph()
    p.text = line
    p.font.size = Pt(15 if i in [0, 3, 6] else 13)
    p.font.bold = i in [0, 3, 6]
    p.font.color.rgb = GREEN if i in [0, 3, 6] else TEXT

add_footer(slide)

# Slide 8 bis 2
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "9. Exemple d’interface de génération : lancer et superviser la production", "La génération doit être pilotable, traçable et contrôlée par l’utilisateur")

left = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(5.8), Inches(4.2))
left.fill.solid(); left.fill.fore_color.rgb = LIGHT; left.line.color.rgb = ACCENT; left.line.width = Pt(1.5)
ltf = left.text_frame
for i, line in enumerate([
    "Commande de génération",
    "Type de génération : dbt / API / vue sémantique / docs",
    "Cible : Snowflake / BigQuery / Postgres / generic",
    "Mode : draft / validation / production",
    "Agents activés : Spec Agent, Validator, Generator, Review",
    "Boutons : Prévisualiser / Valider / Générer / Approuver",
    "Observabilité : logs, diff, score, santé des dépendances"
]):
    p = ltf.paragraphs[0] if i == 0 else ltf.add_paragraph()
    p.text = line
    p.font.size = Pt(15 if i == 0 else 13)
    p.font.bold = i == 0
    p.font.color.rgb = ACCENT if i == 0 else TEXT

right = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.0), Inches(1.9), Inches(5.4), Inches(4.0))
right.fill.solid(); right.fill.fore_color.rgb = WHITE; right.line.color.rgb = ORANGE; right.line.width = Pt(1.5)
rtf = right.text_frame
for i, line in enumerate([
    "Sortie attendue",
    "- modèles staging",
    "- tests de qualité",
    "- docs de publication",
    "- vue métier / catalog",
    "- rapport de validation",
    "- diff de changements",
    "Dernier statut : validation réussie + approbation humaine en attente"
]):
    p = rtf.paragraphs[0] if i == 0 else rtf.add_paragraph()
    p.text = line
    p.font.size = Pt(15 if i == 0 else 13)
    p.font.bold = i == 0
    p.font.color.rgb = ORANGE if i == 0 else TEXT

add_footer(slide)

# Slide 8 bis 3 : architecture technique for devs
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
add_title(slide, "10. Architecture technique pour les développeurs", "De l’interface jusqu’au runtime cible, avec orchestration et validation explicites")

layers = [
    (Inches(0.3), Inches(2.0), Inches(1.5), Inches(1.0), "UI / Portal\nStreamlit / Web"),
    (Inches(1.9), Inches(2.0), Inches(1.7), Inches(1.0), "App Service\nAPI / workflow"),
    (Inches(3.8), Inches(2.0), Inches(1.8), Inches(1.0), "Orchestrator\nagent planner"),
    (Inches(5.8), Inches(2.0), Inches(1.8), Inches(1.0), "Spec Store\nYAML / metadata"),
    (Inches(7.8), Inches(2.0), Inches(1.8), Inches(1.0), "Skills\nDBT / SQL / API"),
    (Inches(9.8), Inches(2.0), Inches(1.6), Inches(1.0), "Agents\nvalidator / review"),
    (Inches(11.7), Inches(2.0), Inches(1.3), Inches(1.0), "CI/CD"),
    (Inches(0.9), Inches(3.8), Inches(2.2), Inches(1.2), "Git Repo\nsource + PRs"),
    (Inches(3.5), Inches(3.8), Inches(2.4), Inches(1.2), "Generator Runtime\npython / templates"),
    (Inches(6.3), Inches(3.8), Inches(2.2), Inches(1.2), "Quality Gate\nunit tests / diff"),
    (Inches(9.0), Inches(3.8), Inches(2.2), Inches(1.2), "Runtime Target\nSnowflake / SQL / API"),
    (Inches(11.6), Inches(3.8), Inches(1.4), Inches(1.2), "Docs & Catalog"),
]
for left, top, w, h, text in layers:
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, w, h)
    box.fill.solid(); box.fill.fore_color.rgb = LIGHT; box.line.color.rgb = BLUE; box.line.width = Pt(1.4)
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text; p.alignment = PP_ALIGN.CENTER; p.runs[0].font.size = Pt(11); p.runs[0].font.bold = True; p.runs[0].font.color.rgb = BLUE

for x1, y1, x2, y2 in [
    (Inches(1.8), Inches(2.5), Inches(1.9), Inches(2.5)),
    (Inches(3.6), Inches(2.5), Inches(3.8), Inches(2.5)),
    (Inches(5.6), Inches(2.5), Inches(5.8), Inches(2.5)),
    (Inches(7.6), Inches(2.5), Inches(7.8), Inches(2.5)),
    (Inches(9.6), Inches(2.5), Inches(9.8), Inches(2.5)),
    (Inches(11.3), Inches(2.5), Inches(11.7), Inches(2.5)),
    (Inches(1.9), Inches(3.1), Inches(1.9), Inches(3.8)),
    (Inches(4.3), Inches(3.1), Inches(4.3), Inches(3.8)),
    (Inches(7.0), Inches(3.1), Inches(7.0), Inches(3.8)),
    (Inches(9.6), Inches(3.1), Inches(9.6), Inches(3.8)),
    (Inches(11.8), Inches(3.1), Inches(11.8), Inches(3.8)),
    (Inches(2.1), Inches(4.9), Inches(3.5), Inches(4.9)),
    (Inches(4.9), Inches(4.9), Inches(6.3), Inches(4.9)),
    (Inches(8.5), Inches(4.9), Inches(9.0), Inches(4.9)),
]:
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    line.line.color.rgb = BLUE; line.line.width = Pt(2)

add_box(slide, Inches(0.6), Inches(5.5), Inches(12.0), Inches(0.9), "Flux technique de production\nBesoin métier → spec YAML → orchestrateur → skills & agents → validation → génération → test → PR → CI/CD → publication dans le runtime cible.", fill_color=LIGHT, line_color=GREEN, font_size=14)
add_footer(slide)

