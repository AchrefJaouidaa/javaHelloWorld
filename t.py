from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

OUT = r"C:\Users\A708004\OneDrive - GROUP DIGITAL WORKPLACE\Desktop\SSM\snowflake\PPO_CPRO\uvp_explo\Framework_Data_Quality_Framework_5_Slides.pptx"
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

NAVY = RGBColor(16, 32, 72)
BLUE = RGBColor(36, 99, 180)
CYAN = RGBColor(88, 165, 214)
GREEN = RGBColor(33, 128, 78)
ORANGE = RGBColor(230, 126, 34)
RED = RGBColor(192, 57, 43)
LIGHT = RGBColor(245, 247, 250)
TEXT = RGBColor(35, 41, 53)
WHITE = RGBColor(255, 255, 255)
GRAY = RGBColor(111, 118, 127)


def section_title(slide, title):
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(12.0), Inches(0.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = NAVY
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.8), Inches(12.2), Inches(0.04))
    line.fill.solid(); line.fill.fore_color.rgb = BLUE; line.line.fill.background()


def add_bullets(slide, x, y, w, h, bullets, font_size=14, color=TEXT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = bullet
        p.level = 0
        p.bullet = True
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.space_after = Pt(8)


def add_card(slide, x, y, w, h, title, body, color, text_color=WHITE):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid(); box.fill.fore_color.rgb = color; box.line.fill.background()
    tb = box.text_frame; tb.word_wrap = True
    p = tb.paragraphs[0]
    p.text = title
    p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = text_color; p.space_after = Pt(8)
    p2 = tb.add_paragraph()
    p2.text = body
    p2.font.size = Pt(10); p2.font.color.rgb = text_color


def add_code_block(slide, x, y, w, h, lines, font_size=10, color=TEXT, fill_color=LIGHT):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid(); box.fill.fore_color.rgb = fill_color; box.line.color.rgb = GRAY
    tb = box.text_frame; tb.word_wrap = True
    for idx, line in enumerate(lines):
        p = tb.paragraphs[0] if idx == 0 else tb.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size); p.font.name = "Consolas"; p.font.color.rgb = color; p.space_after = Pt(1)


# Slide 1 - Framework definition
slide = prs.slides.add_slide(prs.slide_layouts[6])
bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.42))
bar.fill.solid(); bar.fill.fore_color.rgb = NAVY; bar.line.fill.background()

title = slide.shapes.add_textbox(Inches(0.7), Inches(0.8), Inches(9.5), Inches(0.6))
tf = title.text_frame
p = tf.paragraphs[0]
p.text = "Data Quality Framework"
p.font.size = Pt(26); p.font.bold = True; p.font.color.rgb = NAVY

sub = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(10.0), Inches(0.7))
sub_tf = sub.text_frame
p = sub_tf.paragraphs[0]
p.text = "Framework technique réutilisable pour contrôler les données en entrée, en transformation et en sortie"
p.font.size = Pt(17); p.font.color.rgb = GRAY

add_bullets(slide, 0.8, 2.1, 6.2, 3.2, [
    "Objectif : standardiser la qualité des données dans tous les flux d’information.",
    "Piliers : gouvernance, règles de validation, automatisation, observabilité et actions correctives.",
    "Principe : les contrôles sont appliqués dès la source, pendant la transformation et avant la diffusion aux usages métier.",
    "Valeur : réduire les erreurs, améliorer la confiance, accélérer l’onboarding de nouvelles équipes."
], font_size=15)

panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.5), Inches(2.1), Inches(4.9), Inches(3.1))
panel.fill.solid(); panel.fill.fore_color.rgb = LIGHT; panel.line.color.rgb = BLUE; panel.line.width = 1
pt = panel.text_frame; pt.word_wrap = True
p = pt.paragraphs[0]; p.text = "4 piliers du framework"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
for i, txt in enumerate(["Gouvernance", "Contrôles standards", "Plateforme technique", "Feedback / actions"]):
    p2 = pt.add_paragraph(); p2.text = txt; p2.bullet = True; p2.font.size = Pt(12); p2.font.color.rgb = TEXT

footer = slide.shapes.add_textbox(Inches(0.7), Inches(6.5), Inches(12.0), Inches(0.45))
ft = footer.text_frame
p = ft.paragraphs[0]
p.text = "Le framework est conçu comme un socle partagé, réutilisable par toutes les équipes sans réécrire la logique à chaque fois."
p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = NAVY

# Slide 2 - Technical architecture
slide = prs.slides.add_slide(prs.slide_layouts[6])
section_title(slide, "1. Architecture technique du framework")

add_card(slide, 0.7, 1.4, 2.2, 1.5, "Sources", "API, fichiers, ERP, CRM, DB, event streams.", BLUE)
add_card(slide, 3.2, 1.4, 2.2, 1.5, "Ingestion", "Load, staging, provenance, ownership.", CYAN)
add_card(slide, 5.7, 1.4, 2.2, 1.5, "Quality Layer", "Rules engine, validation, scoring, alerts.", GREEN)
add_card(slide, 8.2, 1.4, 2.2, 1.5, "Transformation", "dbt, SQL, Python, business logic.", ORANGE)
add_card(slide, 10.7, 1.4, 2.0, 1.5, "Output", "KPI, tables, dashboards, downstream usage.", RED)

for x in [2.9, 5.4, 7.9, 10.4]:
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(2.9), Inches(0.9), Inches(0.25))
    arrow.fill.solid(); arrow.fill.fore_color.rgb = BLUE; arrow.line.fill.background()

add_code_block(slide, 0.8, 3.6, 5.8, 2.2, [
    "# Exemple de règle métier",
    "rules:",
    "  - name: customer_id_required",
    "    type: completeness",
    "    field: customer_id",
    "    threshold: 99.5",
    "    severity: critical",
    "    owner: finance_team"
], fill_color=LIGHT)

add_code_block(slide, 6.8, 3.6, 5.7, 2.2, [
    "SELECT",
    "  COUNT(*) AS invalid_rows",
    "FROM {{ ref('customer_dim') }}",
    "WHERE customer_id IS NULL",
    "   OR status NOT IN ('ACTIVE', 'INACTIVE')",
], fill_color=WHITE)

add_bullets(slide, 0.9, 6.0, 11.8, 0.6, [
    "Le socle technique repose sur SQL/dbt, YAML, orchestration de jobs et exposition des résultats via tables/référentiels et alertes."
], font_size=11)

# Slide 3 - Controls by layer
slide = prs.slides.add_slide(prs.slide_layouts[6])
section_title(slide, "2. Contrôles par couche")

left = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.2), Inches(5.8), Inches(5.3))
left.fill.solid(); left.fill.fore_color.rgb = LIGHT; left.line.color.rgb = BLUE
lt = left.text_frame
p = lt.paragraphs[0]; p.text = "Contrôles d’entrée"; p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = NAVY
entry = [
    "Validation du schéma : colonnes, types, nullabilité, encodage.",
    "Contrôle de format : date, montant, codes pays, identifiants.",
    "Conformité métier : valeurs autorisées, référentiel, statut validé.",
    "Qualité de volume : seuils minimum/maximum par batch et par partition.",
    "Détection de doublons / clés invalides / enrichissement incomplet.",
    "SLA de réception : pas de retard > tolérance ou absence de livraison."
]
add_bullets(slide, 0.9, 2.0, 5.0, 3.9, entry, font_size=12)

right = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.2), Inches(5.8), Inches(5.3))
right.fill.solid(); right.fill.fore_color.rgb = LIGHT; right.line.color.rgb = GREEN
rt = right.text_frame
p = rt.paragraphs[0]; p.text = "Contrôles de transformation et de sortie"; p.font.bold = True; p.font.size = Pt(16); p.font.color.rgb = GREEN
outcome = [
    "Cohérence métier : joins, agrégations, règles calculaires, référentiels.",
    "Contrôle de duplication : clés métier, clés de fusion, partitions non cohérentes.",
    "Contrôle de résultat : volumes, montants, ratios et distribution par dimension.",
    "Reconciliation source -> cible : delta, lignes absentes, différences de clés.",
    "Contrôle de publication : score final, seuil d’acceptation, alertes avant diffusion.",
    "Signal de qualité : tableau de bord, non-conformité, runbook de correction."
]
add_bullets(slide, 7.1, 2.0, 5.0, 3.9, outcome, font_size=12)

# Slide 4 - How teams use it
slide = prs.slides.add_slide(prs.slide_layouts[6])
section_title(slide, "3. Comment l’utiliser par les autres équipes")

add_card(slide, 0.7, 1.5, 2.8, 1.8, "Équipe métier", "Définit les règles, la criticité, les seuils et les exceptions.", BLUE)
add_card(slide, 3.9, 1.5, 2.8, 1.8, "Data engineer", "Déploie les tests dbt, jobs et orchestrations.", CYAN)
add_card(slide, 7.1, 1.5, 2.8, 1.8, "Platform / governance", "Publie le socle partagé et les standards.", GREEN)
add_card(slide, 10.3, 1.5, 2.1, 1.8, "Support / Ops", "Traite les alertes et les écarts.", RED)

shared = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.0), Inches(4.1), Inches(7.2), Inches(1.5))
shared.fill.solid(); shared.fill.fore_color.rgb = LIGHT; shared.line.color.rgb = NAVY
st = shared.text_frame
p = st.paragraphs[0]; p.text = "Modèle d’usage partagé"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
p2 = st.add_paragraph(); p2.text = "Chaque équipe apporte son domaine et sa criticité ; la plateforme met à disposition les règles standards, l’orchestration, l’observabilité et les alertes."; p2.font.size = Pt(12); p2.font.color.rgb = TEXT

add_bullets(slide, 0.8, 5.8, 12.0, 0.55, [
    "Le framework est réutilisable : il ne s’applique pas à un seul flux mais à tout portefeuille de données avec le même mécanisme d’exécution et de reporting."
], font_size=11)

# Slide 5 - Deployment & backlog
slide = prs.slides.add_slide(prs.slide_layouts[6])
section_title(slide, "4. Déploiement et backlog de mise en œuvre")

boxes = [
    (0.8, 1.8, 2.0, 1.1, "Dev", BLUE),
    (3.1, 1.8, 2.0, 1.1, "UAT", CYAN),
    (5.4, 1.8, 2.0, 1.1, "Preprod", GREEN),
    (7.7, 1.8, 2.0, 1.1, "Prod", ORANGE),
    (10.0, 1.8, 2.0, 1.1, "Monitor", RED),
]
for x, y, w, h, label, color in boxes:
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid(); box.fill.fore_color.rgb = color; box.line.fill.background()
    tb = box.text_frame; p = tb.paragraphs[0]; p.text = label; p.font.bold = True; p.font.color.rgb = WHITE; p.font.size = Pt(15)
for x1, x2 in [(2.8,3.1), (5.1,5.4), (7.4,7.7), (9.7,10.0)]:
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x1), Inches(2.2), Inches(x2-x1), Inches(0.25))
    arrow.fill.solid(); arrow.fill.fore_color.rgb = BLUE; arrow.line.fill.background()

add_bullets(slide, 0.8, 3.4, 5.4, 2.6, [
    "0-30 jours : cartographie des flux et des data owners.",
    "30-60 jours : implémentation du socle technique et des règles standards.",
    "60-90 jours : pilotage sur 2 à 3 flux, alertes et tableau de bord.",
    "90-120 jours : industrialisation pour plusieurs équipes et standardisation des seuils.",
    "120+ jours : extension, amélioration continue et gouvernance partagée."
], font_size=12)

add_code_block(slide, 6.8, 3.4, 5.6, 2.4, [
    "dbt test --select tag:data_quality",
    "dbt run --select tag:data_quality",
    "python dq_runner.py --env prod --team marketing",
    "notify alerting --domain marketing --severity high",
], fill_color=LIGHT)

check = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.1), Inches(12.0), Inches(0.55))
check.fill.solid(); check.fill.fore_color.rgb = LIGHT; check.line.color.rgb = NAVY
ct = check.text_frame
p = ct.paragraphs[0]; p.text = "Résultat attendu : un framework standardisé, automatisé, déployable par plusieurs équipes avec des règles partagées et une observabilité commune."; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = NAVY

prs.save(OUT)
print(f"PowerPoint created: {OUT}")

