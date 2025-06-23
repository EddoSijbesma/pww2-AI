import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from io import BytesIO
import re

st.set_page_config(page_title="PowerPoint Generator", layout="wide")
st.title("📊 PowerPoint Generator Praktijkopdracht")

# === Helper functie voor dummy AI extractie ===
def extract_risico_maatregelen(text):
    # Zoekt regels die met 'risico:' of 'maatregel:' beginnen, case insensitive
    risicoes = re.findall(r'risico:\s*(.+)', text, flags=re.IGNORECASE)
    maatregelen = re.findall(r'maatregel:\s*(.+)', text, flags=re.IGNORECASE)
    risico_maatregelen = []
    for i in range(min(len(risicoes), len(maatregelen))):
        risico_maatregelen.append((risicoes[i], maatregelen[i]))
    return risico_maatregelen

# === Stap 1: Gegevens voor eerste dia ===
st.header("🧾 Gegevens voor eerste dia")

col1, col2 = st.columns(2)
with col1:
    student_naam = st.text_input("Naam student")
    student_nummer = st.text_input("Studentnummer")
    project_naam = st.text_input("Naam project")
    praktijkopdracht_titel = st.text_input("Titel praktijkopdracht")

with col2:
    project_locatie = st.text_input("Locatie project")
    leerbedrijf = st.text_input("Leerbedrijf")
    leermeester = st.text_input("Leermeester")
    inleverdatum = st.date_input("Inleverdatum")

# === Stap 2: Dia's invoeren ===
st.header("📝 Dia's invoeren (behalve dia 4 en 9 t/m 18)")

slides = []

# Vragenlijst voor dia 9 t/m 18
vragen = [
    "Beschrijf hier wat je hebt gedaan.",
    "Waarom heb je het zo gedaan.",
    "Wat is was een leerpunt.",
    "Instructies voor je collega (wat is belangrijk om op te letten?).",
    "(Laat deze tekst op je foto’s aansluiten)",
    "Je kunt hierbij ook pijltjes toevoegen",
    "Voeg een “let op!” toe vanuit de deelopdracht."
]

for i in range(1, 26):
    if i == 4:
        continue  # Dia 4 is aparte stap
    if 9 <= i <= 18:
        st.markdown(f"### Dia {i} – Vaste vragen invullen")
        antwoorden = []
        for idx, vraag in enumerate(vragen):
            antwoord = st.text_area(f"{vraag}", key=f"dia_{i}_vraag_{idx}")
            antwoorden.append(antwoord)
        image = st.file_uploader(f"📷 Afbeelding dia {i} (optioneel)", type=["png", "jpg", "jpeg"], key=f"img_{i}")
        slides.append({
            "title": f"Stap {i - 8}",
            "content": antwoorden,
            "image": image,
            "is_vast_tekst": True
        })
    else:
        with st.expander(f"Dia {i} (Vrije tekst en afbeelding)"):
            title = st.text_input(f"🔹 Titel dia {i}", key=f"title_{i}")
            content = st.text_area(f"✏️ Inhoud dia {i}", key=f"content_{i}")
            image = st.file_uploader(f"📷 Afbeelding dia {i} (optioneel)", type=["png", "jpg", "jpeg"], key=f"img_{i}")
            slides.append({
                "title": title,
                "content": content,
                "image": image,
                "is_vast_tekst": False
            })

# === Stap 3: Dia 4 - Risico's en maatregelen ===
st.header("⚠️ Dia 4 – Risicoanalyse (tabel)")

st.markdown(
    "_Beschrijf de risico’s bij deze praktijkopdracht (veiligheid, tijdsdruk, e.a.)_  \n"
    "_En geef aan welke maatregelen je hebt getroffen om ze te beheersen._"
)

# Initialiseer risico_maatregelen in sessiestate
if "risico_maatregelen" not in st.session_state:
    st.session_state.risico_maatregelen = [("", "") for _ in range(8)]

cols = st.columns(2)
for i in range(8):
    with cols[0]:
        risico = st.text_input(f"Risico {i+1}", value=st.session_state.risico_maatregelen[i][0], key=f"risico_{i}")
    with cols[1]:
        maatregel = st.text_input(f"Maatregel {i+1}", value=st.session_state.risico_maatregelen[i][1], key=f"maatregel_{i}")
    st.session_state.risico_maatregelen[i] = (risico, maatregel)

# === AI Hulpmiddel ===
st.header("🧠 AI Hulpmiddel: Automatisch risico’s en maatregelen invullen")

ai_tekst = st.text_area("Voer hier een uitgebreide tekst in, bijvoorbeeld projectbeschrijving of analyse met risico’s en maatregelen")

if st.button("🤖 Vul risico’s en maatregelen automatisch in"):
    auto_risico_maatregelen = extract_risico_maatregelen(ai_tekst)
    if auto_risico_maatregelen:
        for i in range(min(len(auto_risico_maatregelen), len(st.session_state.risico_maatregelen))):
            st.session_state.risico_maatregelen[i] = auto_risico_maatregelen[i]
        st.success("Risico’s en maatregelen automatisch ingevuld!")
    else:
        st.warning("Geen risico’s en maatregelen gevonden in de tekst.")

# === Stap 4: PowerPoint genereren ===
def maak_pptx():
    prs = Presentation()
    layout = prs.slide_layouts[5]

    # --- Dia 1 – Titelpagina ---
    slide = prs.slides.add_slide(layout)

    box1 = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
    tf1 = box1.text_frame
    p = tf1.add_paragraph()
    p.text = "Stappenplan praktijkopdracht"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 51, 102)
    p.alignment = PP_ALIGN.CENTER

    box2 = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(9), Inches(0.7))
    tf2 = box2.text_frame
    p2 = tf2.add_paragraph()
    p2.text = praktijkopdracht_titel
    p2.font.size = Pt(24)
    p2.font.italic = True
    p2.font.color.rgb = RGBColor(100, 100, 100)
    p2.alignment = PP_ALIGN.CENTER

    box3 = slide.shapes.add_textbox(Inches(1), Inches(2.2), Inches(8), Inches(3))
    tf3 = box3.text_frame
    gegevens = [
        f"Naam student: {student_naam}",
        f"Studentnummer: {student_nummer}",
        f"Naam project: {project_naam}",
        f"Locatie project: {project_locatie}",
        f"Leerbedrijf: {leerbedrijf}",
        f"Leermeester: {leermeester}",
        f"Inleverdatum: {inleverdatum.strftime('%d-%m-%Y')}"
    ]
    for regel in gegevens:
        p = tf3.add_paragraph()
        p.text = regel
        p.font.size = Pt(18)

    # --- Dia 2 t/m 3 ---
    for slide_data in slides[:2]:
        if not slide_data["title"] and not slide_data["content"]:
            continue
        slide = prs.slides.add_slide(layout)
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
        tf = title_box.text_frame
        p = tf.add_paragraph()
        p.text = slide_data["title"] if slide_data["title"] else ""
        p.font.size = Pt(28)
        p.font.bold = True

        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(6.5), Inches(4))
        tf2 = content_box.text_frame
        p2 = tf2.add_paragraph()
        p2.text = slide_data["content"] if slide_data["content"] else ""
        p2.font.size = Pt(20)

        if slide_data["image"]:
            slide.shapes.add_picture(slide_data["image"], Inches(7), Inches(1.5), Inches(2.5), Inches(2.5))

    # --- Dia 4 - Risicoanalyse tabel ---
    slide = prs.slides.add_slide(layout)
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
    tf_title = title_box.text_frame
    p_title = tf_title.add_paragraph()
    p_title.text = "Risicoanalyse"
    p_title.font.size = Pt(32)
    p_title.font.bold = True
    p_title.font.color.rgb = RGBColor(0, 51, 102)
    p_title.alignment = PP_ALIGN.CENTER

    rows = 9  # 8 risico’s + header
    cols = 2
    left = Inches(0.5)
    top = Inches(1.5)
    width = Inches(9)
    height = Inches(4)

    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    table.columns[0].width = Inches(4.5)
    table.columns[1].width = Inches(4.5)

    # Header
    table.cell(0, 0).text = "Risico"
    table.cell(0, 1).text = "Genomen maatregel"
    for col in range(cols):
        cell = table.cell(0, col)
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.size = Pt(16)

    # Vul tabel met data uit sessiestate
    for i, (risico, maatregel) in enumerate(st.session_state.risico_maatregelen):
        table.cell(i+1, 0).text = risico if risico else ""
        table.cell(i+1, 1).text = maatregel if maatregel else ""
        for col in range(cols):
            cell = table.cell(i+1, col)
            cell.text_frame.paragraphs[0].font.size = Pt(14)

    # --- Dia 5 t/m 8 ---
    for slide_data in slides[2:6]:
        if not slide_data["title"] and not slide_data["content"]:
            continue
        slide = prs.slides.add_slide(layout)
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
        tf = title_box.text_frame
        p = tf.add_paragraph()
        p.text = slide_data["title"] if slide_data["title"] else ""
        p.font.size = Pt(28)
        p.font.bold = True

        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(6.5), Inches(4))
        tf2 = content_box.text_frame
        p2 = tf2.add_paragraph()
        p2.text = slide_data["content"] if slide_data["content"] else ""
        p2.font.size = Pt(20)

        if slide_data["image"]:
            slide.shapes.add_picture(slide_data["image"], Inches(7), Inches(1.5), Inches(2.5), Inches(2.5))

    # --- Dia 9 t/m 18 (vaste vragen ingevuld) ---
    for idx, i in enumerate(range(9, 19)):
        slide_data = slides[6 + idx]
        slide = prs.slides.add_slide(layout)

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
        tf_title = title_box.text_frame
        p_title = tf_title.add_paragraph()
        p_title.text = slide_data.get("title", f"Stap {i - 8}")
        p_title.font.size = Pt(28)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(0, 51, 102)

        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(5))
        tf = content_box.text_frame

        antwoorden = slide_data.get("content", [])
        for antwoord in antwoorden:
            p = tf.add_paragraph()
            p.text = antwoord if antwoord else "-"
            p.font.size = Pt(18)
            p.space_after = Pt(6)

        if slide_data.get("image"):
            slide.shapes.add_picture(slide_data["image"], Inches(7), Inches(1.5), Inches(2.5), Inches(2.5))

    # --- Dia 19 t/m 25 ---
    for slide_data in slides[16:]:
        if not slide_data["title"] and not slide_data["content"]:
            continue
        slide = prs.slides.add_slide(layout)
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
        tf = title_box.text_frame
        p = tf.add_paragraph()
        p.text = slide_data["title"] if slide_data["title"] else ""
        p.font.size = Pt(28)
        p.font.bold = True

        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(6.5), Inches(4))
        tf2 = content_box.text_frame
        p2 = tf2.add_paragraph()
        p2.text = slide_data["content"] if slide_data["content"] else ""
        p2.font.size = Pt(20)

        if slide_data["image"]:
            slide.shapes.add_picture(slide_data["image"], Inches(7), Inches(1.5), Inches(2.5), Inches(2.5))

    # --- Return bestand als bytes ---
    pptx_bytes = BytesIO()
    prs.save(pptx_bytes)
    pptx_bytes.seek(0)
    return pptx_bytes

if st.button("📥 Genereer PowerPoint"):
    pptx_file = maak_pptx()
    st.download_button("Download PowerPoint bestand", pptx_file, file_name="praktijkopdracht.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
