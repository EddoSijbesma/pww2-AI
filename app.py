import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from io import BytesIO

st.set_page_config(page_title="AI PowerPoint Generator", layout="wide")

st.title("📊 AI PowerPoint Generator - Praktijkopdracht")
st.markdown("Vul de gegevens in, inclusief risicoanalyse, en download een nette PowerPoint!")

# --- Formuliergegevens
st.header("📌 Gegevens student & project")
student_naam = st.text_input("Naam student")
student_nummer = st.text_input("Studentnummer")
project_naam = st.text_input("Naam project")
project_locatie = st.text_input("Locatie project")
leerbedrijf = st.text_input("Leerbedrijf")
leermeester = st.text_input("Leermeester")
inleverdatum = st.date_input("Inleverdatum")

st.header("📊 Dia 4 - Risicoanalyse")
risico_maatregelen = []
for i in range(8):
    col1, col2 = st.columns(2)
    with col1:
        risico = st.text_input(f"Risico {i+1}", key=f"risico_{i}")
    with col2:
        maatregel = st.text_input(f"Maatregel {i+1}", key=f"maatregel_{i}")
    if risico or maatregel:
        risico_maatregelen.append((risico, maatregel))

st.header("🖼️ Foto's toevoegen")
uploaded_images = st.file_uploader("Upload afbeeldingen voor dia's (optioneel)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

# --- PPTX genereren
def maak_pptx():
    prs = Presentation()
    layout = prs.slide_layouts[5]

    # Dia 1: Titel
    slide = prs.slides.add_slide(layout)
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = f"Stappenplan praktijkopdracht\n{project_naam}"
    p.font.size = Pt(36)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER

    # Dia 2: Gegevens
    slide = prs.slides.add_slide(layout)
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
    tf = textbox.text_frame
    tf.text = "Projectgegevens"
    tf.paragraphs[0].font.size = Pt(28)
    tf.paragraphs[0].font.bold = True
    tf.add_paragraph().text = f"Naam student: {student_naam}"
    tf.add_paragraph().text = f"Studentnummer: {student_nummer}"
    tf.add_paragraph().text = f"Naam project: {project_naam}"
    tf.add_paragraph().text = f"Locatie project: {project_locatie}"
    tf.add_paragraph().text = f"Leerbedrijf: {leerbedrijf}"
    tf.add_paragraph().text = f"Leermeester: {leermeester}"
    tf.add_paragraph().text = f"Inleverdatum: {inleverdatum}"

    # Dia 3: Leeg met eventueel afbeelding
    slide = prs.slides.add_slide(layout)
    tf = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
    tf.text = "Introductie"
    
    # Dia 4: Risicoanalyse met mooie layout
    slide = prs.slides.add_slide(layout)

    # Titel bovenaan
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
    tf = title_box.text_frame
    p = tf.add_paragraph()
    p.text = "Risicoanalyse Praktijkopdracht"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 51, 102)
    p.alignment = PP_ALIGN.CENTER

    # Subtitel
    sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(1), Inches(9), Inches(0.6))
    tf_sub = sub_box.text_frame
    p_sub = tf_sub.add_paragraph()
    p_sub.text = "Overzicht van risico’s en genomen maatregelen"
    p_sub.font.size = Pt(20)
    p_sub.font.italic = True
    p_sub.font.color.rgb = RGBColor(90, 90, 90)
    p_sub.alignment = PP_ALIGN.CENTER

    # Tabel
    if risico_maatregelen:
        rows = len(risico_maatregelen) + 1
        cols = 2
        table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.7), Inches(9), Inches(4.5)).table
        table.columns[0].width = Inches(4.5)
        table.columns[1].width = Inches(4.5)
        table.cell(0, 0).text = "Risico"
        table.cell(0, 1).text = "Genomen maatregel"
        for col in range(cols):
            header_cell = table.cell(0, col)
            header_cell.text_frame.paragraphs[0].font.bold = True
            header_cell.text_frame.paragraphs[0].font.size = Pt(16)
            header_cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        for i, (risico, maatregel) in enumerate(risico_maatregelen):
            table.cell(i+1, 0).text = risico
            table.cell(i+1, 1).text = maatregel

    # Dia’s 5 t/m 25: lege dia’s of met afbeelding
    for i in range(21):
        slide = prs.slides.add_slide(layout)
        if i < len(uploaded_images):
            image = uploaded_images[i]
            slide.shapes.add_picture(image, Inches(1), Inches(1.5), width=Inches(7.5))
        else:
            tf = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5)).text_frame
            tf.text = f"Dia {i+5}"

    return prs

# --- Genereer en download
if st.button("🎬 Genereer PowerPoint"):
    prs = maak_pptx()
    pptx_io = BytesIO()
    prs.save(pptx_io)
    st.success("✅ PowerPoint gegenereerd!")
    st.download_button("📥 Download PowerPoint", data=pptx_io.getvalue(), file_name="praktijkopdracht.pptx")
