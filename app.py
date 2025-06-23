import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from io import BytesIO

st.set_page_config(page_title="PowerPoint Generator", layout="wide")
st.title("📊 PowerPoint Generator zonder AI")

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
st.header("📝 Dia's invoeren")

slides = []
for i in range(1, 26):
    if i == 4:
        continue  # Dia 4 is gereserveerd voor risico's en maatregelen
    with st.expander(f"Dia {i}"):
        title = st.text_input(f"🔹 Titel dia {i}", key=f"title_{i}")
        content = st.text_area(f"✏️ Inhoud dia {i}", key=f"content_{i}")
        image = st.file_uploader(f"📷 Afbeelding dia {i} (optioneel)", type=["png", "jpg", "jpeg"], key=f"img_{i}")
        slides.append({
            "title": title,
            "content": content,
            "image": image
        })

# === Stap 3: Dia 4 - Risico's en maatregelen ===
st.header("⚠️ Dia 4 – Risicoanalyse (tabel)")

st.markdown(
    "_Beschrijf de risico’s bij deze praktijkopdracht (veiligheid, tijdsdruk, e.a.)_  \n"
    "_En geef aan welke maatregelen je hebt getroffen om ze te beheersen._"
)

risico_maatregelen = []
cols = st.columns(2)
for i in range(8):
    with cols[0]:
        risico = st.text_input(f"Risico {i+1}", key=f"risico_{i}")
    with cols[1]:
        maatregel = st.text_input(f"Maatregel {i+1}", key=f"maatregel_{i}")
    risico_maatregelen.append((risico, maatregel))

# === Stap 4: PowerPoint genereren ===
def maak_pptx():
    prs = Presentation()
    layout = prs.slide_layouts[5]

    # Dia 1 – Titelpagina
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

    # Dia 2–3
    for slide_data in slides[:2]:
        slide = prs.slides.add_slide(layout)
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
        tf = title_box.text_frame
        p = tf.add_paragraph()
        p.text = slide_data["title"]
        p.font.size = Pt(28)
        p.font.bold = True

        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(6.5), Inches(4))
        tf2 = content_box.text_frame
        p2 = tf2.add_paragraph()
        p2.text = slide_data["content"]
        p2.font.size = Pt(20)

        if slide_data["image"]:
            slide.shapes.add_picture(slide_data["image"], Inches(7), Inches(1.5), Inches(2.5), Inches(2.5))

    # Dia 4 – Risicotabel
    slide = prs.slides.add_slide(layout)
    title_shape = slide.shapes.title
    title_shape.text = "Risicoanalyse"

    rows = len(risico_maatregelen) + 1
    cols = 2
    left = Inches(0.5)
    top = Inches(1.5)
    width = Inches(9)
    height = Inches(4)

    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    table.columns[0].width = Inches(4.5)
    table.columns[1].width = Inches(4.5)

    table.cell(0, 0).text = "Risico"
    table.cell(0, 1).text = "Genomen maatregel"

    for col in range(cols):
        cell = table.cell(0, col)
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.size = Pt(16)

    for i, (risico, maatregel) in enumerate(risico_maatregelen):
        table.cell(i+1, 0).text = risico
        table.cell(i+1, 1).text = maatregel
        for col in range(cols):
            cell = table.cell(i+1, col)
            cell.text_frame.paragraphs[0].font.size = Pt(14)

    # Dia 5 t/m 26
    for slide_data in slides[2:]:
        slide = prs.slides.add_slide(layout)

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
        tf = title_box.text_frame
        p = tf.add_paragraph()
        p.text = slide_data["title"]
        p.font.size = Pt(28)
        p.font.bold = True

        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(6.5), Inches(4))
        tf2 = content_box.text_frame
        p2 = tf2.add_paragraph()
        p2.text = slide_data["content"]
        p2.font.size = Pt(20)

        if slide_data["image"]:
            slide.shapes.add_picture(slide_data["image"], Inches(7), Inches(1.5), Inches(2.5), Inches(2.5))

    # Opslaan
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

if st.button("🎉 Genereer PowerPoint"):
    pptx_bestand = maak_pptx()
    st.success("✅ PowerPoint gegenereerd!")
    st.download_button(
        label="📥 Download PowerPoint",
        data=pptx_bestand,
        file_name="presentatie.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

