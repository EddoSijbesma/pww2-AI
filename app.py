import streamlit as st
import json
from io import BytesIO
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN

def add_textbox(slide, text, left, top, width, height, font_size=12, bold=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.alignment = PP_ALIGN.LEFT
    return txBox

def create_presentation(data):
    prs = Presentation()

    # DIA 1: Gegevens
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_textbox(slide, "Gegevens", Pt(50), Pt(20), Pt(600), Pt(40), font_size=24, bold=True)

    content = (
        f"Naam student: {data['naam_student']}\n"
        f"Studenten nummer: {data['studenten_nummer']}\n"
        f"Naam project: {data['naam_project']}\n"
        f"Locatie project: {data['locatie_project']}\n"
        f"Leerbedrijf: {data['leerbedrijf']}\n"
        f"Leermeester: {data['leermeester']}\n"
        f"Inleverdatum: {data['inleverdatum']}\n"
    )
    add_textbox(slide, content, Pt(50), Pt(80), Pt(600), Pt(200), font_size=14)

    # DIA 2: Praktijkopdracht
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_textbox(slide, "Welke praktijkopdracht heb je gemaakt?", Pt(50), Pt(20), Pt(600), Pt(40), font_size=20, bold=True)

    po = data['praktijkopdracht']
    fields = [
        ("Wat heb je gemaakt.", po.get('wat_heb_je_gemaakt', '')),
        ("Waarom heb je deze praktijkopdracht gemaakt?", po.get('waarom', '')),
        ("Wat voor type werk was het?", po.get('type_werk', '')),
        ("Hoe was de werksituatie?", po.get('werksituatie', '')),
        ("Hoe groot was je ploeg?", po.get('groot_ploeg', ''))
    ]

    top = 80
    for label, tekst in fields:
        text = f"{label}\n{tekst}"
        add_textbox(slide, text, Pt(50), Pt(top), Pt(600), Pt(60), font_size=12)
        top += 70

    # DIA 3: Risico's
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_textbox(slide, "Beschrijf de risico’s bij deze praktijkopdracht\n(en de genomen maatregelen)", Pt(50), Pt(20), Pt(600), Pt(40), font_size=20, bold=True)

    lefts = [Pt(50), Pt(300)]
    tops = [Pt(80 + i*25) for i in range(8)]
    widths = [Pt(230), Pt(230)]
    heights = Pt(25)

    add_textbox(slide, "Risico", lefts[0], Pt(50), widths[0], heights, font_size=14, bold=True)
    add_textbox(slide, "Genomen maatregel", lefts[1], Pt(50), widths[1], heights, font_size=14, bold=True)

    for i, item in enumerate(data['risicos']):
        add_textbox(slide, item.get('risico', ''), lefts[0], tops[i], widths[0], heights, font_size=12)
        add_textbox(slide, item.get('maatregel', ''), lefts[1], tops[i], widths[1], heights, font_size=12)

    # DIA 4: Materiaalstaat
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_textbox(slide, "Materiaalstaat", Pt(50), Pt(20), Pt(600), Pt(40), font_size=20, bold=True)

    lefts = [Pt(50), Pt(300), Pt(450)]
    headers = ["Materiaal", "Maat", "Aantal"]
    for i, header in enumerate(headers):
        add_textbox(slide, header, lefts[i], Pt(50), Pt(140), Pt(25), font_size=14, bold=True)

    tops = [Pt(80 + i*25) for i in range(12)]
    for i, mat in enumerate(data['materiaalstaat']):
        add_textbox(slide, mat.get('materiaal', ''), lefts[0], tops[i], Pt(140), Pt(25), font_size=12)
        add_textbox(slide, mat.get('maat', ''), lefts[1], tops[i], Pt(100), Pt(25), font_size=12)
        add_textbox(slide, mat.get('aantal', ''), lefts[2], tops[i], Pt(100), Pt(25), font_size=12)

    # DIA 5: Gereedschapslijst
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_textbox(slide, "Gereedschapslijst", Pt(50), Pt(20), Pt(600), Pt(40), font_size=20, bold=True)

    lefts = [Pt(50), Pt(350)]
    headers = ["Gereedschap", "Gebruikt voor"]
    for i, header in enumerate(headers):
        add_textbox(slide, header, lefts[i], Pt(50), Pt(280), Pt(25), font_size=14, bold=True)

    tops = [Pt(80 + i*25) for i in range(12)]
    for i, gereedschap in enumerate(data['gereedschapslijst']):
        add_textbox(slide, gereedschap.get('gereedschap', ''), lefts[0], tops[i], Pt(280), Pt(25), font_size=12)
        add_textbox(slide, gereedschap.get('gebruikt_voor', ''), lefts[1], tops[i], Pt(280), Pt(25), font_size=12)

    # Save presentation to BytesIO object
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

st.title("PowerPoint Generator voor Praktijkopdrachten")

st.write("""
Upload een JSON bestand met je projectdata, of plak de JSON hieronder.
De app genereert een PowerPoint-bestand dat je kunt downloaden.
""")

uploaded_file = st.file_uploader("Upload JSON bestand", type=["json"])

json_text = None

if uploaded_file is not None:
    json_text = uploaded_file.read().decode("utf-8")
else:
    json_text = st.text_area("Of plak hier je JSON data", height=300)

if json_text:
    try:
        data = json.loads(json_text)
        if st.button("Genereer PowerPoint"):
            pptx_io = create_presentation(data)
            st.success("PowerPoint is aangemaakt!")
            st.download_button(
                label="Download PowerPoint",
                data=pptx_io,
                file_name="praktijkopdracht.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    except Exception as e:
        st.error(f"Fout bij verwerken JSON: {e}")

