import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import tempfile

st.title("PowerPoint generator met vaste eerste dia en 25 dia's")

# Eerste dia vaste gegevens invullen
st.header("Eerste dia: Gegevens")

# Veld voor praktijkopdracht titel (valt onder de "Titel van de praktijkopdracht")
praktijkopdracht_titel = st.text_input("Titel van de praktijkopdracht:")

student_naam = st.text_input("Naam student:")
student_nummer = st.text_input("Studenten nummer:")
project_naam = st.text_input("Naam project:")
project_locatie = st.text_input("Locatie project:")
leerbedrijf = st.text_input("Leerbedrijf:")
leermeester = st.text_input("Leermeester:")
inleverdatum = st.text_input("Inleverdatum:")

# Dia's 2 t/m 25
NUM_SLIDES = 25

slides_content = []
for i in range(2, NUM_SLIDES+1):
    st.header(f"Slide {i}")
    title = st.text_input(f"Slide {i} titel", key=f"title_{i}")
    content = st.text_area(f"Slide {i} inhoud", key=f"content_{i}")
    img = st.file_uploader(f"Upload afbeelding voor slide {i} (optioneel)", type=["png","jpg","jpeg"], key=f"img_{i}")
    slides_content.append({"title": title, "content": content, "image": img})

def maak_pptx(first_slide_data, slides):
    prs = Presentation()

    # Eerste dia: vaste tekst + gegevens
    slide_layout = prs.slide_layouts[5]  # lege slide
    sld = prs.slides.add_slide(slide_layout)
    # Titel vaste tekst
    if sld.shapes.title:
        sld.shapes.title.text = "Stappenplan praktijkopdracht"
    else:
        textbox = sld.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
        textbox.text_frame.text = "Stappenplan praktijkopdracht"

    # Subtitel praktijkopdracht titel
    subtitle_box = sld.shapes.add_textbox(Inches(0.5), Inches(1), Inches(9), Inches(0.5))
    subtitle_box.text_frame.text = first_slide_data['praktijkopdracht_titel'] or "Titel van de praktijkopdracht"

    # De rest van de gegevens
    inhoud = (
        f"Naam student: {first_slide_data['student_naam']}\n"
        f"Studenten nummer: {first_slide_data['student_nummer']}\n"
        f"Naam project: {first_slide_data['project_naam']}\n"
        f"Locatie project: {first_slide_data['project_locatie']}\n"
        f"Leerbedrijf: {first_slide_data['leerbedrijf']}\n"
        f"Leermeester: {first_slide_data['leermeester']}\n"
        f"Inleverdatum: {first_slide_data['inleverdatum']}\n"
    )
    txBox = sld.shapes.add_textbox(Inches(0.5), Inches(1.6), Inches(9), Inches(4))
    txBox.text_frame.text = inhoud

    # Overige slides
    for slide in slides:
        sld = prs.slides.add_slide(prs.slide_layouts[5])
        if sld.shapes.title:
            sld.shapes.title.text = slide['title'] or ""
        else:
            textbox = sld.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
            textbox.text_frame.text = slide['title'] or ""
        txBox = sld.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(6), Inches(3))
        txBox.text_frame.text = slide['content'] or ""

        if slide['image'] is not None:
            try:
                with tempfile.NamedTemporaryFile(delete=True, suffix=".png") as tmp:
                    tmp.write(slide['image'].getbuffer())
                    tmp.flush()
                    sld.shapes.add_picture(tmp.name, Inches(7), Inches(1.5), width=Inches(2), height=Inches(2))
            except Exception as e:
                print(f"Fout bij toevoegen afbeelding: {e}")

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

if st.button("Genereer PowerPoint"):
    if not praktijkopdracht_titel or not student_naam or not project_naam:
        st.error("Vul minimaal Titel van de praktijkopdracht, Naam student en Naam project in de eerste dia in.")
    else:
        pptx_file = maak_pptx(
            {
                "praktijkopdracht_titel": praktijkopdracht_titel,
                "student_naam": student_naam,
                "student_nummer": student_nummer,
                "project_naam": project_naam,
                "project_locatie": project_locatie,
                "leerbedrijf": leerbedrijf,
                "leermeester": leermeester,
                "inleverdatum": inleverdatum,
            },
            slides_content,
        )
        st.success("PowerPoint is klaar!")
        st.download_button(
            label="Download PowerPoint (.pptx)",
            data=pptx_file,
            file_name="mijn_presentatie.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
