from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO

def generate_pptx():
    prs = Presentation()

    # Slide 1 – Titelpagina
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = f"Instructieboek: {projectnaam}"
    slide.placeholders[1].text = f"Student: {naam} - {studentnummer}"

    # Slide 2 – Gegevens
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Gegevens"
    content = (
        f"Naam: {naam}\n"
        f"Studentnummer: {studentnummer}\n"
        f"Project: {projectnaam}\n"
        f"Locatie: {locatie}\n"
        f"Leerbedrijf: {leerbedrijf}\n"
        f"Leermeester: {leermeester}\n"
        f"Inleverdatum: {inleverdatum}"
    )
    slide.placeholders[1].text = content

    # Slide 3 – Praktijkopdracht
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Over de Praktijkopdracht"
    opdracht_info = (
        f"Opdracht: {opdracht}\n\n"
        f"Wat heb je gemaakt: {wat_gemaakt}\n\n"
        f"Waarom: {waarom_gemaakt}\n\n"
        f"Type werk: {type_werk}\n"
        f"Werksituatie: {werksituatie}\n"
        f"Ploeggrootte: {ploeggrootte}"
    )
    slide.placeholders[1].text = opdracht_info

    # Slide 4 – Risico's
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Risico's en Maatregelen"
    slide.placeholders[1].text = f"Risico’s:\n{risicos}\n\nMaatregelen:\n{maatregelen}"

    # Slides 5+ – Stappenplan
    for i in range(1, 11):
        titel = st.session_state.get(f"stap{i}_titel", "")
        wat = st.session_state.get(f"stap{i}_wat", "")
        waarom = st.session_state.get(f"stap{i}_waarom", "")
        leer = st.session_state.get(f"stap{i}_leer", "")
        instructie = st.session_state.get(f"stap{i}_instructie", "")
        letop = st.session_state.get(f"stap{i}_letop", "")
        if titel.strip() == "" and wat.strip() == "":
            continue  # sla lege stappen over

        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"Stap {i}: {titel}"
        slide.placeholders[1].text = (
            f"Wat gedaan: {wat}\n\n"
            f"Waarom: {waarom}\n\n"
            f"Leerpunten: {leer}\n\n"
            f"Instructie: {instructie}\n\n"
            f"Let op!: {letop}"
        )

    # Slide – Reflectie
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Reflectie"
    reflectie_text = (
        st.session_state.get("Hoeveel hulp had je nodig en wat kon je zelfstandig?", "") + "\n" +
        st.session_state.get("Wanneer stuurde een collega je bij?", "") + "\n" +
        st.session_state.get("Welke tips heb je gekregen?", "") + "\n" +
        st.session_state.get("Wat waren je leerpunten?", "") + "\n" +
        st.session_state.get("Wat waren je sterke punten?", "")
    )
    slide.placeholders[1].text = reflectie_text

    # Opslaan naar buffer
    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# Download knop
st.subheader("⬇️ Download jouw PowerPoint")
if st.button("Genereer PowerPoint (.pptx)"):
    pptx_file = generate_pptx()
    st.download_button(
        label="Download PowerPoint",
        data=pptx_file,
        file_name="instructieboek_praktijkopdracht.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

