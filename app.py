import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from io import BytesIO

st.title("PowerPoint Generator zonder AI")

st.write("Plak hieronder je grote tekst (zoals jouw voorbeeld).")

# Tekst invoer
grote_tekst = st.text_area("Voer hier je tekst in:", height=400)

def generate_pptx(text):
    prs = Presentation()

    def add_slide(title, content_lines):
        slide_layout = prs.slide_layouts[1]  # Titel + inhoud
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = title
        text_box = slide.shapes.placeholders[1].text_frame
        text_box.clear()
        for line in content_lines:
            p = text_box.add_paragraph()
            p.text = line
            p.font.size = Pt(14)

    # Splits tekst in paragrafen (dubbele enters)
    paragrafen = [p.strip() for p in text.split('\n\n') if p.strip()]

    # Let op: pas deze indexen aan afhankelijk van jouw tekstindeling
    # Dit voorbeeld gaat uit van ongeveer dezelfde structuur als jouw voorbeeld
    try:
        add_slide("Gegevens", paragrafen[0].split('\n'))
        add_slide("Praktijkopdracht", paragrafen[1:6])
        add_slide("Risico’s en maatregelen", paragrafen[6:8])
        add_slide("Materiaalstaat", paragrafen[8:12])
        add_slide("Gereedschapslijst", paragrafen[12:16])
        add_slide("Werkschema en urenverantwoording", paragrafen[16:20])
        add_slide("Stap 1 – Lezen en begrijpen van de werktekening", paragrafen[20:27])
        add_slide("Reflectie: persoonlijk", paragrafen[27:34])
        add_slide("Reflectie: Uitvoering en samenwerken", paragrafen[34:40])
    except IndexError:
        # Als tekst niet lang genoeg is, sla door zonder crash
        pass

    bio = BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

if grote_tekst:
    pptx_file = generate_pptx(grote_tekst)
    st.download_button(
        label="Download PowerPoint",
        data=pptx_file,
        file_name="praktijkopdracht.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
else:
    st.info("Plak een grote tekst om een PowerPoint te genereren.")

