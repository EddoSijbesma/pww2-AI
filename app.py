import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from io import BytesIO
import openai
import os
import json

openai.api_key = os.getenv("OPENAI_API_KEY")  # Zorg dat je hier je API key hebt staan

def add_textbox(slide, text, left, top, width, height, font_size=12, bold=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.alignment = PP_ALIGN.LEFT
    return txBox

def parse_text_to_fields(text):
    prompt = f"""
Je krijgt een grote tekst van een praktijkopdracht. Verdeel deze informatie in een JSON object met de volgende velden:

{{
  "naam_student": "...",
  "studenten_nummer": "...",
  "naam_project": "...",
  "locatie_project": "...",
  "leerbedrijf": "...",
  "leermeester": "...",
  "inleverdatum": "...",
  "praktijkopdracht": {{
    "wat_heb_je_gemaakt": "...",
    "waarom": "...",
    "type_werk": "...",
    "werksituatie": "...",
    "groot_ploeg": "..."
  }},
  "risicos": [
    {{"risico": "...", "maatregel": "..."}},
    {{"risico": "...", "maatregel": "..."}}
  ],
  "materiaalstaat": [
    {{"materiaal": "...", "maat": "...", "aantal": "..."}}
  ],
  "gereedschapslijst": [
    {{"gereedschap": "...", "gebruikt_voor": "..."}}
  ]
}}

Hier is de tekst:
\"\"\"
{text}
\"\"\"

Geef alleen het JSON object terug.
"""
    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
        max_tokens=1000
    )
    content = response['choices'][0]['message']['content']
    return content

def create_presentation(data):
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_textbox(slide, "Gegevens", Pt(50), Pt(20), Pt(600), Pt(40), font_size=24, bold=True)

    content = (
        f"Naam student: {data.get('naam_student','')}\n"
        f"Studenten nummer: {data.get('studenten_nummer','')}\n"
        f"Naam project: {data.get('naam_project','')}\n"
        f"Locatie project: {data.get('locatie_project','')}\n"
        f"Leerbedrijf: {data.get('leerbedrijf','')}\n"
        f"Leermeester: {data.get('leermeester','')}\n"
        f"Inleverdatum: {data.get('inleverdatum','')}\n"
    )
    add_textbox(slide, content, Pt(50), Pt(80), Pt(600), Pt(200), font_size=14)

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_textbox(slide, "Welke praktijkopdracht heb je gemaakt?", Pt(50), Pt(20), Pt(600), Pt(40), font_size=20, bold=True)

    po = data.get('praktijkopdracht', {})
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

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_textbox(slide, "Beschrijf de risico’s bij deze praktijkopdracht\n(en de genomen maatregelen)", Pt(50), Pt(20), Pt(600), Pt(40), font_size=20, bold=True)

    lefts = [Pt(50), Pt(300)]
    tops = [Pt(80 + i*25) for i in range(8)]
    widths = [Pt(230), Pt(230)]
    heights = Pt(25)

    add_textbox(slide, "Risico", lefts[0], Pt(50), widths[0], heights, font_size=14, bold=True)
    add_textbox(slide, "Genomen maatregel", lefts[1], Pt(50), widths[1], heights, font_size=14, bold=True)

    risicos = data.get('risicos', [])
    for i in range(8):
        if i < len(risicos):
            item = risicos[i]
            risico = item.get('risico', '')
            maatregel = item.get('maatregel', '')
        else:
            risico = ""
            maatregel = ""
        add_textbox(slide, risico, lefts[0], tops[i], widths[0], heights, font_size=12)
        add_textbox(slide, maatregel, lefts[1], tops[i], widths[1], heights, font_size=12)

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_textbox(slide, "Materiaalstaat", Pt(50), Pt(20), Pt(600), Pt(40), font_size=20, bold=True)

    lefts = [Pt(50), Pt(300), Pt(450)]
    headers = ["Materiaal", "Maat", "Aantal"]
    for i, header in enumerate(headers):
        add_textbox(slide, header, lefts[i], Pt(50), Pt(140), Pt(25), font_size=14, bold=True)

    materiaalstaat = data.get('materiaalstaat', [])
    for i in range(12):
        if i < len(materiaalstaat):
            mat = materiaalstaat[i]
            materiaal = mat.get('materiaal', '')
            maat = mat.get('maat', '')
            aantal = mat.get('aantal', '')
        else:
            materiaal = ""
            maat = ""
            aantal = ""
        add_textbox(slide, materiaal, lefts[0], Pt(80 + i*25), Pt(140), Pt(25), font_size=12)
        add_textbox(slide, maat, lefts[1], Pt(80 + i*25), Pt(100), Pt(25), font_size=12)
        add_textbox(slide, aantal, lefts[2], Pt(80 + i*25), Pt(100), Pt(25), font_size=12)

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    add_textbox(slide, "Gereedschapslijst", Pt(50), Pt(20), Pt(600), Pt(40), font_size=20, bold=True)

    lefts = [Pt(50), Pt(350)]
    headers = ["Gereedschap", "Gebruikt voor"]
    for i, header in enumerate(headers):
        add_textbox(slide, header, lefts[i], Pt(50), Pt(280), Pt(25), font_size=14, bold=True)

    gereedschapslijst = data.get('gereedschapslijst', [])
    for i in range(12):
        if i < len(gereedschapslijst):
            gereedschap = gereedschapslijst[i]
            naam = gereedschap.get('gereedschap', '')
            gebruikt_voor = gereedschap.get('gebruikt_voor', '')
        else:
            naam = ""
            gebruikt_voor = ""
        add_textbox(slide, naam, lefts[0], Pt(80 + i*25), Pt(280), Pt(25), font_size=12)
        add_textbox(slide, gebruikt_voor, lefts[1], Pt(80 + i*25), Pt(280), Pt(25), font_size=12)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

st.title("PowerPoint Generator van Lap Tekst")

text = st.text_area("Plak hier je grote tekst van de praktijkopdracht", height=400)

if st.button("Genereer PowerPoint"):
    if not text.strip():
        st.error("Voer eerst wat tekst in.")
    else:
        with st.spinner("Bezig met verwerken..."):
            try:
                json_str = parse_text_to_fields(text)
                data = json.loads(json_str)
                pptx_io = create_presentation(data)
                st.success("PowerPoint is aangemaakt!")
                st.download_button(
                    label="Download PowerPoint",
                    data=pptx_io,
                    file_name="praktijkopdracht.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"Er ging iets mis: {e}")
