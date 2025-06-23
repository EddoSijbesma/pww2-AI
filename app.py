import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from io import BytesIO
import openai
import os
import json

st.set_page_config(page_title="AI PowerPoint Generator", layout="wide")
st.title("🤖 AI PowerPoint Generator voor Praktijkopdracht")

# --- OpenAI API key invoer (zorg dat je jouw API key invult of in omgeving zet) ---
if "OPENAI_API_KEY" not in st.session_state:
    st.session_state.OPENAI_API_KEY = ""

def set_api_key():
    st.session_state.OPENAI_API_KEY = st.text_input("Vul hier je OpenAI API key in (Begin met 'sk-'):", type="password")

if not st.session_state.OPENAI_API_KEY:
    set_api_key()
    st.stop()

openai.api_key = st.session_state.OPENAI_API_KEY

# --- Stap 1: Algemene opdracht invoeren ---
st.header("Stap 1: Voer de opdracht of het onderwerp in")

opdracht = st.text_area(
    "Typ hier de omschrijving van je praktijkopdracht of het onderwerp voor de presentatie.",
    height=150
)

# --- Functie om AI te laten genereren ---
def genereer_slides_via_ai(opdracht_tekst):
    # Prompt om gestructureerde output te krijgen (JSON) voor slides
    prompt = f"""
Je bent een assistent die een PowerPoint-presentatie maakt voor een praktijkopdracht.
Maak een presentatie van 20 dia's, inclusief titelpagina.
Elke dia moet een titel en inhoud hebben met 3 tot 5 korte kernpunten.
Lever de output in JSON met deze structuur:

{{
  "slides": [
    {{
      "title": "Titel van dia 1",
      "content": ["punt 1", "punt 2", "..."]
    }},
    ...
  ]
}}

De opdracht of het onderwerp is: "{opdracht_tekst}"

Schrijf de JSON zonder extra uitleg.
"""

    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "Je bent een behulpzame presentatiegenerator."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.7,
        max_tokens=1500,
    )

    text = response['choices'][0]['message']['content']

    # Probeer JSON te parsen, fallback op leeg
    try:
        slides_data = json.loads(text)
    except Exception as e:
        st.error("Fout bij het parsen van de AI-output, probeer opnieuw.")
        st.write(text)
        return None

    return slides_data.get("slides", [])

# --- Functie om pptx te maken ---
def maak_pptx(slides):
    prs = Presentation()
    layout = prs.slide_layouts[5]

    for slide_data in slides:
        slide = prs.slides.add_slide(layout)

        # Titel
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = slide_data.get("title", "")
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 51, 102)
        p.alignment = PP_ALIGN.CENTER

        # Content (lijst van punten)
        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(4))
        tf_content = content_box.text_frame
        tf_content.clear()  # Zorg dat het leeg is

        inhoud_punten = slide_data.get("content", [])
        for punt in inhoud_punten:
            p = tf_content.add_paragraph()
            p.text = punt
            p.font.size = Pt(20)
            p.level = 0

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- Knop om te genereren ---
if st.button("🎉 Genereer PowerPoint met AI"):
    if not opdracht.strip():
        st.warning("Vul eerst de opdracht of het onderwerp in.")
    else:
        with st.spinner("AI is bezig met genereren... dit kan enkele seconden duren"):
            slides = genereer_slides_via_ai(opdracht)
            if slides:
                pptx_file = maak_pptx(slides)
                st.success("✅ PowerPoint succesvol gegenereerd!")
                st.download_button(
                    label="📥 Download PowerPoint",
                    data=pptx_file,
                    file_name="praktijkopdracht_ai_presentatie.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            else:
                st.error("Kon geen slides genereren.")

# --- Optioneel: uitleg ---
st.markdown("---")
st.markdown(
    """
    **Uitleg:**  
    - Vul je OpenAI API-key in (te vinden op https://platform.openai.com/account/api-keys).  
    - Geef een duidelijke opdracht of onderwerp voor je praktijkopdracht.  
    - Klik op de knop om automatisch 20 dia's te laten genereren door AI.  
    - Download het PowerPoint-bestand en open het in PowerPoint of een compatible app.
    """
)

