import streamlit as st
from openai import OpenAI
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import requests

# Gebruik hier jouw API key
client = OpenAI(api_key="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")

UNSPLASH_ACCESS_KEY = "jouw_unsplash_key"

st.title("AI PowerPoint Generator")

topic = st.text_input("Onderwerp")
num_slides = st.slider("Aantal dia's", 3, 20, 5)

def zoek_unsplash_afbeelding(query):
    url = f"https://api.unsplash.com/photos/random?query={query}&client_id={UNSPLASH_ACCESS_KEY}"
    resp = requests.get(url)
    if resp.status_code == 200:
        data = resp.json()
        return data['urls']['small']
    return None

def maak_pptx(slides):
    prs = Presentation()
    for slide in slides:
        sld = prs.slides.add_slide(prs.slide_layouts[5])
        title_shape = sld.shapes.title
        if title_shape:
            title_shape.text = slide['title']
        else:
            textbox = sld.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
            textbox.text_frame.text = slide['title']
        txBox = sld.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(3))
        tf = txBox.text_frame
        tf.text = slide['content']

        if 'image' in slide:
            try:
                img_data = requests.get(slide['image']).content
                img_stream = BytesIO(img_data)
                sld.shapes.add_picture(img_stream, Inches(7), Inches(1.5), width=Inches(2), height=Inches(2))
            except:
                pass

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

if st.button("Genereer PowerPoint"):
    if not topic:
        st.error("Vul een onderwerp in.")
    else:
        prompt = f"Maak een PowerPoint-presentatie over '{topic}' met {num_slides} slides. Geef voor elke slide een titel en korte uitleg in het formaat: Slide 1: Titel - Inhoud"
        try:
            response = client.chat.completions.create(
                model="gpt-4",
                messages=[{"role": "user", "content": prompt}],
            )
            text = response.choices[0].message.content

            slides = []
            for line in text.split('\n'):
                if "Slide" in line and ":" in line:
                    parts = line.split(":", 1)[1].strip().split("-", 1)
                    if len(parts) == 2:
                        title = parts[0].strip()
                        content = parts[1].strip()
                        slides.append({"title": title, "content": content})

            for slide in slides:
                img_url = zoek_unsplash_afbeelding(slide['title'])
                if img_url:
                    slide['image'] = img_url

            pptx_file = maak_pptx(slides)
            st.success("Presentatie is klaar!")
            st.download_button(
                label="Download PowerPoint (.pptx)",
                data=pptx_file,
                file_name=f"{topic}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"Fout bij genereren: {e}")
