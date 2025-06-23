import streamlit as st
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import requests

# Zet hier je keys neer (pas aan naar jouw echte keys)
OPENAI_API_KEY = "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"
UNSPLASH_ACCESS_KEY = "YOUR_UNSPLASH_ACCESS_KEY"

client = OpenAI(api_key=OPENAI_API_KEY)

st.title("AI PowerPoint Generator met OpenAI & Unsplash")

topic = st.text_input("Onderwerp van je presentatie:")
num_slides = st.slider("Aantal dia's", min_value=5, max_value=20, value=12)

def zoek_unsplash_afbeelding(query):
    url = f"https://api.unsplash.com/photos/random?query={query}&client_id={UNSPLASH_ACCESS_KEY}"
    try:
        resp = requests.get(url)
        resp.raise_for_status()
        data = resp.json()
        return data.get('urls', {}).get('small')
    except:
        return None

def maak_pptx(slides):
    prs = Presentation()
    for slide in slides:
        sld = prs.slides.add_slide(prs.slide_layouts[5])
        if sld.shapes.title:
            sld.shapes.title.text = slide['title']
        else:
            textbox = sld.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
            textbox.text_frame.text = slide['title']
        txBox = sld.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(3))
        txBox.text_frame.text = slide['content']
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
    if not topic.strip():
        st.error("Vul een onderwerp in aub.")
    else:
        with st.spinner("AI is aan het werk..."):
            prompt = (
                f"Maak een PowerPoint presentatie over '{topic}' met {num_slides} slides. "
                "Voor elke slide geef je een titel en een korte uitleg in dit format:\n"
                "Slide 1: Titel - Inhoud\n"
                "Slide 2: Titel - Inhoud\n"
                "Ga zo door tot alle slides beschreven zijn."
            )
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

