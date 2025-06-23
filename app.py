import streamlit as st
import openai
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import requests

# Zet hier je oude OpenAI API-sleutel
openai.api_key = "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"

# Unsplash API key (vervang door jouw sleutel)
UNSPLASH_ACCESS_KEY = "jouw_unsplash_key"

st.title("AI PowerPoint Generator (oude OpenAI SDK)")

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
        sld = prs.slides.add_slide(prs.slide_layouts[5])  # lege layout
        title = sld.shapes.title
        if not title:
            title = sld.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1)).text_frame
        title.text = slide['title']
        
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9)
        height = Inches(3)
        
        # Voeg tekst toe
        txBox = sld.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.text = slide['content']
        
        # Voeg afbeelding toe als die er is
        if 'image' in slide:
            img_url = slide['image']
            try:
                img_data = requests.get(img_url).content
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
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[{"role": "user", "content": prompt}],
            )
            text = response['choices'][0]['message']['content']
            
            slides = []
            for line in text.split('\n'):
                if "Slide" in line and ":" in line:
                    parts = line.split(":", 1)[1].strip().split("-", 1)
                    if len(parts) == 2:
                        title = parts[0].strip()
                        content = parts[1].strip()
                        slides.append({"title": title, "content": content})
            
            # Zoek afbeeldingen via Unsplash
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
