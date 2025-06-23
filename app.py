import streamlit as st
from openai import OpenAI
import os
from utils import search_unsplash_image, create_styled_pptx, convert_pptx_to_pdf

# Zet hier je API-sleutels
OPENAI_API_KEY = "sk-..."  # Vervang met jouw OpenAI sleutel
UNSPLASH_ACCESS_KEY = "unsplash-..."  # Vervang met jouw Unsplash API key

# Initialiseer OpenAI client
client = OpenAI(api_key=OPENAI_API_KEY)

st.set_page_config(page_title="AI PowerPoint Generator", layout="centered")
st.title("🎓 AI PowerPoint Generator met Afbeeldingen en PDF")

# Invoer
topic = st.text_input("🧠 Onderwerp")
num_slides = st.slider("📄 Aantal dia's", min_value=3, max_value=20, value=5)
export_pdf = st.checkbox("📤 Exporteer ook naar PDF")

# Startknop
if st.button("🚀 Genereer PowerPoint"):
    if not OPENAI_API_KEY or not UNSPLASH_ACCESS_KEY:
        st.error("❌ API-sleutels ontbreken.")
    elif not topic.strip():
        st.error("❌ Vul een onderwerp in.")
    else:
        with st.spinner("💡 Genereert dia-inhoud..."):
            prompt = (
                f"Maak een PowerPoint-presentatie over '{topic}' met {num_slides} slides. "
                f"Geef elke slide een duidelijke titel en een korte uitleg of bulletpoints. "
                f"Format: Slide 1: Titel - Inhoud"
            )

            response = client.chat.completions.create(
                model="gpt-4",
                messages=[{"role": "user", "content": prompt}],
            )
            text = response.choices[0].message.content

        # Parse dia's
        slides = []
        for line in text.split("\n"):
            if "Slide" in line and ":" in line:
                parts = line.split(":", 1)
                content_parts = parts[1].split("-", 1)
                if len(content_parts) == 2:
                    title = content_parts[0].strip()
                    content = content_parts[1].strip()
                    slides.append({"title": title, "content": content})

        # Afbeeldingen toevoegen
        with st.spinner("🖼️ Zoekt afbeeldingen..."):
            for slide in slides:
                image = search_unsplash_image(slide["title"], UNSPLASH_ACCESS_KEY)
                if image:
                    slide["image"] = image

        # PowerPoint aanmaken
        pptx_io = create_styled_pptx(slides)

        # Downloadknop
        st.success("✅ Presentatie klaar!")
        st.download_button(
            label="📥 Download PowerPoint (.pptx)",
            data=pptx_io,
            file_name=f"{topic}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

        # PDF (optioneel)
        if export_pdf:
            try:
                with st.spinner("🔄 Converteert naar PDF..."):
                    pdf_path = convert_pptx_to_pdf(pptx_io)
                    with open(pdf_path, "rb") as f:
                        st.download_button(
                            label="📥 Download PDF",
                            data=f,
                            file_name=f"{topic}.pdf",
                            mime="application/pdf",
                        )
            except Exception as e:
                st.error(f"❌ PDF-generatie mislukt: {e}")

