import streamlit as st
import openai
import os
from utils import search_unsplash_image, create_styled_pptx, convert_pptx_to_pdf

# ⚠️ Hardcoded API key (NIET AANBEVOLEN voor productie)
openai.api_key = "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"
UNSPLASH_ACCESS_KEY = "VUL_HIER_JE_UNSPLASH_KEY_IN"

st.set_page_config(page_title="AI PowerPoint Generator", layout="centered")
st.title("🧠 AI PowerPoint Generator + Afbeeldingen + PDF")

topic = st.text_input("🎯 Onderwerp")
num_slides = st.slider("📄 Aantal slides", 3, 10, 5)
pdf_export = st.checkbox("📤 Exporteer ook naar PDF")

if st.button("✨ Genereer PowerPoint"):
    if not openai.api_key or not UNSPLASH_ACCESS_KEY:
        st.error("❌ OpenAI of Unsplash API key ontbreekt.")
    elif not topic.strip():
        st.error("❌ Vul een onderwerp in.")
    else:
        with st.spinner("🧠 GPT genereert inhoud..."):
            prompt = f"Maak een PowerPoint over '{topic}' met {num_slides} slides. Geef per slide een titel en inhoud. Format: Slide 1: Titel - Inhoud"
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[{"role": "user", "content": prompt}],
            )
            text = response['choices'][0]['message']['content']

        # Parse slides
        slides = []
        for line in text.split("\n"):
            if "Slide" in line and ":" in line:
                parts = line.split(":", 1)
                title_content = parts[1].strip().split("-", 1)
                if len(title_content) == 2:
                    title = title_content[0].strip()
                    content = title_content[1].strip()
                    slides.append({"title": title, "content": content})

        with st.spinner("📷 Haalt afbeeldingen op..."):
            for slide in slides:
                img = search_unsplash_image(slide["title"], UNSPLASH_ACCESS_KEY)
                if img:
                    slide["image"] = img

        pptx_io = create_styled_pptx(slides)

        st.success("✅ PowerPoint gegenereerd!")

        st.download_button(
            "📥 Download PowerPoint (.pptx)",
            data=pptx_io,
            file_name=f"{topic}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

        if pdf_export:
            try:
                with st.spinner("📄 Converteert naar PDF..."):
                    pdf_path = convert_pptx_to_pdf(pptx_io)
                    with open(pdf_path, "rb") as f:
                        st.download_button(
                            "📥 Download PDF",
                            data=f,
                            file_name=f"{topic}.pdf",
                            mime="application/pdf",
                        )
            except Exception as e:
                st.error(f"PDF export mislukt: {e}")
