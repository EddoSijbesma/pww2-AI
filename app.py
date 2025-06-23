import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
from PIL import Image
import tempfile

st.title("PowerPoint generator met tot 25 dia's + afbeeldingen")

NUM_SLIDES = 25

slides_content = []
for i in range(NUM_SLIDES):
    st.header(f"Slide {i+1}")
    title = st.text_input(f"Slide {i+1} titel", key=f"title_{i}")
    content = st.text_area(f"Slide {i+1} inhoud", key=f"content_{i}")
    # Upload afbeelding of plakken (streamlit laat uploaden toe; plakken via ctrl+v werkt in upload widget)
    img = st.file_uploader(f"Upload afbeelding voor slide {i+1} (optioneel)", type=["png","jpg","jpeg"], key=f"img_{i}")
    slides_content.append({"title": title, "content": content, "image": img})

def maak_pptx(slides):
    prs = Presentation()
    for slide in slides:
        sld = prs.slides.add_slide(prs.slide_layouts[5])  # lege slide
        if sld.shapes.title:
            sld.shapes.title.text = slide['title'] or ""
        else:
            textbox = sld.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
            textbox.text_frame.text = slide['title'] or ""
        txBox = sld.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(6), Inches(3))
        txBox.text_frame.text = slide['content'] or ""

        if slide['image'] is not None:
            try:
                # image = slide['image'].read() # bytes
                # PowerPoint vereist een bestand of bytes, we kunnen het tijdelijk opslaan
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
    # Controleer of tenminste 1 slide titel en inhoud heeft (of je wil stricte check per slide)
    if all(not s['title'] and not s['content'] for s in slides_content):
        st.error("Vul tenminste voor één slide een titel of inhoud in.")
    else:
        pptx_file = maak_pptx(slides_content)
        st.success("PowerPoint is klaar!")
        st.download_button(
            label="Download PowerPoint (.pptx)",
            data=pptx_file,
            file_name="mijn_presentatie.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

