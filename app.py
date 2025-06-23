import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO

st.title("PowerPoint generator zonder API")

num_slides = st.slider("Aantal dia's", 1, 20, 5)

slides_content = []
for i in range(num_slides):
    st.header(f"Slide {i+1}")
    title = st.text_input(f"Slide {i+1} titel", key=f"title_{i}")
    content = st.text_area(f"Slide {i+1} inhoud", key=f"content_{i}")
    slides_content.append({"title": title, "content": content})

def maak_pptx(slides):
    prs = Presentation()
    for slide in slides:
        sld = prs.slides.add_slide(prs.slide_layouts[5])  # lege slide
        if sld.shapes.title:
            sld.shapes.title.text = slide['title']
        else:
            textbox = sld.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
            textbox.text_frame.text = slide['title']
        txBox = sld.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(3))
        txBox.text_frame.text = slide['content']
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

if st.button("Genereer PowerPoint"):
    if any(not s['title'] or not s['content'] for s in slides_content):
        st.error("Vul voor elke slide titel en inhoud in.")
    else:
        pptx_file = maak_pptx(slides_content)
        st.success("PowerPoint is klaar!")
        st.download_button(
            label="Download PowerPoint (.pptx)",
            data=pptx_file,
            file_name="mijn_presentatie.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
