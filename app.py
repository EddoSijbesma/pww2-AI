from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

def style_text_frame(tf, font_name="Arial", font_size=Pt(24), bold=False, italic=False, color=RGBColor(0, 0, 0), align=PP_ALIGN.LEFT):
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    font = run.font
    font.name = font_name
    font.size = font_size
    font.bold = bold
    font.italic = italic
    font.color.rgb = color
    return p

def maak_pptx(first_slide_data, slides):
    prs = Presentation()

    slide_layout = prs.slide_layouts[5]  # lege slide
    sld = prs.slides.add_slide(slide_layout)

    # Titel "Stappenplan praktijkopdracht"
    title_box = sld.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1.5))
    title_tf = title_box.text_frame
    p = title_tf.add_paragraph()
    p.text = "Stappenplan praktijkopdracht"
    p.alignment = PP_ALIGN.CENTER
    font = p.runs[0].font
    font.name = "Arial"
    font.size = Pt(40)
    font.bold = True
    font.color.rgb = RGBColor(0, 51, 102)  # donkerblauw

    # Subtitel praktijkopdracht titel (cursief en lichtgrijs)
    subtitle_box = sld.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(9), Inches(0.7))
    subtitle_tf = subtitle_box.text_frame
    p_sub = subtitle_tf.add_paragraph()
    p_sub.text = first_slide_data['praktijkopdracht_titel'] or "Titel van de praktijkopdracht"
    p_sub.alignment = PP_ALIGN.CENTER
    font_sub = p_sub.runs[0].font
    font_sub.name = "Arial"
    font_sub.size = Pt(28)
    font_sub.italic = True
    font_sub.color.rgb = RGBColor(150, 150, 150)

    # Vaste gegevens netjes uitgelijnd links, kleinere lettergrootte
    inhoud_box = sld.shapes.add_textbox(Inches(1), Inches(2.7), Inches(8), Inches(3))
    inhoud_tf = inhoud_box.text_frame
    inhoud_tf.word_wrap = True

    inhoudsregels = [
        f"Naam student: {first_slide_data['student_naam']}",
        f"Studenten nummer: {first_slide_data['student_nummer']}",
        f"Naam project: {first_slide_data['project_naam']}",
        f"Locatie project: {first_slide_data['project_locatie']}",
        f"Leerbedrijf: {first_slide_data['leerbedrijf']}",
        f"Leermeester: {first_slide_data['leermeester']}",
        f"Inleverdatum: {first_slide_data['inleverdatum']}"
    ]

    for regel in inhoudsregels:
        p = inhoud_tf.add_paragraph()
        p.text = regel
        p.font.size = Pt(18)
        p.font.name = "Calibri"
        p.font.color.rgb = RGBColor(50, 50, 50)
        p.space_after = Pt(5)
        p.alignment = PP_ALIGN.LEFT

    # De rest van de slides zoals eerder
    for slide in slides:
        sld = prs.slides.add_slide(prs.slide_layouts[5])
        # Titel slide
        title_box = sld.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
        title_tf = title_box.text_frame
        p = title_tf.add_paragraph()
        p.text = slide['title'] or ""
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.name = "Arial"
        p.alignment = PP_ALIGN.LEFT

        # Inhoud slide
        content_box = sld.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(6), Inches(4))
        content_tf = content_box.text_frame
        content_tf.word_wrap = True
        p_content = content_tf.add_paragraph()
        p_content.text = slide['content'] or ""
        p_content.font.size = Pt(20)
        p_content.font.name = "Calibri"

        # Afbeelding toevoegen indien aanwezig
        if slide['image'] is not None:
            try:
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

