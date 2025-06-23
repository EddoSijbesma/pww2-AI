import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from io import BytesIO

st.title("PowerPoint Generator zonder AI — 10 stappen")

st.write("Plak hieronder je grote tekst. De app maakt 10 stappen-dia’s plus overige dia’s.")

grote_tekst = st.text_area("Voer hier je tekst in:", height=600)

def set_slide_background(slide, rgb):
    """Zet een effen achtergrondkleur op de slide"""
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.shapes import MSO_FILL
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*rgb)

def add_colored_slide(prs, title, content_lines, bg_color=(91, 155, 213)):
    slide_layout = prs.slide_layouts[1]  # Titel + inhoud
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, bg_color)
    slide.shapes.title.text = title

    text_frame = slide.shapes.placeholders[1].text_frame
    text_frame.clear()
    for line in content_lines:
        p = text_frame.add_paragraph()
        p.text = line
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(255, 255, 255)  # wit tekst

    return slide

def generate_pptx(text):
    prs = Presentation()

    # Split op paragrafen (dubbele enters)
    paragrafen = [p.strip() for p in text.split('\n\n') if p.strip()]

    # Dia 1: Gegevens
    if len(paragrafen) > 0:
        add_colored_slide(prs, "Gegevens", paragrafen[0].split('\n'), bg_color=(31, 73, 125))

    # Dia 2: Praktijkopdracht
    if len(paragrafen) > 6:
        add_colored_slide(prs, "Praktijkopdracht", paragrafen[1:6], bg_color=(79, 129, 189))

    # Dia 3: Risico’s en maatregelen
    if len(paragrafen) > 8:
        add_colored_slide(prs, "Risico’s en maatregelen", paragrafen[6:8], bg_color=(91, 155, 213))

    # Dia 4: Materiaalstaat
    if len(paragrafen) > 12:
        add_colored_slide(prs, "Materiaalstaat", paragrafen[8:12], bg_color=(104, 161, 230))

    # Dia 5: Gereedschapslijst
    if len(paragrafen) > 16:
        add_colored_slide(prs, "Gereedschapslijst", paragrafen[12:16], bg_color=(117, 176, 247))

    # Dia 6: Werkschema en urenverantwoording
    if len(paragrafen) > 20:
        add_colored_slide(prs, "Werkschema en urenverantwoording", paragrafen[16:20], bg_color=(130, 190, 255))

    # Dia 7 t/m 16: Stap 1 t/m 10
    # Voor elke stap 1 t/m 10 een dia maken.
    # Voor elke stap verwachten we 7 paragrafen (beschrijving, waarom, leerpunt, instructies, let op!)
    # Dus totaal 7*10 = 70 paragrafen nodig; we doen zoveel mogelijk

    start_index = 20
    step_title_base = "Stap {} – Lezen en begrijpen van de werktekening"

    for i in range(1, 11):
        idx_start = start_index + (i - 1) * 7
        idx_end = idx_start + 7
        if len(paragrafen) >= idx_end:
            # Titel per stap, kan ook dynamisch per stap naam aangepast worden
            title = f"Stap {i} – Lezen en begrijpen van de werktekening"
            add_colored_slide(prs, title, paragrafen[idx_start:idx_end], bg_color=(31, 73, 125))
        else:
            break

    # Dia 17: Reflectie: persoonlijk
    ref_pers_index = start_index + 7 * 10
    if len(paragrafen) > ref_pers_index + 7:
        add_colored_slide(prs, "Reflectie: persoonlijk", paragrafen[ref_pers_index:ref_pers_index+7], bg_color=(79, 129, 189))

    # Dia 18: Reflectie: Uitvoering en samenwerken
    ref_uitv_index = ref_pers_index + 7
    if len(paragrafen) > ref_uitv_index + 5:
        add_colored_slide(prs, "Reflectie: Uitvoering en samenwerken", paragrafen[ref_uitv_index:ref_uitv_index+5], bg_color=(91, 155, 213))

    bio = BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

if grote_tekst:
    pptx_file = generate_pptx(grote_tekst)
    st.download_button(
        label="Download PowerPoint",
        data=pptx_file,
        file_name="praktijkopdracht.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
else:
    st.info("Plak een grote tekst om een PowerPoint te genereren.")


