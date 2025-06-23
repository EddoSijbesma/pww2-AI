import streamlit as st
from pptx import Presentation
from io import BytesIO
import datetime

# Functie om placeholders in tekst te vervangen
def vervang_tekst(tekst, vervangingen):
    for sleutel, waarde in vervangingen.items():
        tekst = tekst.replace(f"{{{{{sleutel}}}}}", str(waarde))
    return tekst

# Functie om een presentatie te genereren op basis van sjabloon
def genereer_powerpoint(vervangingen, template_path="Verslag Praktijkopdracht stappenplan variant (003).pptx"):
    prs = Presentation(template_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                originele_tekst = shape.text
                nieuwe_tekst = vervang_tekst(originele_tekst, vervangingen)
                if originele_tekst != nieuwe_tekst:
                    shape.text = nieuwe_tekst

    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# --- Streamlit app interface ---
st.title("🛠️ Praktijkopdracht PowerPoint Generator")

st.markdown("Vul hieronder de gegevens in. De PowerPoint wordt automatisch ingevuld op basis van jouw sjabloon.")

# Gegevens invullen
naam = st.text_input("Naam student")
studentnummer = st.text_input("Studentnummer")
project = st.text_input("Projectnaam")
locatie = st.text_input("Locatie project")
leerbedrijf = st.text_input("Leerbedrijf")
leermeester = st.text_input("Leermeester")
inleverdatum = st.date_input("Inleverdatum", datetime.date.today())

# Praktijkopdracht info
opdracht = st.text_area("Welke praktijkopdracht heb je gemaakt?")
wat_gemaakt = st.text_area("Wat heb je gemaakt?")
waarom = st.text_area("Waarom heb je deze praktijkopdracht gekozen?")
type_werk = st.text_input("Wat voor type werk was het?")
werksituatie = st.text_area("Hoe was de werksituatie?")
ploeggrootte = st.text_input("Hoe groot was je ploeg?")

# Risico’s en maatregelen
risicos = st.text_area("Beschrijf de risico’s")
maatregelen = st.text_area("Welke maatregelen heb je genomen?")

# Verzamelen van alle placeholders en bijbehorende waarden
vervangingen = {
    "naam": naam,
    "studentnummer": studentnummer,
    "project": project,
    "locatie": locatie,
    "leerbedrijf": leerbedrijf,
    "leermeester": leermeester,
    "inleverdatum": inleverdatum.strftime('%d-%m-%Y'),
    "opdracht": opdracht,
    "wat_gemaakt": wat_gemaakt,
    "waarom": waarom,
    "type_werk": type_werk,
    "werksituatie": werksituatie,
    "ploeggrootte": ploeggrootte,
    "risicos": risicos,
    "maatregelen": maatregelen
}

# Downloadknop
st.subheader("⬇️ Download jouw ingevulde PowerPoint")

if st.button("📤 Genereer PowerPoint"):
    try:
        pptx_bestand = genereer_powerpoint(vervangingen)
        st.download_button(
            label="🎓 Download PowerPoint",
            data=pptx_bestand,
            file_name="praktijkopdracht_ingevuld.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except Exception as e:
        st.error(f"Er is een fout opgetreden bij het genereren van de PowerPoint: {e}")

