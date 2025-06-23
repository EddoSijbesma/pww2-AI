import streamlit as st

st.title("Praktijkopdracht Instructieboek Generator")

st.header("📋 Gegevens")
naam = st.text_input("Naam student")
studentnummer = st.text_input("Studentnummer")
projectnaam = st.text_input("Naam project")
locatie = st.text_input("Locatie project")
leerbedrijf = st.text_input("Leerbedrijf")
leermeester = st.text_input("Leermeester")
inleverdatum = st.date_input("Inleverdatum")

st.file_uploader("Upload hier foto's van jezelf tijdens het werk", accept_multiple_files=True)

st.header("🛠️ Over de praktijkopdracht")
opdracht = st.text_area("Welke praktijkopdracht heb je gemaakt?")
wat_gemaakt = st.text_area("Wat heb je gemaakt?")
waarom_gemaakt = st.text_area("Waarom heb je deze praktijkopdracht gekozen?")
type_werk = st.selectbox("Wat voor type werk was het?", ["Nieuwbouw", "Aanbouw", "Renovatie", "Onderhoud", "Anders"])
werksituatie = st.text_area("Hoe was de werksituatie? (bijv. samenwerking, tijdsdruk, weer)")
ploeggrootte = st.text_input("Hoe groot was je ploeg?")

st.file_uploader("Upload hier foto’s van het eindresultaat", accept_multiple_files=True)

st.header("⚠️ Risico’s en maatregelen")
risicos = st.text_area("Beschrijf de risico’s bij deze praktijkopdracht")
maatregelen = st.text_area("Welke maatregelen heb je getroffen?")

st.header("📐 Werktekening")
st.file_uploader("Upload hier je werktekening", type=["jpg", "png", "pdf"])

st.header("🧰 Materiaal en gereedschap")
materialen = st.text_area("Materiaalstaat")
gereedschap = st.text_area("Gereedschapslijst")
werkuur = st.text_area("Werkschema en urenverantwoording")

# Dynamisch stappenplan
st.header("🪜 Stappenplan")
for i in range(1, 11):
    with st.expander(f"Stap {i}"):
        st.text_input(f"Stap {i} – Titel", key=f"stap{i}_titel")
        st.text_area(f"Wat heb je gedaan?", key=f"stap{i}_wat")
        st.text_area(f"Waarom heb je het zo gedaan?", key=f"stap{i}_waarom")
        st.text_area(f"Wat was een leerpunt?", key=f"stap{i}_leer")
        st.text_area(f"Instructies voor je collega", key=f"stap{i}_instructie")
        st.text_area(f"Let op!", key=f"stap{i}_letop")
        st.file_uploader("Voeg hier foto's toe", accept_multiple_files=True, key=f"stap{i}_foto")

st.header("🔍 Reflectie: Persoonlijk")
st.text_area("Hoeveel hulp had je nodig en wat kon je zelfstandig?")
st.text_area("Wanneer stuurde een collega je bij?")
st.text_area("Welke tips heb je gekregen?")
st.text_area("Wat waren je leerpunten?")
st.text_area("Wat waren je sterke punten?")

st.header("👥 Reflectie: Samenwerken")
st.text_area("Wat werd er van je verwacht?")
st.text_area("Wat heb je zelfstandig gedaan?")
st.text_area("Wat zou je de volgende keer anders doen?")
st.text_area("Wat wil je nog leren?")
st.text_area("Tips van je collega?")

st.header("📤 Afronding")
if st.button("Genereer concept instructieboek"):
    st.success("Concept gegenereerd! (functie om alles samen te voegen en te downloaden kan hier nog worden toegevoegd.)")
