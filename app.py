from pptx import Presentation
from pptx.util import Inches, Pt

# Jouw grote tekst, hier als voorbeeld
grote_tekst = """
Naam student: Jan Jansen
Studenten nummer: 123456
Naam project: Bouwproject X
Locatie project: Amsterdam
Leerbedrijf: Bouwbedrijf Y
Leermeester: Piet de Vries
Inleverdatum: 23-06-2025

Welke praktijkopdracht heb je gemaakt?
Ik heb een nieuwe aanbouw gemaakt aan een woning.
Waarom heb je deze praktijkopdracht gemaakt?
Om praktijkervaring op te doen met nieuwbouw.
Wat voor type werk was het? (nieuwbouw, aanbouw etc.)
Aanbouw.
Hoe was de werksituatie? (waren er bijzonderheden zoals tijdsdruk, samenwerking, het weer)
We hadden tijdsdruk maar goede samenwerking.
Hoe groot was je ploeg?
4 personen.

Beschrijf de risico’s bij deze praktijkopdracht (veiligheid, tijdsdruk, e.a.)
Risico: Valgevaar
Genomen maatregel: Gebruik van steigers en harnassen
Risico: Tijddruk
Genomen maatregel: Planning strak houden

Materiaalstaat
Materiaal: Hout
Maat: 2x4
Aantal: 50
Materiaal: Beton
Maat: 30kg zak
Aantal: 20

Gereedschapslijst
Gereedschap: Zaag
Gebruikt voor: Hout zagen
Gereedschap: Betonmixer
Gebruikt voor: Beton mengen

Werkschema en urenverantwoording
Stap: Lezen werktekening
Aantal uren: 2
Opmerkingen: Uitgebreid bestudeerd

Stap: Materiaal klaarzetten
Aantal uren: 1
Opmerkingen: Alles verzameld

Stap 1 – Lezen en begrijpen van de werktekening
Beschrijf hier wat je hebt gedaan.
Ik heb de werktekening zorgvuldig bestudeerd.
Waarom heb je het zo gedaan.
Om fouten te voorkomen.
Wat is was een leerpunt.
Nauwkeurig werken is cruciaal.
Instructies voor je collega (wat is belangrijk om op te letten?).
Let op de maten en details.
Let op!

Reflectie: persoonlijk
Hoeveel hulp had je nodig en wat kon je zelfstandig?
Ik kon veel zelfstandig.
Op welke momenten stuurde de Praktijkbegeleider/ collega je bij?
Bij de details.
Welke tips heb je gekregen?
Werk nauwkeuriger.
Welke leerpunten had je?
Betere communicatie.
Wat waren je sterke punten?
Doorzettingsvermogen.

Reflectie: Uitvoering en samenwerken
Wat werd er van je verwacht?
Zelfstandig werken.
Wat heb je daadwerkelijk zelfstandig gedaan?
Alle stappen.
Wat zou je de volgende keer anders doen?
Meer vragen stellen.
Wat zou je nog willen leren?
Betere planning.
Tips van je collega.
Blijf communiceren.
"""

# Maak een nieuwe presentatie
prs = Presentation()

# Helperfunctie om een dia met titel + tekst te maken
def add_slide(title, content_lines):
    slide_layout = prs.slide_layouts[1]  # Titel + inhoud
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    text_box = slide.shapes.placeholders[1].text_frame
    text_box.clear()
    for line in content_lines:
        p = text_box.add_paragraph()
        p.text = line
        p.font.size = Pt(14)

# Splits grote tekst in blokken op basis van double linebreaks (paragraphs)
paragrafen = [p.strip() for p in grote_tekst.split('\n\n') if p.strip()]

# Maak dia’s volgens jouw structuur (voorbeeld, aan te passen)
add_slide("Gegevens", paragrafen[0].split('\n'))
add_slide("Praktijkopdracht", paragrafen[1:6])
add_slide("Risico’s en maatregelen", paragrafen[6:8])
add_slide("Materiaalstaat", paragrafen[8:12])
add_slide("Gereedschapslijst", paragrafen[12:16])
add_slide("Werkschema en urenverantwoording", paragrafen[16:20])
add_slide("Stap 1 – Lezen en begrijpen van de werktekening", paragrafen[20:27])
add_slide("Reflectie: persoonlijk", paragrafen[27:34])
add_slide("Reflectie: Uitvoering en samenwerken", paragrafen[34:40])

# Sla het bestand op
prs.save("praktijkopdracht.pptx")
print("PowerPoint is gegenereerd als praktijkopdracht.pptx")

