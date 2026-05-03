# Vortrag-KI-VHS

Projektstand: 3. Mai 2026

Dieses Repository enthaelt die komplette Arbeitsbasis fuer einen ca. 60-minuetigen VHS-Vortrag zum Thema Kuenstliche Intelligenz: inhaltliche Konzeption, fertiges Folienskript, Copilot-Masterprompts, Bildprompts pro Folie sowie Python-Skripte zur (teil-)automatisierten PPTX-Erstellung.

## Ziel des Repository

- Eine verstaendliche, didaktisch starke KI-Einfuehrung fuer Erwachsene mit heterogenem Vorwissen.
- Produzierbare Praesentation mit konsistentem Bildstil und Sprechernotizen.
- Wiederverwendbarer Workflow fuer Aktualisierung und Re-Generierung der Folien.

## Repository-Inhalt im Ueberblick

### 1) Kerndokumente (Markdown)

- `01-Ki-Vortrag-Prompt.md`  
	Ausgangsauftrag fuer die Erstellung einer vollstaendigen, didaktisch strukturierten KI-Praesentation.

- `02-KI-Praesentation.md`  
	Hauptquelle der Praesentation (36 Folien + Anhang/Glossar), inkl.:
	- Folientitel und Bulletpoints
	- Sprechernotizen
	- Bildprompts je Kernfolie
	- Quellenhinweise

- `04_clean_pptx_prompt.md`  
	Bereinigte/aufbereitete Prompt-Version zur Umwandlung in eine Praesentation.

- `06_PPT_Copilot_Masterprompt_KI_RealisticStyle.md`  
	Masterprompt fuer Microsoft 365 Copilot in PowerPoint (realistischer, freundlicher Illustrationsstil).

### 2) Bildprompt-Sammlung

- Ordner: `Prompts_Bilder/`
- Enthalten:
	- `00_STIL_FUER_ALLE_BILDER.txt` (globaler Stilzusatz)
	- Einzeldateien je Folie (z. B. `10_Was_bedeutet_GPT.txt`, `29_Datenschutz.txt`)
- Muster: englischer Bildprompt + standardisierter Stilzusatz fuer visuelle Konsistenz.

### 3) Python-Generatoren

- `Python/generate_pptx.py`
	- Umfangreicher PPTX-Generator auf Basis von `python-pptx`.
	- Parst Markdown-Folienstruktur (Abschnitte, Tabellen, Bullets, Notizen, Bildprompts).
	- Enthält eigenes Designsystem (Farben, Layout-Logik, adaptive Schriftgroessen).
	- Bietet zusaetzlich eine Vorlagen-Variante (`main_vorlage`).

- `Python/generate_03_pptx.py`
	- Generator fuer template-basierten Vortragsexport.
	- Extrahiert und speichert Bildprompts als einzelne TXT-Dateien.
	- Legt zusaetzlich eine Stil-Datei fuer Bildgenerierung ab.

### 4) Praesentationsartefakte

- Root:
	- `2026-KI-Praesentation_gehalten.pptx`
	- `2026-KI-Praesentation_fuer_PDF_export.pptx`
	- `2026-KI-Praesentation_fuer_PDF_export.pdf`

- Ordner `Powerpoint/`:
	- `KI_VHS_60min_im_Stil_Vorlage01.pptx` (Template)
	- `KI_VHS_Vollversion_Aus_Markup 1.pptx`
	- `Layout_Demos_KI_VHS 1.pptx`

### 5) Sonstiges

- `Bilder/` sind die KI-generierten Bilder 

## Inhaltliche Struktur der Praesentation

Die Praesentation besteht aus **36 Folien + Glossar** und ist in 11 Teile gegliedert.
Quelldatei: `02-KI-Praesentation.md`

---

**Titelfolie & Ueberblick**
- Folie 1 – Kuenstliche Intelligenz *(Titelfolie)*
- Folie 2 – Was erwartet Sie heute?

---

**Teil 1 – Einleitung**
- Folie 3 – Ein Moment, der alles veraendert hat *(ChatGPT-Launch Nov. 2022, 1 Mio. User in 5 Tagen)*
- Folie 4 – Warum jetzt? Die drei Schluessel *(Daten · Rechenpower · Transformer-Algorithmus)*

**Teil 2 – Geschichte der KI**
- Folie 5 – Die Geburt einer Idee (1950er) *(Alan Turing, Dartmouth-Konferenz 1956)*
- Folie 6 – KI-Winter und KI-Fruehling *(Zeittafel 1950–2022)*

**Teil 3 – Machine Learning**
- Folie 7 – Was ist maschinelles Lernen? *(Regeln vs. Muster aus Beispielen)*
- Folie 8 – Lernen aus Beispielen *(Spam-Filter-Analogie)*
- Folie 9 – Vergleich: Wie ein Kind lernt *(Gegenuebertstellung Kind / KI-Modell)*

**Teil 4 – GPT in die Tiefe**
- Folie 10 – Was bedeutet GPT? *(Generative · Pre-trained · Transformer · LLM-Begriff)*
- Folie 11 – Was sind Token? *(Wortfragmente, Kontextfenster, Live-Demo-Idee)*
- Folie 12 – Das Geheimnis: Wahrscheinlichkeiten *(naechstes Wort, Kaffee/Tee-Analogie)*
- Folie 13 – Transformer: Die Aufmerksamkeit *(Attention-Mechanismus, „Bank am Fluss")*
- Folie 14 – Kein Denken. Kein Verstehen. *(GPT berechnet, mustert – hat kein Bewusstsein)*

**Teil 5 – Warum GPT trotzdem intelligent wirkt**
- Folie 15 – Sprache ist unsere Superkraft *(Sprache als Traeger von Intelligenz-Wirkung)*
- Folie 16 – Muster, Stil, Tonalitaet *(GPT imitiert Schreibstile)*

**Teil 6 – Prompting**
- Folie 17 – Was ist ein Prompt? *(Restaurantbestellung-Analogie)*
- Folie 18 – Best Practice: Die 5 Zutaten *(Rolle · Kontext · Format · Iterieren · Schritt fuer Schritt)*
- Folie 19 – Schlechter vs. guter Prompt *(Vergleichstabelle mit Beispielen)*
- Folie 20 – Fortgeschrittene Techniken *(Chain of Thought · Chain of Draft · Umgekehrte Befragung · Negatives Prompting)*

**Teil 7 – Was KI gut kann**
- Folie 21 – Halluzinationen: Ein reales Warnsignal *(Anwalts-Beispiel USA)*
- Folie 22 – Halluzinationen: Warum passiert das? *(Wahrscheinlichkeit statt Wahrheitspruefung)*
- Folie 23 – Halluzinationen: So schuetzen Sie sich *(5 konkrete Regeln)*
- Folie 24 – Bias: Die KI spiegelt uns *(Vorurteile in Trainingsdaten, Amazon-Beispiel)*
- Folie 25 – LLM vs. RAG *(statisches Wissen vs. Live-Suche; Perplexity, Copilot)*
- Folie 26 – Die Staerken der KI *(Schreiben · Uebersetzen · Brainstormen · Strukturieren)*
- Folie 27 – KI in der Praxis: Durchbrueche & Warnsignale *(AlphaFold-Nobelpreis · Krebserkennung · Amazon-Bias · Gesichtserkennung)*

**Teil 8 – Was KI nicht kann**
- Folie 28 – Die Grenzen der KI *(kein echtes Verstehen, keine Empathie, keine Verantwortung)*

**Teil 9 – Datenschutz & Sicherheit**
- Folie 29 – Datenschutz *(Faustregel: KI wie eine Postkarte behandeln)*

**Teil 10 – In den Kinderschuhen**
- Folie 30 – Wie jung ist KI wirklich? *(Vergleich: Dampfmaschine · Auto · Computer · Internet)*
- Folie 31 – Wohin geht die Reise? *(Multimodalitaet · Agenten · Personalisierung · Wissenschaft)*
- Folie 32 – Reasoning-Modelle: KI, die erst nachdenkt *(OpenAI o3, DeepSeek R1 – Denkpause vor Antwort)*

**Teil 11 – Fazit**
- Folie 33 – Was wir heute gelernt haben
- Folie 34 – Ihr erster Schritt *(konkrete Handlungsaufforderung)*
- Folie 35 – Empfohlene Einstiegs-Tools *(ChatGPT · Copilot · Gemini · Claude · Perplexity)*
- Folie 36 – Schlussbild *(„KI ist nicht die Zukunft. KI ist die Gegenwart.")*

---

**Anhang: Glossar** – 12 Begriffe: KI, Machine Learning, GPT, Token, Halluzination, Bias, Prompt, Transformer, LLM, RAG, Reasoning-Modell, DSGVO.

## Empfohlener Arbeitsablauf

### A) Inhalt pflegen

1. `02-KI-Praesentation.md` als inhaltliche Wahrheit nutzen und aktualisieren.
2. Optional Copilot-Workflows mit `04_clean_pptx_prompt.md` und `06_PPT_Copilot_Masterprompt_KI_RealisticStyle.md` verwenden.

### B) Bilder erzeugen

1. Prompts aus `Prompts_Bilder/` verwenden.
2. Stilzusatz aus `Prompts_Bilder/00_STIL_FUER_ALLE_BILDER.txt` konsequent anhaengen.
3. Bilder in die jeweiligen Folien einsetzen.

### C) PPTX erzeugen (Python)

Voraussetzung:

```bash
python3 -m pip install python-pptx
```

Generator starten (vom Projektroot):

```bash
python3 Python/generate_pptx.py
```

Optionaler Template-Generator:

```bash
python3 Python/generate_03_pptx.py
```

## Wichtige technische Hinweise

- Die Python-Skripte arbeiten mit festen Dateinamen und erwarten je nach Variante bestimmte Eingabedateien relativ zum Skriptpfad.
- Wenn Eingabedateien verschoben wurden, muessen Pfade in den Skripten angepasst werden.
- Fuer reproduzierbare Ergebnisse empfiehlt sich eine konsistente Reihenfolge:
	1. Markdown finalisieren
	2. Bildprompts finalisieren
	3. PPTX generieren
	4. Manuelle Design-Feinkorrektur in PowerPoint

## Dateistatistik (Textquellen)

- Markdown gesamt (4 Hauptdateien): 1.813 Zeilen
- Python-Skripte: 1.483 Zeilen
- Bildprompt-Dateien: 29 Dateien (inkl. globaler Stildatei)

## Lizenz

[![CC BY 4.0](https://licensebuttons.net/l/by/4.0/88x31.png)](https://creativecommons.org/licenses/by/4.0/deed.de)

Copyright (c) 2026 Hagen Gloetter – lizenziert unter [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/deed.de).

Nutzung, Weitergabe und Bearbeitung sind frei erlaubt – nicht kommerziell –, solange der Urheber genannt wird.
Beispiel: „Basierend auf Material von Hagen Gloetter (2026), CC BY 4.0"

Details: [LICENSE](LICENSE)
