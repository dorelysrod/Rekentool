#  Infraroodpanelen Rekentool

Een klant heeft zich gemeld bij **Praeter** met de vraag om een **rekentool** te ontwikkelen voor het plaatsen van infraroodpanelen in zijn volledige bedrijfsruimte. De berekening voor deze maatregel is eerder al uitgevoerd door een van onze Praeter-collega’s, maar op termijn is het natuurlijk efficiënter en gebruiksvriendelijker om deze berekening in een applicatie te kunnen invoeren.

---

##  Doel van de rekentool

Het advies dat we aan de klant geven, bestaat uit **twee onderdelen**, met een additionele verdieping.

### Opdracht 1 – Verbruiksvergelijking

We voeren een **vergelijking uit tussen het energieverbruik van de klant en het gemiddelde verbruik van vergelijkbare gebouwen elders in Nederland**. Dit geeft inzicht in:

- Hoeveel hoger of lager het huidige verbruik van de klant is ten opzichte van de sector.
- Of er al sprake is van een relatief energiezuinig gebouw, of dat hier nog substantiële besparingen mogelijk zijn.

---

### Opdracht 2 – Financiële berekening rentabiliteit

We voeren een **financiële berekening uit met betrekking tot de rentabiliteit van de panelen**. Hierin berekenen we onder andere:

- **IRR (Internal Rate of Return):** geeft inzicht in het totale rendement van de investering.
- **REV (Rendement op Eigen Vermogen):** toont het rendement op het door de klant zelf geïnvesteerde vermogen.
- **PAT (Profit After Tax):** geeft de winst na belasting weer.
- **TVT (Terugverdientijd):** laat zien in hoeveel jaar de investering is terugverdiend.

---

###  Additionele opdracht – Financiële inzichten over de looptijd

Naast de standaardberekeningen leveren we **extra financiële inzichten over de volledige looptijd van de investering**, zoals:

- **Cashflow-overzicht per jaar**, inclusief cumulatieve cashflow.
- **Gevoeligheidsanalyse van REV in relatie tot rente VV (vreemd vermogen)**: aangezien de rente op vreemd vermogen niet vaststaat, tonen we hoe veranderingen in het rentepercentage impact hebben op het rendement op eigen vermogen.

---

##  Projectstructuur

De repository heeft de volgende structuur:

project-root/
├── data/                 # (added to .gitignore)
│   └── PraeterBV_Case.xlsx
├── scripts/
├── outputs/
├── requirements.txt
└── .gitignore


##  Implementatie

Deze tool is geïmplementeerd in Python en bestaat uit:

- **opdracht1_sector_comparison.py**  
  Vergelijkt klantverbruik met CBS-data en genereert een visualisatie.

- **opdracht2_financial_calculation.py**  
  Voert financiële berekeningen uit en genereert inzichten en gevoeligheidsanalyses.

---

##  Installatie en gebruik 

### 1. Clone de repository

```bash
git clone 
cd 
```

### 2. Maak en activeer een virtual environment
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Installeer de vereisten
```bash
pip install -r requirements.txt
```
### 4. Controleer je bestanden

Zorg dat het bestand PraeterBV_Case.xlsx aanwezig is in de map data/ of pas het pad aan in de scripts indien nodig.

### 5. Run Opdracht 1

```bash
python3 sector_comparison.py
```
### 6. ### 5. Run Opdracht 2
```bash
python3 financial_calculation.py
```
### 7. Aditionel Opdracht

```bash
python3 export_cashflows.py.
```
---
## Opmerking

```bash
source venv/bin/activate
```
(Op Windows gebruik je: venv\Scripts\activate)
