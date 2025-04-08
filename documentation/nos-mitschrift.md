## NOS 💚 Deadlines ☠️ and other 👻 important things:
- **❄️ Frozen Zone (NOS) ❄️** = **mindestens 13T vor Auslieferdatum** => **14. Tag vor Auslieferdatum => Schieben möglich**
- **📦 Bei Abgabe der Produktions_KW => 1 Woche hinter $Hersteller_H => gleich mit $Hersteller_B (Bsp.: KW_11 für KW_14)**
- **📦 Produktions_Datum für Auftrag = im Idealfall -3 Werktage vor Ausliefer_Datum**
- **📦 Bereitstell_Datum +2 Werktag zum Produktions_Datum** => Eg. **-1 Werktage vor Ausliefer_Datum** => **Bereitstelldatum_Uhrzeit** = (06:00 Uhr)
- **📗 Avisieren** immer **täglich** für **für alle fertigen Touren** => **normalerweise ca. 3 Folgewochen sind Avisiert.**
- **♻️ Import der Cargo-Support Datensätze** = **[Täglich um 06:00 & 10:00 & 14:00 Uhr]**
- **🛢️ WAB'S** => **bei 20-25m³ voll** => meist **nicht auf 30m³ füllbar**
- **📅 Kalender-Woche** bei **>= 10m³** = **NICHT FAHREN** => **ist mit Vertreib abklärt ? Peter fragen? - Gebiete ?**
- **📅 Unterschiedliche Auslifer-Kalenderwochen für Gebiete** => **NEIN** => **Alles (Mo - Fr)**
- **📦 Produktions_Datum manchmal unterschiedlich => DONE ALEX**
- **🔨 Einlastungsdatum:** => **-3 Werktage** == **📦 Prod_Datum**  => Liste in Teams mit Alex/Helmuth in Teams ablegen für WAB-Nr. 
- **⏰ WAB_Bereitstell_Uhrzeit: 06:00 Uhr pro Gebiet** | **Tour ab 08:00**
- **📑 Frachtbriefe und Auslieferliste für SC-Leiter** => Helmut macht Listen mit **Ladelisten + Tourenplan + Deckblatt**
    - **Wöchtenlich** => Immer **Mittwoch for Folge_KW2**
    - Neue Lösung für **Frachtbriefe & Tourenpläne** für SC-Leiter und Hersteller **[Siehe Unten](#)**
- **🚛 Fuhrpark => Fahrzeuge und Speditionen pro Gebiet**
- **📦 Produktions_Export_Datum** und **[WOCHENTAG]** => Nicht vorhanden abarbeiten der Listen! ca. 4_Folge_KW's
- **Postversand / DIREKT-Versand Termine pro Gebiete** => Immer **DONNERSTAG in der KW** **(keine Fixtermine)**
- **Postversand => Do. in der KW (ALLE GEBIETE) => Alle in eine Tour!** => Prod = Montag
- **↗️ 4 Rampen vorhanden die gleichzeit LKW beladen können (Neudörfel)**
- **⚠️ Fixtermine** => **Alle zum einhalten**
- **🔫 Schnellschuss** ab wann ? Eigenheiten der **📦 Produktion** => **🔫 Schnellschuss NUR bei Reklas & Eilaufträge**
    - Besprechen mit **Auftragsabwickling** => **Datum in CS beachten**
- **Mokledruck (Etiketten) => Irrelevant aufgrund => Produktion verlagt es nicht**
- **📗 Aviso-Datum => auf 3h Zeit-Fenster setzten!** => **MÖGLICH** => Siehe **Test-Mail**
- **🛑 Not Tested but told not working:** => **Kann 2020 geteilte WAB-Aufträge übernehmen ?**

## 📦 Einlastung der Produktion (Status 200) => Liefervorschlag Druck => Liefertermin festgelget CS
- **Gewählter Liefer-Termin_System (Status 200) => max. -1T vor Liefertermin**
- **Liefervorschläge werden anschließend Ausgedrückt => täglich bis max. 11:00Uhr**
- **Reklas kommen rein => System schlägt nächstmöglich vor => händisch vorgerückt**
- **Rückstände werden von Alex notiert. = Keine Möglichkeit im 2020**
- **Vergleich NOS_Tourenkonzept & Kapazitätreports => Entscheidung für neuen Liefertermin**
- **Reklas wird nach Dispo gerichtet.**
- **2020 gibt selbe Auftragsbestätigung wie Cargo-Support** => **Nachbessern im System => Falscheingabe des Verkäfers.**
- **Alex retuniert => Mails an Verkäufer.**
- **Täglich immer bis spät. 11:00 Uhr**
- **Täglich anschließend dann druck der Ladelisten** => **Status Ändert sich nicht.**
- **Liefervorschläge Deckblätter** nach Touren sortiert.
- **Deckblatt vonm Heinz == N=NUR Fahrer / J=MIT MONTEUR**
- **Heinz-Fach => Alles nicht vermekrte Rückstände oder auf Plant SC-Leiter**
- meistens **Dienstag & Mittwochs** => wird **Folge_KW_2** geschlossen!
- Dann wird werden **Liefervorschläge** gedruckt.
- Täglich um **ca. 11:00 Uhr werden Nachträge** gedruckt.

## Touren-Namens Schema

| NAME | SCHEME | INFO |
| --- | --- | --- |
| **ABH Kundenamen TT.MM** | **ABH $KUNDE $ORT $WT(TT)** | Selbstabholungs-Tour am Donnerstag |
| **Postversand TT.MM** | **POSTVERSAND $ORT $KW** | Alle Postversand ohne Fixtermin in KW_12 |
| **Kundenname KundenOrt TT.MM** | **$KUNDE $ORT $WT(TT)** | Direkt-LKW Tour zur Baustelle am Donnerstag |
| **SC INNSBRUCK WAB TT.MM** | **SC WAB für $Fahrzeuge Touren $WT(TT)** | Lager-WAB Mittwoch <br> für LKW & 7,5t & 3,5t Touren <br> "RAMPE" in "Bemerkung Transport" |
| **SC 7,5t Tour Innsbruck - Telfs TT.MM** | **SC $Fahrzeuge Tour $ORT_S $ORT_E $WT(TT)** | Untertour 7,5t Fahrzeug <br> Laden von Lager-WAB |
| **SC LKW Tour Fügen - Jenbach TT.MM** | **SC $Fahrzeuge Tour $ORT_S $ORT_E $WT(TT)**  | Untertour LKW Fahrzeug <br> Laden von Lager-WAB|
| **Plant SC Leiter / 6 Kunden** |  | Unter-Tour mit 6 Kunden <br> die SC-Leiter Plant |
| **SC WAB für ULGB MI** | **SC WAB für ULGB $WT(TT)** | **[⛰️ ULGB Voradelberg ist Anders](%EF%B8%8F-vorarlberg-ist-anders)** <br> Lager-WAB Mittwoch <br> Voradelberg **immer "Platz"** |
| **ULGB Feldkirch - Dornbirn FR** | **ULGB $ORT_S $ORT_E $WT(TT)** | **[⛰️ ULGB Voradelberg ist Anders](%EF%B8%8F-vorarlberg-ist-anders)** <br> Tour am Freitag (lt. ULGB) |



## 🛢️ Interner-WAB-Nummern-Kreis & Schema
- **WAB-Nummer = OO_W_AA**
- **OO** = **ORTSKENNZAHL** **[Siehe Auftragsnummern-Tabelle]()**
- **W** = **WOCHENTAG = 1-5 = Mo-Fr**
- **AA** = ANZAHL DER CONT PRO TAG **01 - 99**
- **Beispiele Innsbruck:**
    - WAB(1) Innsbruck Montags = **15_1_01**
    - WAB(2) Innsbruck Montags = **15_1_02**
    - WAB(3) Innsbruck Montags = **15_1_03**
    - WAB(1) Graz Dienstags = **17_2_01**
    - WAB(1) Graz Mittwochs = **17_3_01**
    - WAB(1) Graz Donnerstags = **17_4_01**
    - WAB(1) Graz Freitags = **17_5_01**

## Interaktive Ausliefer-Karte und Postleitzahlen Zuordnung

**[Klicke Hier](https://kingslayer9988.github.io/austria-post-and-area-code/) für Interaktive Karte**

**[![Interactive Dispatcher Map](https://github.com/BMLZellEr/bgo_montage_und_logistik/raw/main/austria_cargo_zone_map/Screenshot_Map.png)](https://bmlzeller.github.io/bgo_montage_und_logistik/austria_cargo_zone_map/index.html)**



| Zone | SC | PLZ |
| --- | --- | --- |
| **A B C E** | **W.Neudorf Roman** |  |
| **H D** | **Linz Öhlinger** | |
| **F** | **Graz Koeck** |  |
| **G** | **Klagenfurt Koeck** |  |
| **I** | **Innsbruck Thonhauser** |  |
| **J** | **Dornbirn ULGB** |  |

## 🚚 Transportbestellung
- Macht aktuell immer Helmuth => Im Team abklären wie es weitergeht
- Vorlage:
````
Sehr geehrte Damen und Herren! [Oder Persöhnliche Anrede]

Abholung von [$WAB_ANZAHL] WAB am [$DATUM_ABH] ab [$UHRZEIT_ABH] Uhr ([$LADELISTE] + [$LADELISTE] + [$LADELISTE])
Zustellung am [$DATUM_ZUS] um [$UHRZEIT_ZUS] Uhr an [$KUNDENNAME] [$KUNDENADRESSE]

Danke im Voraus.
````

## ♻️ Import der Aufträge / Datensätze
- **Alex sendet E-Mails an Helmuth mit Änderungen für Aufträge => Leitet mir Innsbruck weiter.**
- **CS-JOB => Datei => Ladungs-Importe => Fenster mit allen Datensätzen die Importiert wurden** => **Links unten werden Reiter "Orange"** wo eine Änderung vorgekommen ist

## Spezielle NOS 💚 Eigenheiten die ich wissen sollte
- **Drucken** => **Einlastungs_Datum** festlegen im **2020** => **CS-Datensätze werden überspielt!**

## Heinz Cargo-Support & Excel
- H richtet **Lieferlisten zusammen mit Deckblatt**
- **wie vielen WAB's & WANN**
- Liste wird von **Liefervorschägen** nach abgearbeitet.
- **-1 Werkttag vor Ausliefertag** => Heinz richtet **Ladelisten Tourenplan & Frachtbiefe** für Monteure

- Bei **Multi-Split => Abkoffern** => und **Checkmark: Zu Tour verbinden**
- Immer **Abkoffern** => Button = **Useless** => Helmuth got it from far
- **bis = Bereitstellungs_Datum  // bis(E) = 06:00 Uhr**
- Bei **Multi-Split** ohne **Checkmark "zu Tour verbinden"** => **Useless**
    - Heinz just teached me Cargo-Support
- Bei Plant_SC_Leiter Touren gibt **Heinz** nachträglich => **Durchführungs_Datum & Monteurname** ein.

## Tourenplan für SC-Leiter
- **Rechtsclick** => **Archivierte Dokuemente** => **Archivierte Dokumente des Auftrages** => **Doppelclick auf Datensatz** => **Dokument_1**
- **Mehrfachdruck testen ??? - Möglich ?**
- **F6** => **Tourenplan** => **Dokument_2**
- **Deckbaltt** => **Dokument_3**
- **Sollte für SC-Leiter reichen => Aufpassen Preise sollten runter => (Archivierte Dokumente)**

## Fixing NOS 💚
- **Excel-Makro für Tourenplan_Helmuth** => Zu **Tourenplan_BML** (Almost working)
- **2020** => **Liefer-Vorschläge export as PDF's ??** => **Sortieren nach Gebieten (Zonen)** => **Auto. Mails** an Zuständigen!
- **Sorting nach neuer PLZ Methode** => **[Siehe Interaktive Aus-Lieferkarte]()**
- Wie am besten **Sortiern:**
    -   Als erstes wird nach **Prod_Datum** sortiert => "KW"
    -   Gebiete => nach **PLZ** sortiert
    -   Filter nach **Gebiet (Zonen)**
    -   Auswahl der Tour nach **Volumen/Gewicht** => Tour anlegen und auf Liste vermerken.
    -   Senden **Alle** =>
- **🔁 Beim Nachplanen:**
    - POST / DIREKT Aufträge
    - Filter nach KW bzw. Prod. Datum
    - Achten auf: Montage-Aufträge mit wenig Gewicht, Blöde Lage => PLANT SC LEITER
    - Aufträge können in Folge_KW verschieben => Muss keine Eingabe ins 202 erfolgen!

## Situation mit hinterlegten Volumen/Gewicht/Montagezeit
- Hinterlegten **Montagezeiten/Volumen/Gewichte** aus ?
- Gibt es einer **📅 Fehler-Liste** ?
- **Alle nachrechnen** => **Montagezeiten vorallem bei Handelswaren (Zulieferant)**
- **🛢️ Volumen** => **Handelswaren (Zulieferant) nicht hinterlegt** => **Liste starten**
- **Riesige-Liste von 100 Artikeln** => Vorhanden => Kösung finden um Daten nachzureichen => Hali Vergleich / Bene fragen?
- **Alex 💚** helfen sie brauchen jemadnen um das zu checken!

## 

## ⛰️ Vorarlberg ist Anders (ULGB)
- **Voralberg** => Wie ist die **Zusammenarbeit mit ULBG** (Wie Isabel ? Wer 📗 Avisiert ?)
- **Helmuth** => Stellt Tour nach Kunden zusammen => Mail mit **Listen wie SC-Leiter** nur ohne **Stops**
- **Nachtrag wird nicht erstellt** = Oliver anrufen absprechen **Listen Rücksendung -> Übernehmen in Cargo - ERIK**

## ⛵ Kärnten ist Anders (Riegler)
- **Riegler** => **1:1** wie **ULGB**
- **Verladereihenfolge** kommt zurück => Wird nicht im Cargo-Support nachgetragen!

## 🍳 L&M Küchennmontagen 
- Heinz schickt => Deckblatt, Liefervorschläge, Pläne** => L&M => Stops, Zeiten => kommen zurück und **wir Avisieren!**

## Auftragsnummern => Ersten 2 Ziffern sind Geschäftstelle GG/AAAAAA/00/00
| Ausftragsnummer | WAB Nummern-Kreis | Geschäftstelle | Service-Center |  Kommentar |
| --- | --- | --- | --- | --- |
| **10** | **12** |  **Werk** | **WNeudorf** | |
| **11** | **12** | **Wien** | **WNeudorf** | |
| **12** | **12** | **Neudörfel** | **WNeudorf** | |
| **13** | **13** | **Linz** | **Linz** | |
| **14** | **13** | **Salzburg** | **Linz soon Salzburg** | |
| **15** | **15** | **Tirol** | **Innsbruck** | |
| **16** | **16** | **Kärnten** | **Klagenfurt** | |
| **17** | **17** | **Steiermark** | **Graz** | |
| **18** | **18** | **Vorarlberg** | **Dornbirn** | |
| **23** | **Je nach Gebiet** | **Händler** | **Je nach Gebiet** | |
| **43** | **43** | **Deutschland** | **Relogg Partner** | |
| **45** | **45** | **Budapest** | **Export** | |
| **47** | **45** | **Budapest** | **Export** | |

## Vorproduktion
- Im Regel bei **📅 Terminverschiebungen durch kunden wird Prod.**
- trotzdem laut altem Datum durchgeführt & **Auf WAB verladen.** (Genug Container vorhanden)
- **Bei Kleinigkeiten => verschieben wir machmal die Produktion**

## 🚛 Fuhpark

| Kennzeichen | Fahrzeugklasse (3,5t/7,5t/LKW) | max. Volumen(m³) | max. Gewicht(kg) | Fahrzeug Modell | Kommentar |
| --- | --- | --- | --- | --- | --- |
| **WY-871AW** | **3,5t** | X | ~300kg | Fiat Ducato | |
| **WY-742AT** | **7,5t** | ~15m³ | ~1500kg | Iveco Eurocargo | SC WNeudorf ? |
| **WY-307AX** | **LKW & Hänger** | 30m³ | X | LKW | |
| **WY-308AX** | **LKW & Hänger** | 30m³ | X | LKW | |
| **WY-450AX** | **LKW & Hänger** | 30m³ | X | LKW | |
| **WY-659AN** | **LKW & Hänger** | 30m³ | X | LKW | |
| **WY-741AT** | **LKW & Hänger** | 30m³ | X | LKW | |


- Innsbruck = Spedition_Kusztrich / Heinz nochmal klären
    - Deutschland => Weiß, Nuri(Kameen) => Mit Silke klären 🛑
    - Neudörfl = Spedition_Kusztrich / ÜBEX (Bei viel Auslastung)
    - 5 WAB-LKW (WY-307AX, WY-308AX, WY-450AX, WY-659AN WY-741AT) alle Hänger möglich.
    - Fahrzeuge: 7,5t (WY-742AT)
    - Fahrzeuge: 3,5t Tonnen Sprinter (WY-871AW) ab KW 13. in Wiener Neudorf = SC Wiener Neudorf Keine Fahrzeuge für Neudörfel (Direkt)
- haben eigenen LKW der 2 WAB laden kann
- 8 Monteure wurden nach SC-Wiener Neustadt geben => 4 Monteure über bei NOS direkt
- Sped_Kusztrich ist sehr zuverlässig laut NOS
- NOS Info in Cargo-Support fast nie 30m³ auf einem WAB => Mehr 20 - 25 m³
- Oft wird Produkion während des einräumen des WAB's fertiggstellet => Aber nicht in der richtigen Reihenfolge = Wird dann für den falschen Stopp verladen
- 🔨 Alex glättet Produktion
- 📑 Alex hat Kapazitätsreport von 2020 => Bekommt Helmuth täglich
- 📅 Team-Gespräch sollte eigeführt werden mind. 1x die Woche
- Deutschland hat NOS 💚 nur 10-15 m³ pro Woche => Lösung muss her
- Export gibt es bei NOS nicht viel manchmal Bratislava, Budapest und (Slovenien => macht Riegler Klagenfurt)
- 💡 Excel-Listen mit Helmuth und Heinz müssen eingeführt werden => Teams & Excel vorbereiten => ALMOST DONE
- 💡 PDF-Parser und Crap-Scripts für automation seems very powerful at NOS

## NOS Verspätungen Produktion 26.03.2025
- Amt der Tiroler LRG => Hali Auftrag für Morgen NOS Produktion hat nicht geklappt => 1. Tag zuvor erst verständigt.
- Pollauf Philipp hat bei NOS nachgefragt. => Nach mehrmaliger nachfrage erst heute verschieben bekannt gegeben.


## 17.03.2025 Start Erik: Innsbruck in Cargo-Support Abgabewoche => KW15

- **Cargo Support NOS💚-Style**
- **Akutelle Aufgabe** nur **Touren** in **Cargo-Support zusammenstellen & 📗 Avisieren & 🔄 Nachplanen**
- **Gebiet (J) => St. Anton am Arlberg => Immer Innsbruck verplanen**
- **Gebiet (H) => Saalfelden & Zell am See & Mittersill => Immer Innsbruck verplanen**

## 31.04.2025 Start Erik: Voradelberg, Kärnten, Steiermark
- **Planung KW18** => **Start**

## 🌍 Gebiete Erik - Hersteller NOS 💚 => + "Salzburg" => Service-Center kommt Itzlinger
- **⛰️ Voradelberg - [VBG] - (SC Dornbirn) - {Partner=ULGB} +  Deutschland Süden (PLZ 8XXXX) [DE] + Liechtenstein [FL] + Schweiz [CH]**
    - **🚀 Untertouren** macht **[⛰️ ULGB Voradelberg ist Anders](https://github.com/Kingslayer9988/bgo_holding/blob/main/documentation/Profi-Tour.md#%EF%B8%8F-vorarlberg-ist-anders)**
    * **NOS = 1-2 WAB pro 📅 KW** lt. Helmuth  04.03.2025
    * **(J)** = Zone in **Cargo-Support**
    * **SC Dornbirn** = SC-Leiter => **Oliver L. (ULGB)**
    * **📅 Liefertage** => Eher **Anfangs der Woche** => Aber Woche geht von **Mo. - Fr.**
    * **❌ Kein Küchenmonteur** => **L&M Küchenmontage**

- **🚠 Tirol - [T] - (SC Innsbruck) + 🇮🇹  Italien [ITA] (Export)** 
    - **1️⃣ Gebiet Erik ab 19.03.2025**
    - **NOS 💚 = 2-4 WAB pro 📅 KW** lt. Helmuth  04.03.2025
    * **(I)** = Zone in **Cargo-Support**
    * **SC-Leiter** => **Thonhauser F. & Agostini T.**
    * **📅 Liefertage** => **WAB's** eher **MO - MI** => Woche geht von **Mo. - Fr.**
    * **❌ Kein Küchenmonteur** => **L&M Küchenmontage**

- **⛵ Kärnten - [KTN] - (SC Klagenfurt) - {Partner=Riegler}**
    - **NOS 💚 = 1-2 WAB pro 📅 KW** lt. Helmuth  04.03.2025
    * **(G)** = Zone in **Cargo-Support**
    * **SC Klagenfurt** => SC-Leiter => **Koeck M.  & Bader S.**
    * **🚀 Untertouren & 📗 Avisieren** macht **[⛵ Riegler Kärnten bei NOS ist Anders ~ Wie ULGB](https://github.com/Kingslayer9988/bgo_holding/blob/main/documentation/Profi-Tour.md#%EF%B8%8F-vorarlberg-ist-anders)**
    * **Küchenmonteur vorhanden ✔️**
    * **Kleines SC-Lager** => **Max. 2 LKW pro Tag (1x Platz & 1x Rampe)**

- **🌳 Steiermark - [STMK] - (SC Graz) + Kroatien [HR] + Slovakei [SI] + Solvenien [SLO]**
    - **NOS 💚 = 5 WAB pro 📅 KW** lt. Helmuth  04.03.2025
    * **(F)** = Zone in **Cargo-Support**
    * **SC Graz** => SC-Leiter => **Koeck M.  & Bader S.**
    * **❌ Kein Küchenmonteur aber gute Monteure (Außnahme) ✔️** 

- **🇩🇪 Deutschland Gesamt {Partner=Relogg}** => **❓ Noch nicht sicher für mich** => (UNSURE ASK DENISA❓ => Teams)
    * **(???) => probably  [DE]** = Zone in **Cargo-Support**
    * **Ähnlich wie Voradelberg (Relogg ~ ULGB) = 🚀 Untertouren & 📗 Avisieren**
 
- **🛖 Neudörfel [BGLD] {Helmuth macht Alles}**
    - **Wandmonteure & Küchenmonteure** => **Beides in Neudörfl vorhanden**

> [!NOTE]
> **SC = Service-Center 🏩**
