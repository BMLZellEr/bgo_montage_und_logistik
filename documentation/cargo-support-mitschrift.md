## 🌍 Gebiete Erik (Hersteller NOS 💚) 
- **⛰️ Voradelberg - [VBG] - (SC Dornbirn) - {Partner=ULGB} +  Deutschland Süden (PLZ 8XXXX) [DE] + Liechtenstein [FL] + Schweiz [CH]**
    - **🚀 Untertouren** macht **[⛰️ ULGB Voradelberg ist Anders](https://github.com/Kingslayer9988/bgo_holding/blob/main/documentation/Profi-Tour.md#%EF%B8%8F-vorarlberg-ist-anders)**
    * **NOS = 1-2 WAB pro 📅 KW** lt. Helmuth  04.03.2025
    * **(J)** = Zone in **Cargo-Support**
    * **SC Dornbirn** = SC-Leiter => **Oliver L. (ULGB)**
    * **❌ Kein Küchenmonteur** => **L&M Küchenmontage**

- **🚠 Tirol - [T] - (SC Innsbruck) + 🇮🇹  Italien [ITA]**
    - **1️⃣ Gebiet das ich verplanen soll**
    - **NOS 💚 = 2-4 WAB pro 📅 KW** lt. Helmuth  04.03.2025
    * **(I)** = Zone in **Cargo-Support**
    * **SC-Leiter** => **Thonhauser F. & Agostini T.**
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

- **🇩🇪 Deutschland [DE] {Partner=Relogg}** => **❓ Noch nicht sicher für mich** => (UNSURE ASK DENISA❓)
    * **(???) => probably  [DE]** = Zone in **Cargo-Support**
    * **Ähnlich wie Voradelberg (Relogg ~ ULGB) = 🚀 Untertouren & 📗 Avisieren**


| NAME | SCHEME | INFO |
| --- | --- | --- |
| **ABH Kundenamen TT.MM** | **ABH $KUNDE $ORT $WT(TT)** | Selbstabholungs-Tour am Donnerstag |
| **Postversand TT.MM** | **POSTVERSAND $ORT $KW** | Alle Postversand ohne Fixtermin in KW_12 |
| **Kundenname KundenOrt TT.MM** | **$KUNDE $ORT $WT(TT)** | Direkt-LKW Tour zur Baustelle am Donnerstag |
| **SC TIROL WAB TT.MM** | **SC WAB für $Fahrzeuge Touren $WT(TT)** | Lager-WAB Mittwoch <br> für LKW & 7,5t & 3,5t Touren <br> "RAMPE" in "Bemerkung Transport" |
| **SC 7,5t Tour Innsbruck - Telfs TT.MM** | **SC $Fahrzeuge Tour $ORT_S $ORT_E $WT(TT)** | Untertour 7,5t Fahrzeug <br> Laden von Lager-WAB |
| **SC LKW Tour Fügen - Jenbach TT.MM** | **SC $Fahrzeuge Tour $ORT_S $ORT_E $WT(TT)**  | Untertour LKW Fahrzeug <br> Laden von Lager-WAB|
| **Plant SC Leiter / 6 Kunden** |  | Unter-Tour mit 6 Kunden <br> die SC-Leiter Plant |
| **SC WAB für ULGB MI** | **SC WAB für ULGB $WT(TT)** | **[⛰️ ULGB Voradelberg ist Anders](%EF%B8%8F-vorarlberg-ist-anders)** <br> Lager-WAB Mittwoch <br> Voradelberg **immer "Platz"** |
| **ULGB Feldkirch - Dornbirn FR** | **ULGB $ORT_S $ORT_E $WT(TT)** | **[⛰️ ULGB Voradelberg ist Anders](%EF%B8%8F-vorarlberg-ist-anders)** <br> Tour am Freitag (lt. ULGB) |


# 🚥 Cargo-Support for Dummies 🚥

## 💎 Infos Cargo-Support

## DB-Ansicht besteht aus 2 Teilen
- Oben => Datenbank => Unverplante Datensätze ==> Pretty much 1:1 Datenbank-Ansicht in Profitour
- Unten => Disponiert => Verplante Datensätze ==> Touren-Fenster in Profitour => in DB-Ansicht sichtbar!
- **Datensätze in Roter Schrift** sind **Regie-Aufträge** => **Keine ❌📦 Produktion**
- Standartansicht => **Std. Dispo Basis** => laut CS-Video
- Aufträge können direkt in **DB-Ansicht gefilterd** werden => Feld: **Auftrags-Nr.**
- Fahrzeug: Dropdown
- Fahrtstatus in Standartansicht => Rechts => Disponiert / Offen => Bei uns Unten/Oben
- Frachtbrief & Tourenplan & Transportauftrag & Auftragsbest => Alle von **DB-Ansicht** aus **druckbar** => **"F6"**
- **Wichtig** = Arbeitsweise ist anders als in Profi-Tour => Zuerst Lager-WAB planen & Anschließend Tour teilen

## Programmstart & Filter der Aufträge
- Filter in der DB-Ansicht ist oben links => Keine "KW" => Montag bis Sonntag der gewünschten KW auswählen.
- Gebiete **ZONE** nach Karte Filter => Karte .pptx per Mail von Isabel
- **Dispotopf** => Filter => **ZONE** Auswählen => Zone siehe **.pptx Karte**
- **⚠️ FIXTERMIN** => Eigene Spalte in **DB-Ansicht** mit **Checkmark ☑️** => Spalte: **"Erstmögliche Lieferung"** = Spalte: **"Letztmögliche Lieferung"**
    - **⚠️ FIXTERMIN** villeicht nur **"Grundfilter"** in **Cargo-Support**
- **Bereitstelldatum == Prod_Datum** in **Cargo-Supprot** => **Wie geht das das bei NOS 💚??** (UNSURE ASK ALEX/HELMUTH)
- **Feld: Frei_Text_1 (Freetext_1)** == **Interne WAB-Nr**
- **Spalte: Entlade_Start** == **Ankunft fixieren (Profi-Tour)** == **Liefer_Uhrzei** => **WICHTIG: Immer bei 1. Stopp in Tour setzen!**
- **Aufpassen = Keine Softwarekontrolle** für **Bereitstell_Datum** ist **später** als **Entlade_Start_Datum**
- **Bermerkungs_Felder** sind **1:1** wie in **Profi-Tour**

## Planung von DIREKT_BAUSTELLE Touren (In DB-Ansicht möglich)
- Bei DIREKT_BAUSTELLE => In Datenbank-Ansicht => Tour markieren => **Zu Tour Verbinden**

- Neues Fenster öffnet sich => Fenster: **[ZU_TOUR_VERBINDEN_FENSTER]()**
    - **Fahrzeug, Fahrer, Frachtführer, Freitext** => Laut CS-Video
    - **Namen** **Namens-Schema == Profitour**
    - Alle anderen Felder **können leer gelassen** werden!

- **Alle Stopps der Tour markieren** => Button **[Planungsdaten Ändern]()**
    - **Bereitstellungsdatum** vergeben! => Wie im Profi-Tour (Freitag für Montag / 1 Tag vorher)
    - Container auswählen (NORMAL, EGAL,JUMBO)
    - Bei **2 Container auf 1 Tour** => 2. Container **händisch in DB-Ansicht** einstellen!
    - **Entladestart == 08:00** | Bereitstellungsdatum + 1 Werktag
    - **Spedition hinzufügen** => Same as in Profi-Tour

## Rechtsclick-Menü** öffnet sich nur bei makierten Datensatz => Häckchen gesetzt.
- Button: **In den Planungspool** = Datensätze in das Virtuelle_Touren_Fenster übernehmen!
- Button: **Letztes_E_Avis_Termin** => 3h Aviszeit für Kunden hinterlegen => wird nicht gesendet!
- Button: **Fahrt Bearbeiten** auf erstellter Tour öffnet Fenster: **Fahrt_Bearbeiten_Fenster**

## ⏱️ Montagezeit kalkulieren bzw. Transportauftrag checken
- Öffne **CS-JOB** => Sollte sowieso bei **jedem Start mitlaufen**
- In **DB-Ansicht** **markieren des Datensatzes** => **Rechtsclick** => Für **Öffne is CS-JOB**
- In **CS-JOB** gehe zum **Reiter** => **Ladungsinhalte**
- Sollte bei **jedem Datensatz / Auftrag** kontrolliert bzw. **nachgerechnet** werden.
- Datensätze die mit **(00:30 h)** eingegeben sind können **gemeinsam kontrolliert** werden => **Stimmen meistens!**
- Aufpassen bei **Trennwänden(RS/RG)**, **Trennwand/Glaswand** etc. mit **SC-Leiter abklären**
- Aufpassen bei **🍳 Küchen** => **Montagezeit** stimmt **Transportauftragsreport** teilen durch **Personen** => **Standzeit fixieren**
- Bei **🍳 Kleinen Küchen** mit **SC-Leiter abklären** ob wir sie selbst machen können!
- **Küchenmonteure teilweise** selbst vorhanden! **[Siehe Hier für NOS💚](NOS.md)**
- In **Microsoft Teams Datein** nach Auftragsnummer suchen => **Pläne downloaden** => **An SC-Leiter weiterleiten**
    - Aufpassen beim **Download von PDF's** => Downloaded gerne altes File! => **Teams is great Software**
- Für **Kalkulations-Skript-Profi-Tour** => **[Siehe Hier](../README.md#working-still-needs-adoption-for-hersteller)**

## Guide für das Erstellen von Unter-Touren
- In DB-Ansicht **gewünschte Datensätze markieren** => Aufträge auf **Lager-WAB's**
- Rechtsclick => Button: In den **[Planungspool]()**
- Sobald Datensätze im **Planungspool** liegen werden sie auf **[Minimap]()** angezeigt!
- Nach Auswahl der Punkte auf **Minimap** => Rechtsclick => Button: **Hinzufügen**
- Ausgewählte Datensätze sind nun im **Virtuelle_Touren_Fenster**
- **Tour** nach **gewünschter Fahrt richten** => Button: **Tour erzeugen**
- Neues Fenster öffnet sich => Fenster: **[Planungsdaten Ändern_ODER_NEUES_FENSTER_1]()** => Namen vergeben (Bsp.: SC WAB für 3,5t Touren)
- In der **DB-Ansicht** wird nun einen **generiete Tourennummer** vergen und in **gleicher Farbe** für die **Tour hinterlegt**
- Bei **2 WAB's (Container) auf einer Tour** => Eg. **Sattel oder LKW + Anhänger** => in der **DB-Ansicht** die Container **Diagonal in die 2 Datensätze** reinziehen. 
- Tipp: **Sortieren nach Tour-Nummer** => legt gewünschte **Tour-Stopps** untereinander!
- **Container-Nummer => WAB-Nr.**
- **Entladestart** ==> Auf **Stopp Nr.1 der Tour** stellen => **Anliefer_Datum** ausfüllen (Bereitstelldatum + 1 WKT)
- Feld: **Entlade_Start** = zeigt **Entlade/Belade Reihenfolge für Stopps => 1, 2, 3, 4** etc.
- Markieren der Tour => Fenster: **[Planungsdaten Ändern_ODER_NEUES_FENSTER_2]()** => Datum und Uhrzeit für LagerWAB einstellen => **06:00 Uhr + Anliefer_Datum**

## Guide für das Erstellen von WAB-Touren NACH ERSTELLEN der UNTERTOUREN
- **Makiere alle Datensätze** die auf den **Lager-WAB** kommen sollen =>
- **In diesem Fall:** **Alle Datensätze der Untertouren** die von **diesem Lager-WAB laden** sollen.
- Rechtsclick => Button: **Multi-Split**
- **Multi-Split** => Teilen des Auftrags **LagerWAB-Tour + Unter-Tour**
- Neues Fenster öffnet sich => Fenster: **Multisplit_Fenster**
- Feld: **Adresse** => **SC-Lager** einstellen => (Bsp.: SC-Graz)
- Feld: **Fahrttype** => **Zustellung** einstellen (NEU bzw. auto. in Profi-Tour) => Möglicherweise noch andere Zustellungsarten (ASK SOMEONE)
- **WAB-Tour** sollte hiermit **abgeschlossen** sein => In **DB-Ansicht** nun im oberen Teil in **Schwarzer-Schrift** und mit **WAB-Nummer** sichtbar!
- Anschließend **markieren der WAB's** => Rechtsclick => **Veraldereihenfolge schreiben** => **Verladeplanung+ in Profi-Tour**

## Guide für das Erstellen von WAB-Touren VOR ERSTELLEN der UNTERTOUREN
- **Makiere alle Datensätze** die auf den **Lager-WAB** kommen sollen => Rechtsclick => Button: **Multi-Split** => (Meist kl. Aufträge <= 15m³)
- **Multi-Split + Dispo Kombi** => **Teilen** des **Datensatzes (Auftrags)** in: **2 Datensätze LagerWAB + Untertour**
- Neues Fenster öffnet sich => Fenster: **Multisplit_Fenster**
- Feld: **Adresse** => **SC-Lager** einstellen => (Bsp.: SC-Graz)
- Feld: **Fahrttype** => **Zustellung** einstellen (NEU bzw. auto. in Profi-Tour) 
- **WAB-Tour** sollte hiermit **abgeschlossen** sein => In **DB-Ansicht** nun im oberen Teil in **Schwarzer-Schrift** und mit **WAB-Nummer** sichtbar!
- Anschließend **markieren der WAB's** => Rechtsclick => **Veraldereihenfolge schreiben** => **Verladeplanung+ in Profi-Tour**

## Planungsdaten Ändern => bzw. Unterfenster zum bennen die sich öffnen
- Abgleichen der Fenster => Logische Namen für Fenster überlegen => Aktuell Unübersichtlich
- **SIND DIESE FENSTER ALLE NOTWENDID ?? => DB-Ansicht ist ja in Cargo-Support editierbar**

## Virtuelle Touren Fenster
- **Einfügen von Stopp - Vor / Nach** => Wie im **BIOS-Boot-Reihenfolge** (F5/F6)=> **Höher/Tiefer** mit **ausgwählten Datensatz**
- **Montagezeiten(Standzeiten)** => Werden auch hier Festgelegt => **[Transportauftragsreport]()** => Kann hier angeshen werden **JA/NEIN** ? (ASK SOMEONE)
- _Note: Nobody knows will try later or at home_
- **Gewicht bzw.- Volumen hier checken** => Je nach Ergebniss => Fahrzeug wählen
- Button: **Löschen = Aus Tour entferenen**
- Button: **Tour_erzeugen** => **Feld: Name** => (Bsp. SC 3,5 t Tour Graz - Gresten) => Again siehe Profi-Tour
- **Feld: Freitext_1** => **WAB-Nummer** wie **Profitour** => Nummern-Kreis

## Multi-Split_Fenster
- **"Datum"** Reiter => **totally Useless**

- **"Split-Information"** Reiter:
    - Feld: **Adresse** = **SC-Lager** hier im Drop-Down Menü auswählen => Bsp. (SC-Graz, SC-Innsbruck)
    - Feld: **Fahrttyp** = Immer **Zustellung** im Drop-Down Menü auswählen => Möglicherweise noch andere (UNSURE ASK ALEX/HELMUTH)
    - **❌ Rest Useless ?**

- **"Fahrzeug-Info"** Reiter:
    - **Fahrzeug** = Working
    - **Anhänger** = ❌ Not used
    - **WAB 1** = ❌ Useful but not working (tested)
    - **WAB 2** = ❌ Useful but not working (un-tested)
    - **Fahrer 1 & Fahrer 2** = ❌ Not used
    - **Frachführer** = Spedition = Working
    - **Freies Kennzeichen** = ❌ Useless

- **"Abrechnung"** Reiter => **❌ totally Useless**
- **"IC-Tochter"** Reiter => **❌ totally Useless**
- **Checkmark ☑️ Zur Tour verbinden** => Auswählen falls **[Unter-Touren]()** noch **nicht erstellt** sind.
- **Tourname** = Hier **Tourname** vergeben => Siehe **[Naming-Schema]()** => (Bsp. SC WAB GRAZ KW12 DI TT.MM.JJ)

## Minimap
- Steuerung ist **Invertiert** im vergleich mit Profi-Tour **[STRG] [SHIFT]**
- Farben sind ohne bedeutung ==> Should be fixed!

## Planungspool
- Find out if this is better or worse than in Profitour.
- Items vom Planungspool können nur one by one gelöscht werden => **Andere Kunden von Cargo-Support**
- Falls Planungspool zu viele Unnötige Items enthält => Planungspool komplett löschen => Neu Anlegen
- **Neues_Fenster_Planungspool** => Mann kann neuen Pool erstellen => Fenster muss geschlossen werden um zu aktualisieren.
- "F5" => Auswahl vornehmen => PLanungspool nun sichbar!
- _Note: Because multiple "Planungspools" are possible maybe make them according to "Zone's"_
- _Note: Should make moving inbetween them way faster ?_

## Doppelclcik-Menü Wird von niemanden verwendet!

## Export zum Hersteller
- Bei **NOS 💚** leider noch **keine Möglichkeit (Schnittstelle)** in Software-2020 zu **exportieren.**
- Export => **Rechtsclick auf Datensatz oder markieren mehrerer Datensätze** => Button: **"Aktion->BGO Tourenfeedback"**  

## Weird Shit
- ❌ Useless_Felder == **Startzeit_ab** & **Startzeit_bis** & **Freitext_2** & **Freitext_3** & **ab** & **von(B)** & **bis (B)** & **bis** & **von (E)** & **bis(E)**
- Felder = **Freitext_1** missing in **Fenster_Multi_Split**
- Felder
- Es soll Felder geben die wenn man sie verstellt sind sie für alle verstellet ??? (UNSURE ASK BIG BOSS) 
- Button: **Tour drehen** & **Tour optimieren** im Virtuelle_Touren Fenster funktionieren nicht. 

## Watch Guides from Web and Read Website - Maybe someuseful Infos
- Übung 1 | done | (CS-JOB) => Watch @ home 
- Übung 2 | done | Partly Useful => Frachtbriefe etc.
- Übung 3 | done | (CS-JOB) => Watch @ home 
- Übung 4 | done | (CS-JOB) => Watch @ home 
- Übung 5 | done | (CS-JOB) => Watch @ home 
- Übung 6 | done | (CS-JOB) => Probably also Useful for (CS-PLAN) =>
    - Spalten-Ausrichtung und Namensänmderung => wie in Profi-Tour + Fixieren möglich
        - Maske bearbeiten => Layout editieren
    - Funktionszeile links unten => Anzahl der Datensätze
    - Gruppieren möglich (Spalte nach oben ziehen => Seems Useless but maybe good => Thats how "Dispo-Fenster" is done probably
        - Gruppieren => Feature für Disponmiert/Offen Spalte => Wieso nicht mit Fahrtstatus ? 
    - Vorsicht Änderungen der Ansicht werden nicht auto. gespeichert => Rechts oben Safe-Icon => Dropdown-Menu
    - Datums-Filter-Einstellungen probably wrong for us "BGO"
    - Grundfilter seems useful => DB-Makro to mark when 2 or more Variables fit the case!
    - Freigabe von Bildschirmeinstellungen
- Übung 7 | done | (CS-JOB) => Watch @ home (Fahrer/Mitarbeiter-Anlage)
- Übung 8 | done | USEFUL (CS-PLAN)
    - Reiter Offen/Disponiert => DB-Ansicht => Unten/Oben
    - Reiter Live-Dispo useful zur Kontrolle
    - Änderungen/TV => Suche über Filter: Auftragsnummer, WAB-Nummer etc.
    - Alle Änderungen für Fahrzeug,Cont, etc. über Dropdown in DB-Ansicht möglich
    - Multi-Druck möglich => Makieren & "F6"
- Übung 9 | done | USEFUL (CS-PLAN)
    - Containerdispo => Zeile oben mit Reitern => Daten ergänzen die in CS-JOB nicht richtig eingeben wurden
    - Adresse ändern => Makieren => Reiter **Wegpunkte** => SC-GRAZ etc. (Excel funktion mit Runterziehen possible) 
    - Rechtsclick -> **Auftragsdaten ergänzen** -> BS_DATUM, AUFN_DATUM etc. => Siehe unten!
    - Rechtsclick -> **Containernummer ändern** => Sub_Window
    - CS_PLAN => Fahrt =! Tour ==> Fahrt = DIREKT_BAUSTELLE || Tour = Lager-WAB + Unter-Touren
    - DIREKT_BAUSTELLE => Auswahl der Datensätze (30m³) => Rechtsclick => **Zur Tour verbinden**
    - ZU_TOUR_VERBINDEN_FENSTER ==> **Siehe unten** == Mitschrift =! Videos 
    - TOUR_UND_NEUE_FAHRT_FENSTER ==> **Open Software and Check out** => Drag&Drop => Video unclear
    - Rechtsclick auf **Datensatz in bestehender Tour** => Fahrt aus Tour entfernen == **Stopp auflösen in Profi-Tour**

- Übung 10 | half done | USEFUL (CS-PLAN) => Getting Tired continue tommorw
- Übung 11 | X | 

## Pros & Cons (Profi-Tour VS Cargo-Support)
- BIG PRO = Fahrzeuge / Datum / Container & mehr kann **alles aus der DB-Ansicht gemacht** werden!
- BIG PRO = Programm ist um einiges **Modularer** als Profi-Tour (Fenster-Settings and how they behaive is still pretty buggy)
- BIG PRO = **Frachtbrief & Tourenplan & Transportauftrag** aus **DB-Ansicht** Druckbar (Siehe oben bei Filter)
- PRO = Gebiete sind viel deutlicher aufgeteilt => **Zonen >>> Bundesländer** (Way better for my Brain, better logic) 
- PRO = Postaufträge werden komplett automatisch verarbeitet! => 0 User Input (saving lots of useless loadtimes also Mistakes)
- PRO = DIREKT_BAUSTELLE Aufträge komplett in DB-Ansicht planbar! (saving lots of useless loadtimes)
- PRO = Definitv mehr Möglichkeiten um Ansichten zu verändern (Dunno how Useful)

- BIG CON = Work is in two seperate Programs => CS-JOB & CS-PLAN (Profi-Tour DB-Ansicht = Teil von CS-JOB)
- BIG CON = Aktualisieren "F5" is necessary
- BIG CON = Programm ist deutlich langsamer als Profi-Tour & Keine Ladeanzeige!
- BIG CON = Minimap is way worse than Profi-Tour => G-Maps API is perfect => But no Coloring of Dots is horrible & Loadingtimes are also horrible
- Tiny CON = Keinen "KW" Filter in DB-Ansicht
- Tiny CON = Steilere Lernkurve
- UNSURE = **Planungspool** - Maybe better maybe worse not sure yet.  


# Everything below this Heading is a Copy from Profi-Tour and will get changed along the Way (For Easy STRG+C & STRG+V)
# 🏢 BGO Mitschrift Erik Zellhofer 🏢

## Chronologischer Ablauf (Will Still have Mistakes - Aktuell für $Hersteller_H)

- **📅 Montag:**
    - ⌛ **Daily Quests** 
    * 📄 **Frachtbriefe & Tourenpläne für SC-Leiter** => Für **$Folge_KW_01** -> Bsp.: **Montag_KW_08 für KW_09** => **[Guide](#-guide-für-frachtbriefe--tourenpläne-2-pdfs)**
    * 📑 **Auslieferlisten für SC-Leiter** => Für **$Folge_KW_01** -> Bsp.: **Montag_KW_08 für KW_09** => **[Guide](#-guide-für-auslieferlisten)**
    * ▶️ **Starten der Planung von mindestens $Folge_KW_4** => Bsp.: **Montag_KW_08 für KW_12** => **[Guide](#-guide-für-das-erstellen-von-wab-touren--direktpost)**
- **📅 Dienstag:**
    - ⌛ **Daily Quests**
    - **📗 Avisieren mindestens $Folge_KW_2**=> Bsp.: **Dienstag_KW08 für KW_10** => **[Guide](#-avisieren)**
- **📅 Mittwoch:**
    - ⌛ **Daily Quests**
- **📅 Donnerstag:**
    - ⌛ **Daily Quests**
    * 📦 **Produktions-Export (WAB's-Export)**  => Für **MINDESTENS $Folge_KW_04** (Bsp: **Donnerstag_KW_08 für KW_12**) => **[Guide](#-guide-für-die-abgabe-der-produktion)**
- **📅 Freitag:**
    - ⌛ **Daily Quests** 
    * 🚚 **Transportbestellung** => Für **$Folge_KW_02** (Bsp: **Freitag_KW_08 für KW_10**) => **[Guide](#-guide-für-die-transportbestellung)**
    * 📑 **Auslieferliste für Hersteller** => Für **$Folge_KW_01** (Bsp: **Freitag_KW_08 für KW_09**) => **[Guide](#-guide-für-auslieferlisten)**
    * 🔨 **Möglicherweise Glättung der Produktion** => Für **Abgabe_KW** der **Produktion** (Bsp.: **Freitag_KW_08 für KW_12**) => **[Guide](#-guide-für-die-glättung-der-produktion)**

- **⌛ Daily Quests:**
    - ♻️ Bei jedem Import von Hersteller_H **06:30 & 14:00 Täglich** (Hr. Roither für Probleme)
        - 🔄 **Nachplanen** => **Alle neuen Aufträge** wieder verplanen für **$Folge_KW_1 & $Folge_KW_2 & $Folge_KW_3 & $Folge_KW_4** => **[Guide](#-nachplanen)**
        - 🔶 Ausnahme **Freitags 14:00: Nur DIREKT/POST** Lieferungen verplanen **(Runde der Mitarbeiter)** 
    - 📺 **TV's** machen die Anfallen! => **[Guide](#-guide-für-tv-tv--termin-Verspätung)**

# 🚥 Profi-Tour for Dummies 🚥

### ⚖️ LAW Number One ➡️ Don't mess with Profi-Tour-Profile from Isabel
### ⚖️ LAW Number Two ➡️ Bei jeder Änderung -> Export für Hersteller
### ⚖️ LAW Number Three ➡️ Don't even try SPECIAL_LETTERS(§$%) in Profitour
### ⚖️ LAW Number Four ➡️ Don't open many Instances because of double "Übernahme in Planung"
---

## 💎 General Infos:

### Aufträge bzw. Datensätze

| ART | INFO | KOMMENTAR |
| --- | --- | --- |
| **[DIR / STK / PST]** | **Alle DIREKT** - **Kein LAGER-WAB** | **als 1. Verplanen** = **Durchläufer** |
| **[PST]** | Postversand **pro KW** oder **FIXTERMIN** | **>= 5m³** = **Falsche Eingabe** von **Hersteller** |
| **[DIR]** | **Keine Fixtermine** | Durchläufer - **Immer eigene Tour** <br> Achtung muss **[Avisiert](#-avisieren)** werden! |
| **[ABH]** | **Selbstabholung** durch Kunden | **Durchläufer** - **Immer eigene Tour** | 
| **Internet-Order (AB-Nr: 034XXX)** | **Fixtermin** = **EGAL** | Fixtermin ist für Kunde **nicht sichtbar** |
| **Neutrale Bestellung (AB-Nr: 04XXXX)** | Für Tischer Unbeschriftet | **Fahrer & Montagepersonal** muss **instruiert** werden |


### 🤵‍♂️ Kunden-spezifischische Infos
- 🟢 => **Kunde wurde bereits Avisiert**
- 🔴 => **Kunde hat altes Aviso = Zeiten können nicht gehalten werden!** => **FORCE_GREEN** oder **Neu Avisieren** **[Guide](#-avisieren)**
- **FORCE_🟢_BUTTON** => Button: **Kunden_Avis_Bestätigen** => **Kein Avis versendet trotzdem 🟢**
- **Tournummer** mitschreiben bei denen **FORCE_🟢_BUTTON** benutzt wurde!

### 🧾 Kommentar Felder
- **Bemerkung Transport** => Infos für die **Spedition** => **Bsp.: (PLATZ, RAMPE)**
- **Bemerkung Versand** => Infos für **Verladung** => **Bsp.: (Manuelle Ladereihenfolge)**
- **Bemerkung Montage** => Infos für **Montage Arbeiter** => **Bsp.: (FIXTERMIN | Hinweise | Adresseänderung/Zufahrt)**
    - Bei **Adress-Änderungen/Fehler** => **Mail an Hersteller** um weitere Fehler zu vermeiden!
    - Falls **Plant-SC-Leiter-Tour** => SC-Leiter anrufen!

### 🕋 Fracht
- Spedition müssen den **Touren richtig zugeordnet** werden (**Links oben** aus dem Fenster ziehen)
- Spedition ist **NOTWENDIG** für **alle Lager-WAB-Touren** => Bsp. **Spedition_Winkler**
- Es kann immer nur **1 Fahrzeug pro Tour** zugeteilt werden!
- Bei **Touren** mit **verschiednen Fahrzeugtypen** => **größtes Fahrzeug nehmen** => (Bsp.: 3,5t + 7,5t Tour => 7,5t Fahrzeug)

#### 🛢️ Container

| Container | Verwendung | Kommentar |
| --- | --- | --- |
| **NORMAL** | Bei **Lager-WAB's**  | Für Untertouren mit LKW oder 3,5/7,5t |
| **EGAL** | Bei **Direkt_Baustelle** | Nur für Direkt_Baustellen | 
| **JUMBO** | Bei **Direkt_Baustelle & 30-35m³** | Nur für Direkt_Baustellen > 30m³ |

#### 🚛 Fahrzeuge
     
| Fahrzeug | Gewicht | Volumen | Name | Kommentar | FS |
| --- | --- | --- | --- | --- | --- |
| **Mont-Bus** | **???** | **4m³** | GRAZ_Direkt | Für Arbeiter um zu Baustelle zu kommen <br> zB. Bei DIREKT_BAUSTELLE | B |
| **3,5 Tonnen** | **600 kg** | ∞ | GRAZ_Bus_1 | Nur für Unter-Touren verwendet | B |
| **7,5 Tonnen** | **1000 kg** | **15 m³** | GRAZ_Bus_1 | Nur für Unter-Touren verwendet | C |
| **LKW** | ∞ | **30 m³** | LKW_KESIC | Für Unter-Touren verwendet oder um Lager-WAB anzuliefern | C |
| **LKW+Hänger** | ∞ | **60 m³** | Kein FZ=SPED_WINKLER | Um 2 Lager-WAB's am selben Tag anzuliefern <br> Nicht auf Rampe | C |
| **Sattel** | ∞ | **60 m³** | Kein FZ=SPED_QUEHENBERGER | Nur bei **DIREKT_BAUSTELLEN** verwendet <br> Billiger als LKW + Anhänger <br> BGO hat **keinen eigenen Sattel** => **[Transportbestellung](#-guide-für-die-transportbestellung)** | C |

- **Fahrzeuge are completely messed up in some areas => Needs fixing - Absprache mit SC-Leiter**
    - **[Excel-Listen](Excel.md)** von Isabel => Eintragen von **3,5t | 7,5t** Spalten
    - **Excel & Profitour & Realität** = STIMMEN **NICHT** ÜBEREIN 💀

## 🗺️ Minimap - Kleine Karte links unten - Enthält ausgewählte Datensätze/Aufträge
- 🔵 Blau = Kunde/Baustelle **Unverplant** oder nur am **Lager-WAB**
- 🔴 Rot = Kunde/Baustelle mit **Trennwand**
- 🟠 Orange = Kunde/Baustell **Verplant**
- ➡️ Pfeile = Zeigen die **Tour von Kunde zu Kunde** (Luftlinie ✈️)
- 💹 Je **größer der Punkt desto mehr 💵💶💴**
- ⌨️ Hold **[STRG]** für **Neue Auswahl**
- ⌨️ Hold **[SHIFT]** für **Zur Auswahl hinzufügen**

## 💣 Guide für das erstellen von WAB-Touren & Direkt/Post:
- Zuerst immer **DIREKT & POST - Linien** erstellen!
    - Datensatz bei **DIREKT | POST | DIREKT_Baustelle** => **NICHT TEILEN**
    - Maximal **1 Postversand_$ORT_Linie - pro KW** (Außnahme PST mit FIXTERMIN) 
    - Bei **DIREKT_Touren** auch **Verlade-Reihenfolge** richten & **WAB-Nummer** vergeben

- Bei **DIREKT_Baustelle_Touren** auch **Verlade-Reihenfolge** richten & **WAB-Nummer** vergeben
    - Aufpassen wenn **3 WAB's oder mehr** gebraucht werden => Immer **maximal 2 WAB's** pro **Tour**
    - Für den **3 bzw 4. WAB** eine neue Tour anlegen & **Ankunftszeit fixieren** (Bsp.: **1&2-WAB - 08:00** | **3&4-WAB - 12:00**) 
    - Aufpassen im **Normalfall** - **nicht mehr als 1 Kunde** => **Spedition** kann normal **nicht mehr als 2h auf Baustelle** warten

- Bei **2 Containeren(WAB's) auf einem LKW + Hänger** bei **Verladereihenfolge** auch **2 WAB-Nummern vergeben** und auf **Hänger** stellen!
- **2 Arten von WAB's** => Eintragen von **[Rampe oder Platz]** in **Bemerkung Transport**
- **Maximal 1 WAB auf Rampe** pro Tag => **2ter WAB** muss auf **Platz**
    
    - **↗️ [Rampe]** (Lager-WAB) => Bei **3,5t & 7,5t Unter-Touren** => **(eher = Bei vielen Kleinkunden)**
        - Container: **$Spedition_NORMAL**
        - Fahrzeug: **3,5t** oder **7,5t - Fahrzeug**
        - **Niemals** einem **Jumbo-WAB** zur **Rampe**
        - Nur **max. 1 LKW** auf **Rampe**

    - **🅿️ [Platz] (Lager & Direkt-WAB)**  => Bei **LKW Unter-Touren** **(eher = Bei wenig Großkunden)**
        - Container: **$Spedition_NORMAL**
        - Fahrzeug: **LKW - Fahrzeug**
        - Aufpassen auf **Zufahrten etc.** oft ist es **nicht möglich** mit **LKW bis zur Baustelle** zu kommen
        
- **🔁 Automatische Verlade-Reihenfolge**
    - **Auto-Button** berücksichtigt **keine Manuellen Änderungen** => **[Reset]**
    - Bei **Manuelle-Ladereihenfolge** verwendet wurde => In Kommentarfeld **Bemerkung Versamd** vermerken!
    - Wieso **Manuelle-Ladereihenfolge** => Wenn **2 Unter-Touren** an **1 Tag** aus **1 WAB geladen** werden.
    - **Verlade-Reihenfolge = 99** = Reihenfolge = **EGAL**
        - Oft bei **LKW + Hänger** auf einer **DIREKT_BAUSTELLE_Tour** angewendet

- **Vorproduktion:** => **WAB in Produktionswoche** anlegen & Bei **Unter-Tour** => Button: **Datum fixieren** => **Lieferdatum auswählen**

- **WAB-Nummern Vergabe** = Siehe **Nummern-Kreis** auf **Ausdruck**
    - Aufpassen: Bei **Vorproduktionen** gibt es einen eigenen **Nummern-Kreis**

- **💸 Tipps um Geld zu sparen**
    - **2 WAB's** mit **LKW + Anhänger an 1 Tag** ist **billiger** als **2 LKW's** mit jeweils **1 WAB** an **2 verschiedenen Tagen**
    - Aufpassen teilweise **Spedition & SC-Leiter** aufgrund von **Auslastung** nicht Happy
    - **LKW gewünscht:** - Manchmal ist es **teurer umzuladen** in **3,5t bzw 7,5t Fahrzeug** - In diesen Fällen mit LKW zu **kleinem Kunden zu fahren!**

> [!NOTE]
> **Excel-Makro** für KW_$$ zur **Fehlererkennung** erstellen => Sobald **2 Touren aus 1 WAB** laden!

## 🚀 Guide für das erstellen von Unter-Touren:

- Datensätze nach **WAB's (WAB-Nummern)** auswählen und **in die Auswahl** übernehmen. (Bsp. Alle Datensätze am "WAB 050")
    - WAB's sollten schon erstellt sein => **📦[Produktions-Abgabe](#-guide-für-die-abgabe-der-produktion)** bzw. **💣[Guide für WAB Touren](#-guide-für-das-erstellen-von-wab-touren--direktpost)**

- **⚠️ Wunschtermine & Fixtermine** bzw. **größere Baustellen** beachten
- **⚠️ FIXTERMIN** Im **Tourennamen** und im Feld **Bemerkung_Montage** vermerken (Außnahme Internet/Postversand)

- **⏱️ [Montagezeit (Transportauftragsreport)](#%EF%B8%8F-montagezeit-kalkulieren-bzw-transportauftrag-checken) nachkalkulieren!**

- **❌ Falls sich die Tour** aufgrund von **Fahrzeiten** oder **Montagezeiten NICHT** ausgeht:
    - Datensätze auf **[Minimap](#%EF%B8%8F-minimap---kleine-karte-links-unten---enthält-ausgewählte-datensätzeaufträge)** ansehen und nach **Orten/Fahrtstrecke** teilen
    - Datensätze im **Bearbeitungs-Fenster** - **aus der Tour lösen** bzw. **Stopp entfernen** bei erstellter Tour.

- **🚛 Fahrzeug** nach **Gewicht, Volumen(m³)** der **Tour** einteilen **( 3,5t | 7,5t | LKW )** **[Siehe Hier](#-fahrzeuge)**

- **☎️ Bei knappen/komischen-Touren** => **Absprache mit SC-Leiter** => **Plant-SC-Leiter-Tour erstellen**
    - **Ausnahme**: **[Vorarlberg ist anders](#%EF%B8%8F-vorarlberg-ist-anders) (ULGB)**

 - 🔨 **[Glättung der Produktion](#-guide-für-die-glättung-der-produktion)** des **jeweiligen Herstelller** berücksichtigen

- **📝 Naming-Scheme für Touren/WAB's:**

| NAME | SCHEME | INFO |
| --- | --- | --- |
| **DIREKT FirmaABC Innsbruck MO** | **DIREKT $KUNDE $ORT $WT(TT)** | Direktversand am Montag |
| **ABH FirmaABC Innsbruck DO** | **ABH $KUNDE $ORT $WT(TT)** | Selbstabholungs-Tour am Donnerstag |
| **Postversand Innsbruck KW12** | **POSTVERSAND $ORT $KW** | Alle Postversand ohne Fixtermin in KW_12 |
| **Postversand Innsbruck MO FIXTERMIN** | **POSTVERSAND $ORT $WT(TT) FIXTERMIN** | Postversand mit <br> Fixtermin am Montag |
| **FirmaABC Innsbruck DO** | **$KUNDE $ORT $WT(TT)** | Direkt-LKW Tour zur Baustelle am Donnerstag |
| **SC WAB für LKW und 7,5t und 3,5t Touren MI** | **SC WAB für $Fahrzeuge Touren $WT(TT)** | Lager-WAB Mittwoch <br> für LKW & 7,5t & 3,5t Touren <br> "RAMPE" in "Bemerkung Transport" |
| **SC 7,5t Tour Innsbruck - Telfs MI** | **SC $Fahrzeuge Tour $ORT_S $ORT_E $WT(TT)** | Untertour 7,5t Fahrzeug <br> Laden von Lager-WAB |
| **SC 3,5t Tour Meran - Brixen MI** | **SC $Fahrzeuge Tour $ORT_S $ORT_E $WT(TT)**  | Untertour 3,5t Fahrzeug <br> Laden von Lager-WAB |
| **SC LKW Tour Fügen - Jenbach DI** | **SC $Fahrzeuge Tour $ORT_S $ORT_E $WT(TT)**  | Untertour LKW Fahrzeug <br> Laden von Lager-WAB|
| **Plant SC Leiter / 6 Kunden** |  | Unter-Tour mit 6 Kunden <br> die SC-Leiter Plant |
| **SC WAB für ULGB MI** | **SC WAB für ULGB $WT(TT)** | **[⛰️ ULGB Voradelberg ist Anders](%EF%B8%8F-vorarlberg-ist-anders)** <br> Lager-WAB Mittwoch <br> Voradelberg **immer "Platz"** |
| **ULGB Feldkirch - Dornbirn FR** | **ULGB $ORT_S $ORT_E $WT(TT)** | **[⛰️ ULGB Voradelberg ist Anders](%EF%B8%8F-vorarlberg-ist-anders)** <br> Tour am Freitag (lt. ULGB) |

> [!NOTE]
> Ich finde **[SC_ORT] & [Datum]** sollte jeder Tour beigefügt werden **[TT.MM.JJJJ]** \
> Erlaubt später alle Gebiete bzw. mehrere KW's auf einmal zu bearbeiten!\
> **SC WAB für [Fahrzeuge]** => Ist **Useless** und muss **oft geändert** werden!\
> Hier sollte **meiner Meinung** nach die **Excel-Liste** angepasst werden!

> [!TIP]
> **Schnelles Übernnahme in die Planung** - Planung Reiter offen lassen & Tour aus Planung Entfernen\
> **Schnelles Verplanen von Änderungen** - Teilen von Aufträgen auch im Unbearbeitet-Fenster gemacht werden\
> **Auswahl in Verladeplanung+ nicht verlieren** - Nach Auswahl zuerst Bereitstell_Datum dannach Produktions_Datum klicken

## 🔄 Nachplanen
- **WAB's richten** und **📄 Frachtbriefe/Tourenpläne** vor dem **[Erneut-Avisieren](#-erneut-avisieren-nach-avisieren)** machen um **geänderte Touren rot markiert** zu haben.
- Aufpassen bei der **🔁 Automatische Verlade-Reihenfolge** => Falls der Tag nicht **mindestens 📅 8 Tage** in der Zukunft liegt
- **🛑 Keine 🔁 Automatische Verlade-Reihenfolge** machen Wegen **Etiketten-Druck**
- **[📄 Frachtbriefe & Tourenpläne für SC-Leiter](#-guide-für-frachtbriefe--tourenpläne-2-pdfs)** müssen neu gesendet falls die Woche bereits gesendet wurde.


## 📙 Vorraus Avisieren (Vor-Avisieren)
- Kann auch **Telefonisch** gemacht werden => Gut für **Stammkunden**
- **E-Mail & Kontakt & Datum** aus Profitour (Datenbank-Ansicht) nehmen und **E-Mail** erfassen!
- Wichtige Daten: **Kunde, KW, Auftragsnummer**
- Mail-Wichtigkeit auf **"High" 🍄**
- Vorlage:
````
Sehr geehrte Damen und Herren! [Oder Persöhnliche Anrede]

In der KW_$$/KW_$$ bekommen sie von der Firma [Hersteller_$] die neuen Büromöbel geliefert und montiert.
Wir würden die Montage gerne am [TTTT], [DATUM] um [UHRZEIT] Uhr beginnen.
Die Montagearbeiten wären am [DATUM]-[Tageszeit] abgeschlossen.

Bitte um eine kurze Antwort, ob der Termin für sie in Ordnung ist.

Danke im Voraus.
````
> [!NOTE]
> Mittels PDF-Parser kann dieses Prozess **easy automatisiert werden**\
> Alle notwendigen Daten könnten aus **Aviso PDF** gezogen werden => Much ⏰ safed here\
> Prozess wird **Lokal automatisiert** => **Keine Anbindung an Profitour** => **Windwos_Contextmenu_PS_Script**)

## 📗 Avisieren
- Filter in **Datenbank-Ansicht** nach **[Prod + Lieferwoche] & [MO-SA]**
- **Avisieren** immer im **Touren-Fenster** => **[OBEN]**
- **Alle Touren** außer **[DIR](#aufträge-bzw-datensätze)** und **[PST](#aufträge-bzw-datensätze)** und **[Lager-WAB-Touren](#-guide-für-das-erstellen-von-wab-touren--direktpost)** Avisieren!
- **Alles 2x Kontrollieren**
- **Datum bei Touren** & **Ladereihenfolge WAB's 3x kontrollieren**
    - **Datum** immer in dem **Bearbeitungs-Fenster** unter **Stopps** kontrollieren!
    - **Datum** kann sich aufgrund von Tagesarbeitszeit oder **Falscher [Montagezeit](#%EF%B8%8F-montagezeit-kalkulieren-bzw-transportauftrag-checken)** auf **nächsten Tag** rutschen!
    - **Fehler passieren hier leicht** und **keine Software-Kontrolle** für Fehler wie diese!

> [!NOTE]
> Probably a big reason that "AI" is not already taking Disponenten-Jobs (needs a sick Algorithm 😄)\
> Bad for me but  => **Local Deepseek with Live Training here 👀👀**

## 📘 Erneut Avisieren (Nach-Avisieren)
- **Meistens** der **Fall** das Kunde den Termin nicht halten kann. => Checken ob **[Vor-Avisiert](#-vorraus-avisieren-vor-avisieren)** wurde  
- Oft der **Fall** gleich nach dem **[Avisieren](#-avisieren)** => Kunden bekommen Aviso und wollen **neuen Termin**
- Filter nach **[Transportaufträge & Auftrag]** => **Auftragsnummer des Kunden** (Siehe **Mail** bzw. oft per **Telefon**)
- Gehe zu **Tour** die den **Kunden(Stopp)** beinhaltet
- Aufpassen ob andere Kunden auf Tour einen **⚠️ Fixtermin** vereinbart haben
- Checke ob schon **alle Kunden** auf **Tour** avisiert wurden
- Falls **JA** checke alle **anderen Kunden(Stopps)** und versuche **Zeit/Datum des Kunden** anzupassen mit **möglichst wenig neues Avisos**
- 🛑 Vorsicht kann **Teufelskreis** werden _(Bsp.: Tour drehen => 5 neue Avisos die den Kunden nicht passen und 1 das passt)_
- Falls **NEIN** schieben ist leichter weil **kein neues Avisieren notwendig** ist!
- Beim schieben auch auf **🔨 [Auslastung des Herstellers](#-guide-für-die-glättung-der-produktion) Rücksicht** geben
- Bei **$Hersteller_H** (Villeicht weitere) auf **[❄️ Frozen Zone & ☑️ Checkmark Eingelastet](#other-things-i-got-told-and-are-probably-useful)** **Rücksicht** geben

## 📄 Guide für Frachtbriefe & Tourenpläne (2-PDF's):
- Filter -> **[Transportaufträge & Prod + Lieferwoche]** => **TT.MM.JJ - TT.MM.JJ** der gewünschten **KW auswählen**
- Für **Hersteller** == Alle Touren außer **[Lager-WAB's] => Immer Freitags für $Folge_KW_1**
- Für **SC-Leiter** == Alle Touren außer **[POST, Direkt, Lager-WAB's]**
- **Auswahl der Tour im Oberen Fenster** -> Frachtbrief -> Print as PDF -> Naming_Scheme: **[TT]_TOURNUMMER**
- **Auswahl der Tour im Oberen Fenster** -> Tourenplan -> Print as PDF -> Naming_Scheme: **[TT]_TOURNUMMER_TB**
- Frachtbrief und Tourenplan für SC-Leiter immer **$Montags_Aktuelle_KW** für **%Nächste_KW** -> Bsp.: **Montag in KW7** -> **Ausdrücke für KW08 an SC-Leiter**
- Immer **Excel-Tabelle der Folge_KW$$** auf Papier am Tisch haben und Touren nach Senden der PDF abhaken! -> **2PDF's pro Touren [_NON-TB & _TB]**
- Bei **Unklaren Anfahrten 🔮** für **Baustellen** => Mit **Hilfe** von SC-Leiter **Anfahrts-PDF** erstellen
    - Enthält **Pläne & Bilder vor Ort** oder **Google-Maps**
- Ablage auf Onedrive 📁 that never syncs even 10MB's 😨😂

## 📑 Guide für Auslieferlisten:
- Filter -> **[Transportaufträge & Prod + Lieferwoche]** => **TT.MM.JJ - TT.MM.JJ** der gewünschten **KW auswählen**
- Für **Hersteller** == Alle Touren außer **[Lager-WAB's] => Immer Freitags für $Folge_KW_1**
- Für **SC-Leiter** == Alle Touren außer **[POST, Direkt, Lager-WAB's]**
    - Auslieferlisten werden von **SC-Leiter** bearbeitet und **alle 3 Herstellern kombiniert**
    - Haben dann **Infos** über genaue Touren, Mitarbeiter, Fahrzeuge etc.
- **Touren** markieren => **Export** => Wähle **Auslieferliste** -> **Saved as .csv on Remote [Profi-Tour-Server]** -> Copy to **[Work-PC]** and open in Excel
- Apply Excel ✨ **Magic-Makros** ✨ **[Strg+Umschalt+A]** = Austria || **[Strg+Umschalt+D]** for "Germany"
- Bei **Plant-SC-Leiter-Touren:** => **Lösche Inhalt der Zellen in Grün** und schreibe **"Plant SC-Leiter"** in **Rot & Fett**
- Save as **.xlsx**
- Naming-Scheme: **ALL_$ORT[WIEN,LINZ,GRAZ,etc]_[KW$$].xlsx**

> [!NOTE]
> Upgrade Makro to do the last 3 Steps also automatically

## ⏱️ Montagezeit kalkulieren bzw. Transportauftrag checken
- In egal welchem Fenster => **Rechtsklick auf Datensatz** => **Transportauftragsreport** => Für **Zeit-Berrechnung bzw. Calc.sh**
- Sollte bei **jedem Datensatz / Auftrag** kontrolliert bzw. **nachgerechnet** werden.
- Datensätze die mit **(00:30 h)** eingegeben sind können **gemeinsam kontrolliert** werden => **Stimmen meistens!**
- Für **Kalkulations-Skript** => **[Siehe Hier](../README.md#working-still-needs-adoption-for-hersteller)**
- Aufpassen bei **Trennwänden(RS/RG)**, **Trennwand/Glaswand** etc. mit **SC-Leiter abklären**
- Aufpassen bei **Küchen** => **Montagezeit** stimmt **Transportauftragsreport** teilen durch **Personen** => **Standzeit fixieren**
- Bei **Kleinen Küchen** mit **SC-Leiter abklären** ob wir sie selbst machen können!
    - **Küchenmonteure teilweise** selbst vorhanden! **[Siehe Hier für NOS💚](NOS.md)**

## 📦 Guide für die Abgabe der Produktion:
- Filter -> **[Transportaufträge & Hali: Lieferwoche]** => **[JJJJ/KW] auswählen**
- Es gibt eine **Liste mit Abgabedaten** für **Hersteller_H**
- Filter **$Folge_KW4** => Alle **Gebiete**
- Im **Unbearbeitet-Fenster** makiere alle Aufträge => **Checke Gewicht und m³**
- Erstelle Alle **[DIREKT & PST & DIREKT_BAUSTELLE & LAGER_WAB-Touren](#-guide-für-das-erstellen-von-wab-touren--direktpost)**
- **KONTROLLE auf 🔫 Schnellschüsse** => Schnellschuss = **Lieferungs_KW** ist **nicht** in **Produktions_KW**
- Links oben auf Export ==> **Produktion des Herstelles wird eingelastet!!**
- **Deadline** verschieden für alle 3 Hersteller => **[NOS💚](NOS.md)**
- Hersteller_H => 3 Wochen unterschied => **Bsp.: spät. Freitag KW7 für KW_11**

## 🔨 Guide für die Glättung der Produktion:
- im **Team abklären** wer **WAB's** leicht **verschieben kann um Auslastung** zu richten!
- Auf **Fixtermine** und **Montage-Auslatung** achten!
- Mit **SC-Leiter abklären** ob mann die Tage scheiben kann!
- Bei **$Hersteller_KW =! $Wirkliche_KW** auf Tage achten (Bsp.: $Hersteller_KW = DI-DI) => **NOS 💚 sollte das nicht zutreffen!**
- Öffne **Verladeplanung+** der **Lager-WAB's** die **geschoben** werden sollen => **Ändere Datum => Fertig**

## 🚚 Guide für die Transportbestellung:
- Filter -> **[Touren & Aufladestelle/Zeitraum]** => **[MO -SA] der gewünschten KW** auswählen!
- Nach dem Filter Alle Gebiete kontrollieren **(2 unteren Spalten in Excel zählen)** => **Direkt & Lager-WAB's**
- **Schnittstelle -> Export in Excel** -> safe as .xlxs Name: **Transportbestellunhg_KW__**
- **Container-Nummer & Container Anzahl & Ladeeinheiten kontrollieren!**
- **Ladereihenfolge checken** => oft **35/36/35** => Delete **35**
- **m³ mit Fahrzeugen abglichen** bei => (Bsp.: **40m³** => **LKW + Hänger** oder **2 LKW's**)
- Transportbestellung im richtigen **MS-Teams** Ordner ablegen!
- **Letzter im Team** fügt die **Excel-Imports der Mitarbeiter zusammen** und sendet sie an **Spedition Winkler**
- **Sonderbestellungen** wie Transmato, Quehenberger(Sattel) **sofort** bei Auftrag an Spedition senden!

## ⛰️ Vorarlberg ist Anders
- **Untertouren werden von ULGB zusammengestellt** nach einer Liste von uns!
- Filter in Datenbank nach **[Dornbirn + FL + CH + DE]**
- Datenbank Ansicht auf **[Isabel-Voradelberg]**
- Oben auf **Schnittstelle** => Export als **Excel file (.csv)**
- Auto-Summe pro WAB erstellen für **[Gewicht und m³]**
- Text: **Anlieferung am $Datum zum Lager oder Direkt zum Kunden (Siehe Excel file oder Profitour)**
- **Jumbo-WAB vermerken**
- Als **E-Mail an Oliver** senden - er sendet die **Untertouren** zurück!
- **Unter-Touren** wieder in **Profi-Tour** eingeben!
* Oliver sendet **PDF via Mail** - **auf "A3" ausdrucken!** => **Untertouren** stehen unter der **Excel-Tabelle**
    - Alle Touren nach Oliver anlegen **(Naming: ULGB $Start_ORT $End_ORT Wochentag[TT])**
    - Alle **Stopps(Kunden)** Richtien & richtige **Ankunft fixieren!**
    - Alle Touren **Nachkalkulieren** & **Standzeit fixieren!** **[Siehe Hier](#%EF%B8%8F-montagezeit-kalkulieren-bzw-transportauftrag-checken)**

> [!NOTE]
> Make One and Only ⛰️ Vorarlberg Excel-Makro for Isabel (STRG+UMSCHALT+V)

## 📺 Guide für TV (TV = Termin-Verspätung)
- Nur möglich wenn **Untertouren fertiggestellt** sind! 
- Filter -> **[Transportaufträge & Auftrag]** => **Auftragsnummer der TV** (Siehe TV)
    - Tipp: **Enter = Zeile hinzufügen** (Filter nach mehreren Aufträge)
- Im **Touren-Fenster** **Tour die TV beinhaltet** auswählen
- Im **Bearbeitungs-Fenster** wähle **Transport_Aufträge** => **Positionen**
- Es gibt **4 verschiedene Fälle von TV's**
---
* 1️⃣ **Wenn Ganzer Transportauftrag (AB-Nummer)** Verzögert ist:
    * Gehe zu **WAB der TV-Tour** beinhaltet => Button: **Aus Tour lösen und Teilung Rückgängig machen**
---
* 2️⃣ Wenn nur **eine Postion im Auftrag** verzögert ist:
    * **Checken** ob bei der **Verzögerten Position Montagestunden** dabei sind und **keine weiteren Produkte**
    * Wenn ❌ **NEIN** => Button: Im **Bearbeitungs-Fenster** => **Neuen Transportauftrag erzeugen**
    * Wenn ✅ **JA** => Gehe zu **WAB der TV-Tour** beinhaltet => Button: **Aus Tour lösen und Teilung Rückgängig machen**
    * Wenn 🔐 die **Verzögerte Position Montagestunden** und andere **nicht verzögerte Produkte** beinhaltet
    * Löse die **Verzögerte Postition & Inklusive Montagestunden aus der Tour**
---
* 3️⃣ Wenn **ein Teil einer Position** verzögert ist:
    * Beim Teilen zuerst die **Anzahl der Teilstücke** dann die **Stückzahl** => (Bsp.: 3/7 Sessel in TV => Teilen 2 Teilstücke mit 4 Rest)
    * Button: Im **Bearbeitungs-Fenster** => **Neuen Transportauftrag erzeugen**
---
* 4️⃣ Wenn **die TV eine Postversendung** ist => **Verladeplanung+** auf neues Datum richten!!
---
- Wenn **WAB-Tour in der Zukunft** => **Verladereihenfolge neu richten**
- Anschließend oben Button: **Selection** (Falls Vergessen = Probleme im Bearbeitungs-Fenster) => Button: **In Auswahl**
- Im Bearbeitungsfenster **Auftrag auswählen** => **Sammelmappe erstellen**
- Naming-Scheme: **[$ORT_$KW_DER_TV_KUNDE]** => (KW: Siehe Links unten auf TV)
- Filter nach **$KW_WO_TV_HINGEHÖRT** in **[Lieferwoche]**
- Öffne **Sammelmappe** => Button: **Übernahme in die Planung**
- Auftrag ist nun im **Unbearbeitet-Fenster** => Splitten, WAB, etc. => **behandeln wie neuen Auftag**

## Profi-Tour besteht aus 2-Teilen bzw. Ansichten (Datenbank & Planungs - Ansicht)

### Ansicht_1 = Datenbank_Ansicht
- Direkt nach dem **Start des Programms** sichtbar!
- Import und **Auswahl der Gebiete** findet hier statt! **(Bsp.: Wien, Linz, Steiermark, etc.)**

### Ansicht_2 = Planungs_Ansicht = Besteht aus 3 Fenstern:

#### Fenster_1 = TOP = **Touren-Fenster**
- Hier findest du deine **Zusammengestellten Touren** (Lager-WAB's, Untertouren, Post-Tour)
- Der Inhalt dieses Fensters wird **Exportiert zum Hersteller** **📦[Produktions-Abgabe](#-guide-für-die-abgabe-der-produktion)**

#### Fenster_2 = MIDDLE = **Bearbeitungs-Fenster**
- **Teilen/Auflösen/Ändern** von **Touren** & **Checken von Transportaufträgen** und vieles mehr
- **Ändernen von Touren** => **Löst Stopp aus WAB & Stopp aus Tour**
    - Im Bearbeitungs-Fenster: **Aus Tour lösen und Teilung Rückgängig machen**
    - Im Unbearbeitet-Fenster: **Teilung Rückgängig machen**     

#### Fenster_3 = BOTTOM = **Unbearbeitet-Fenster**
- Im Prinizip wie **Datenbank-Ansicht** nur mit **übernommenen Datensätzen** haben

## Other Things I got told and are probably Useful:
- **❄️ Frozen Zone (Hali) ❄️**
    - Ab **Folge_KW** der **Abgabewoche** !
    - Aufträge können von **Hersteller_H NICHT mehr bearbeitet werden.**
    - BGO kann noch Schieben aber **nur nach hinten (später)** - Nur bei **[Checkmark-Eingelastet]** ☑️

- **Trennwände: (Speical_Montage_Workers)**
    - Montage => pro 1 Laufmeter = 3h
    - Demontage => pro 1 Laufmeter = 1,5h
    - [RG] **Glaswand** => Kommt aktuell nur von **$Hersteller_B** => Bei anderen Hersteller => **Falsche Eingabe**

- **Microsoft Teams**
    - Aufpassen beim **Download von PDF's** => Downloaded gerne altes File! => **Teams is great Software**
  
- **Konflikte** mit **BGO & SC-Leiter** (UNSURE ASK BIG-BOSS 😕)
    - Wer ruft Kunden bei welchen Problem an ? 💀
    - Abschieben auf $Hersteller & $SC-Leiter ==> **und Umgekehrt!**  💀
    - **WHO THE FUCK IS TO BLAME AND FOR WHAT 💀**
    - **Keine Verantwortung übernehmen :question: => Alles Problem von Hersteller** :question:
    - Maybe ask all the Colleges ?

- **$HERSTELLER**
    - NOS -  🟢 - Hersteller_N -> 🔵 Schrift
    - HALI - 🔵 - Hersteller_H -> 🟢 Schrift
    - BENE - 🔴 - Hersteller_B -> ⚫ Schrift

> [!NOTE]
> Later make Guides switchable per $_Hersteller and Color them accordingly 🌈
