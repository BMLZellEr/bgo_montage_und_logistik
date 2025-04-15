# 💎 Infos Cargo-Support

## Cargo Support Tickets - Live-System

## 🐞 Bugs and Possible Bugs

- **🐞 Auto. Verladereihenfolge**  
  Live-System is behaving unexpectedly → Could be user error → Tested with Klaus → **He believes it's not user error**

- **🐞 Split-Menu**  
  Issues with System/Split/Übergabe/Liefer_Date functionality → **No significant impact on NOS**

- **BIG🐞 After Spliting the Tour "Fixtermin="ContainsFixDate" only stays at the Überstellungs-Tour**  
  SHOULD BE IN BOTH TOURS! → Possibly connected to the date calculation issue

- **MAYBE?🐞 Montage/Verlade/Gesamt-Zeit at Tours that are Split**
  Affects **Date&Time** Calculation for Container_Tours → **Entlade_Zeit ≠ Montage_Zeit**

## 😀 Quality of Life Improvements / Non-Bugs

- **😀 Mitarbeiter needs to be set on Auftragsebene** => **Calculation is slower than my Grandma**
  Very annoying → Need either button: **"Planungsdaten ändern"** or change on **"Tour-Ebene"**

- **😀 Same issue for "Entlade_Start" / "UnloadingStart"**  
  Add at "Entlade_Start" input in button: **"Planungsdaten ändern"**

- **😀 Split-Menu**  
  No field for **Container-Nummer** → Add **"FreeText1"** in Split Menu → **Both tours will be visibly connected**

- **😀 Doppelklick Menü deaktivieren**  
  Add a checkbox option → Current behavior is very annoying → I Open this Menu by accident around 10x a Day

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

## Übernahme Aufträge die Hänhgengeblieben sind
- CS-Job => Ladung Import Aufträge => Orange Felder sind neu

## Export der DB in Excel
- Rechtsklick auf Spalte => Exportiere Tabelle

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

## Zeit/Stop-Planung von Unter-Touren
- Tour markieren => Rechtsklick => Tour => **"Tour Zeiten/Einschränkungen Anpassen"** => Verladung & Tour richten
- Bsp.: 3 Kunden von WAB => Kunde1 - Kunde2 - Kunde3 Aufkoffern => Kunde3 -Kunde 2 - Kunde 1 => Entladen
- Nach richten Aktualisieren => Rechtsclick => **"Aktion: Live Dispo: Aktivität für Tour erzeugen"**

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

## Personal
- As I use it more and more its actually sad that the Software got cancelet it could have been great Software.
- It needs big fixes and a lot of performace imporvements but the featureset is actually fucking big.
- I wonder can closed source software actually be a good thing ?
- Is there a Open-Source Tourenplanung-Software ? If not maybe a good buisness idea

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

