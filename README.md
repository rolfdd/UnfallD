# UnfallD – Unfalldaten binär komprimiert

Dieses Repository enthält die **deutschen Unfalldaten (2016–2025)** in einem extrem kompakten Binärformat sowie die Werkzeuge, um sie zu erzeugen, zu durchsuchen und zu visualisieren.

## Daten

Die Rohdaten stammen von der amtlichen Unfallstatistik (DESTATIS) und wurden in einem mehrstufigen Prozess aufbereitet:

- **Quelle:** Excel-Tabellen (ca. 400 MB)
- **Aufbereitung:** VBA-Makro extrahiert, filtert, sortiert und komprimiert die Daten
- **Ergebnis:** `UnfallD_16-25.bin` – 36 MB, 16 Byte pro Datensatz
- **Index:** `UnfallD_16-25_index.csv` – Gemeindeschlüssel → letzte Zeile

## Warum binär?

| Format | Größe | Bemerkung |
|--------|-------|-----------|
| Excel (Roh) | ~400 MB | Nicht versionierbar, langsam |
| CSV | ~180 MB | Textbasiert, aber ineffizient |
| **Binär (16 Byte)** | **36 MB** | **~78 % kleiner, extrem schnell** |

Durch Bit-Packing (Art+Typ, Kat+Licht+Zust, Booleans) und kompakte Datentypen (Byte, Long) wird jeder Datensatz auf nur 16 Byte reduziert.

## Makro (`BinaerExport.vba`)

Das VBA-Makro in `docs/BinaerExport.vba` erzeugt aus den Excel-Rohdaten die Binärdatei und den Index:

1. **Filter:** Bundesländer (1–16), optional Rad– oder Stadtfilter
2. **Sortierung:** Gemeindeschlüssel, Jahr, Monat, Tag, Stunde …
3. **Spaltenanpassung:** Entfernen überflüssiger Felder, Umsortierung
4. **Bit-Packing:** 
   - Art (4 Bit) + Typ (4 Bit) → 1 Byte
   - Kat (2 Bit) + Licht (2 Bit) + Zust (2 Bit) → 1 Byte
   - 7 Booleans (Rad, Pkw, … PLST) → 1 Byte
   - Reserve-Byte → 1 Byte
   - lon/lat × 1.000.000 als Long → 8 Byte
5. **Export:** Binärdatei + Index (Gemeindeschlüssel → letzte Zeile)

## Web-Anwendung

Die Datei `index.html` lädt die Binärdatei per `fetch` + `Range`-Request:
- Nur der benötigte Bereich wird geladen (z.B. Dresden)
- TypedArray (`Uint8Array` + `Int32Array`) für schnelles Parsing
- Anzeige als CSV oder auf einer Leaflet-Karte

## Repo-Struktur

```
UnfallD/
├── data/
│   ├── UnfallD_16-25.bin                              (Binärdaten, 36 MB)
│   └── UnfallD_16-25_index.csv                        (Index Gemeindeschlüssel)
├── docs/
│   ├── UnfallD_roh_16-25_Testdaten.xlsx               (Excel-Testdatei)
│   └── BINAER_Export_Standard_MitIndex_OhneHeader.vba (Makro als Text)
├── index.html                                         (Streaming + + Pivottable + Leaflet)
└── README.md
```

