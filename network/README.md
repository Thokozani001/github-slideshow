# North West site network — KMZ map and improved spreadsheet

Everything in this folder was rebuilt from the screenshot of the **CURRENT TECHNOLOGY**
sheet (rows 40–145 = 106 records). Nothing was typed in by hand from a second source:
the sheet is the only source, and every derived value is traceable back to its row number.

## Deliverables

| File | What it is |
|---|---|
| `NW_Site_Network_Pack.zip` | **One-click bundle** — the KMZ, the Excel report, the KML, the CSV and this README in a single download. |
| `NW_Network_Sites.kmz` | Google Earth map. 20 high sites, each with a **5 km coverage radius**; 86 connected sites; a link line from every connected site to the high site that serves it. Five switchable layers, popups with the full detail per site. |
| `NW_Network_Site_Database.xlsx` | The improved spreadsheet / report — 7 sheets, print-ready A4 landscape with headers, footers and page numbers. |
| `NW_Network_Sites.kml` | Same map as plain KML (re-import into Google Earth, My Maps, QGIS, ArcGIS). |
| `nw_sites_clean.csv` | Cleaned flat data set for any other system. |
| `index.html` | Static download page (big download buttons at the top) plus a summary and a coverage panel for every town. |
| `source_sites.csv` | The captured source rows, verbatim, including the errors. |
| `build.py` | Regenerates every file above from `source_sites.csv`. |

Regenerate with:

```bash
pip install numpy openpyxl
python3 network/build.py
```

## The workbook

1. **Overview** — headline numbers, what is in the workbook, method, legend.
2. **Sites** — all 106 sites: classification, bearer, coordinates, serving high site,
   the distance stated in the sheet, the distance recalculated from the coordinates, the
   variance, the 5 km check, a data-quality note per site, its source row, and a Google Maps link.
3. **High Sites** — the 20 high sites with connected / inside / outside counts, the farthest
   connected site, the coverage radius and any remote links noted in the sheet margin.
4. **Towns** — sites, coverage and bearers per town, plus a chart.
5. **Remote Links** — the 14 long-range links listed in the margin of the source sheet.
6. **Data Quality** — all 32 findings and the action taken for each.
7. **Source (as captured)** — rows 40–145 exactly as they appear in the original, for audit.

## The map (KMZ)

| Layer | Contents |
|---|---|
| 01 – High Sites | 20 high sites, grouped by town |
| 02 – 5 km coverage radius | one 5 km ring per high site, true size on the ground |
| 03 – Connected sites | 86 connected sites, grouped by town (dark red = outside its 5 km ring) |
| 04 – Links | a line from every connected site to its high site (grey inside 5 km, red dashed outside) |
| 05 – No high site listed | the three Delareyville sites the sheet gives no high site for |

## Numbers

- 106 sites · 20 high sites · 86 connected sites · 19 towns
- 74 connected sites fall **inside** their high site's 5 km radius, **9 outside**, 3 have no high site listed.
- 9 connected sites outside the radius: Jericho Library (8.97 km), Chaneng Library (8.69),
  Rustenburg Office of the Premier (7.81), Lethlabile Library (7.25), Lethlabile Education (7.23),
  Ganyesa Public Works (6.32), Rustenburg Arts Culture Sports & Recreation (5.97),
  Mmakau Library (5.90), Mmabatho Head Office Library & Archives (5.16).
  The Lethlabile / Jericho trio is served by Lethabong Maboloka, whose own coordinates had to be
  estimated — those three distances come straight from the sheet, so the 7–9 km hop is real.

## What the source sheet gets wrong (all listed on the Data Quality sheet)

**High severity — confirm on site**

- **Brits Library** and **Lethabong Maboloka Library** are high sites with no coordinates.
  Both were fixed from their own three connected sites and the distances stated in the sheet.
  A closed-form solve has a mirror ambiguity here, so the fix is the point that actually
  reproduces the measured distances: mean residual **0.03 km** for Brits Library
  (0.23 / 1.49 / 3.21 km against 0.21 / 1.46 / 3.25 km) and **0.01 km** for Lethabong Maboloka
  (7.25 / 7.23 / 8.97 km against 7.26 / 7.21 / 8.97 km — and ~1 km from Maboloka village).
  Both are marked *Estimated* everywhere (KML popup, both workbook columns).
- **Lehurutse Library** latitude is entered as **+25.486423**, i.e. north of the equator.
  Corrected to -25.486423 to match its own connected sites (0.7–1.1 km away).
- **Rustenburg Office of the Premier** longitude is **26.181303**, which plots ~108 km west of
  its own town — a duplicated digit from row 112 (Lichtenburg Weigh Bridge). Corrected to
  27.181303, which agrees with the sheet's own 2.90 km distance, and flagged.

**Medium — data inconsistencies**

- **Delareyville** (3 sites) has distances but no high site anywhere in the sheet, so those
  sites are left unassigned rather than given an invented one.
- **Bloemhof – Coverdale Clinic** is plotted at exactly the same coordinates as the Bloemhof
  high site but is listed 1.53 km away.
- **Ganyesa Hospital (6.35 km)** and **Ganyesa Public Works (2.30 km)** look **transposed** —
  their coordinates say 2.28 km and 6.32 km respectively.
- 5 further sites disagree with their stated distance by more than 1 km
  (Bloemhof Agriculture & Rural, Bloemhof Library, Taung Agriculture & Rural,
  Rustenburg Office of the Premier, and the Ganyesa pair above).

**Low**

- `Radio Linnk` (row 77) and `Radio Radio` (row 141) → `Radio Link`.
- `Mogwase Agriculture & Rural Development (Radio Distrib` — name truncated in the source.
- `Mahikeng` is spelled `Mahikeng` in the town column but `Mafikeng` in the high-site name —
  both left as they are so they still match the source.
