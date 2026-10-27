#!/usr/bin/env python3
"""
Build the deliverable network pack for the "CURRENT TECHNOLOGY" site database.

Inputs   : network/source_sites.csv   (verbatim capture of the source sheet, rows 40-145)
Outputs  : network/NW_Network_Sites.kmz
           network/NW_Network_Sites.kml
           network/NW_Network_Site_Database.xlsx
           network/nw_sites_clean.csv
           network/index.html          (static download / summary page)

Requires : numpy, openpyxl
Usage    : python3 network/build.py
"""

import csv
import math
import os
import re
import zipfile
from datetime import datetime

import numpy as np
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "source_sites.csv")

RADIUS_KM = 5.0          # required coverage radius around every high site
EARTH_R = 6371.0088

# --------------------------------------------------------------------------
# 1. Load the verbatim capture
# --------------------------------------------------------------------------


def load_source():
    rows = []
    with open(SRC, newline="", encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            def num(key):
                v = (r[key] or "").strip()
                return float(v) if v else None

            rows.append(
                {
                    "row_ref": int(r["row_ref"]),
                    "name": r["site_name"].strip(),
                    "town": r["town"].strip(),
                    "link_raw": r["link_type_raw"].strip(),
                    "klass": r["classification"].strip(),
                    "lat_raw": num("latitude_raw"),
                    "lon_raw": num("longitude_raw"),
                    "dist_raw": num("distance_km_raw"),
                    "comment_raw": (r["comment_raw"] or "").strip(),
                }
            )
    return rows


# --------------------------------------------------------------------------
# 2. Geometry helpers
# --------------------------------------------------------------------------


def haversine(lat1, lon1, lat2, lon2):
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = p2 - p1
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * EARTH_R * math.asin(math.sqrt(a))


def destination(lat, lon, bearing_deg, dist_km):
    br = math.radians(bearing_deg)
    p1 = math.radians(lat)
    l1 = math.radians(lon)
    dr = dist_km / EARTH_R
    p2 = math.asin(math.sin(p1) * math.cos(dr) + math.cos(p1) * math.sin(dr) * math.cos(br))
    l2 = l1 + math.atan2(
        math.sin(br) * math.sin(dr) * math.cos(p1),
        math.cos(dr) - math.sin(p1) * math.sin(p2),
    )
    return math.degrees(p2), (math.degrees(l2) + 540) % 360 - 180


def circle_ring(lat, lon, radius_km, steps=72):
    return [destination(lat, lon, i * 360.0 / steps, radius_km) for i in range(steps)]


def trilaterate(points):
    """Best-fit point for (lat, lon, distance_km) triples.

    A closed-form least-squares solve has a mirror ambiguity when the circles are
    nearly collinear, and it can land on the wrong branch.  A coarse grid search over
    the area the circles cover, followed by progressive refinement, always returns the
    point that actually agrees with the measured distances.
    """

    def cost(lat, lon):
        return sum((haversine(lat, lon, a, b) - d) ** 2 for a, b, d in points)

    lat_c = sum(p[0] for p in points) / len(points)
    lon_c = sum(p[1] for p in points) / len(points)
    span = max(p[2] for p in points)
    dlat = span / 80.0
    dlon = span / (80.0 * max(math.cos(math.radians(lat_c)), 0.2))

    lat0, lat1 = lat_c - dlat, lat_c + dlat
    lon0, lon1 = lon_c - dlon, lon_c + dlon
    step = max((lat1 - lat0) / 90.0, 0.0005)
    best = (cost(lat_c, lon_c), lat_c, lon_c)
    lat = lat0
    while lat <= lat1:
        lon = lon0
        while lon <= lon1:
            c = cost(lat, lon)
            if c < best[0]:
                best = (c, lat, lon)
            lon += step
        lat += step

    # refine
    _, lat, lon = best
    for div in (10, 10, 10):
        step_lat, step_lon = dlat / (90.0 * div), dlon / (90.0 * div)
        improved = True
        while improved:
            improved = False
            for dlat2, dlon2 in (
                (-step_lat, 0), (step_lat, 0), (0, -step_lon), (0, step_lon),
                (step_lat, step_lon), (-step_lat, -step_lon),
                (step_lat, -step_lon), (-step_lat, step_lon),
            ):
                c = cost(lat + dlat2, lon + dlon2)
                if c < cost(lat, lon):
                    lat, lon, improved = lat + dlat2, lon + dlon2, True
    residual = math.sqrt(cost(lat, lon) / len(points))
    return lat, lon, residual


# --------------------------------------------------------------------------
# 3. Structure of the network: which CPE hangs off which high site
#    (derived from the town grouping + the distance column in the source)
# --------------------------------------------------------------------------

ASSIGNMENTS = {
    43: range(44, 49),
    49: range(50, 53),
    53: range(54, 61),
    61: range(62, 71),
    71: range(72, 75),
    75: range(76, 81),
    81: range(82, 84),
    84: range(85, 88),      # Mothutlung (Brits) - Damonsville, Mmakau, Mothutlung Agric
    88: range(89, 92),      # Brits Library - town centre CPEs
    92: range(93, 96),      # Lethabong Maboloka - Lethlabile / Jericho
    96: range(97, 99),
    99: range(100, 103),
    103: range(104, 111),
    111: range(112, 116),
    116: range(117, 121),
    121: range(122, 125),
    125: range(126, 128),
    128: range(129, 137),
    137: range(138, 140),
    140: range(141, 146),
}

# Column K notes in the source sheet: long-range links from a high site to an
# out-lying community site (target coordinates are not in the source sheet).
REMOTE_LINKS = [
    (53, "Pudimoe Library", 15.6),
    (54, "Pudimong CHC", 15.7),
    (61, "Amalia Library", 27.7),
    (71, "Devondale Library", None),
    (75, "Tlaping Library", 10.22),
    (76, "Southey Library", 25.8),
    (81, "Khunwana Library", 23.3),
    (82, "Madibogopan Library", 25.0),
    (84, "Hebron Library", 14.23),
    (88, "Sonop Old Age", 8.23),
    (89, "Sonop Center for Children", 8.69),
    (128, "Boitekong Health Center", 7.44),
    (129, "Boitekong Library", 7.97),
    (140, "Matlwang Library", 13.23),
]

LINK_FIXES = {"Radio Linnk": "Radio Link", "Radio Radio": "Radio Link"}
NAME_FIXES = {
    "Mogwase Agriculture & Rural Development (Radio Distrib": (
        "Mogwase Agriculture & Rural Development (Radio Distribution)"
    )
}


def build():
    rows = load_source()
    by_row = {r["row_ref"]: r for r in rows}
    issues = []          # (row_ref, site, severity, field, issue, action)

    # ---- 3a. clean obvious transcription / data-entry faults -------------
    for r in rows:
        if r["link_raw"] in LINK_FIXES:
            issues.append(
                (r["row_ref"], r["name"], "Low", "Link Type",
                 f'"{r["link_raw"]}" is a typo', f'normalised to "{LINK_FIXES[r["link_raw"]]}"')
            )
            r["link"] = LINK_FIXES[r["link_raw"]]
        else:
            r["link"] = r["link_raw"]

        if r["name"] in NAME_FIXES:
            issues.append(
                (r["row_ref"], r["name"], "Low", "Site Name",
                 "name truncated in source", "completed to " + NAME_FIXES[r["name"]])
            )
            r["name_clean"] = NAME_FIXES[r["name"]]
        else:
            r["name_clean"] = r["name"]

    # latitude entered as a positive value (north) for a Southern-Hemisphere site
    r = by_row[96]
    if r["lat_raw"] and r["lat_raw"] > 0:
        issues.append(
            (96, r["name"], "High", "Latitude",
             "+25.486423 is north of the equator", "sign corrected to -25.486423")
        )
        r["lat_fix"] = -abs(r["lat_raw"])

    # longitude that puts the site ~108 km outside its own town
    r = by_row[134]
    hs = by_row[128]
    true_d = haversine(hs["lat_raw"], hs["lon_raw"], r["lat_raw"], r["lon_raw"])
    if true_d > 50:
        issues.append(
            (134, r["name"], "High", "Longitude",
             f"26.181303 plots {true_d:.0f} km west of East End high site",
             "corrected to 27.181303 - VERIFY on site")
        )
        r["lon_fix"] = 27.181303

    # ---- 3b. estimate the two high sites that have no coordinates --------
    estimates = {}
    for hs_ref, cpe_refs in ASSIGNMENTS.items():
        hs = by_row[hs_ref]
        if hs["lat_raw"] is not None:
            continue
        pts = [
            (by_row[x]["lat_raw"], by_row[x]["lon_raw"], by_row[x]["dist_raw"])
            for x in cpe_refs
            if by_row[x]["lat_raw"] is not None and by_row[x]["dist_raw"]
        ]
        lat, lon, rres = trilaterate(pts)
        estimates[hs_ref] = (lat, lon, rres)
        issues.append(
            (hs_ref, hs["name"], "High", "Coordinates",
             "latitude/longitude blank in source",
             f"estimated {lat:.6f}, {lon:.6f} from its {len(pts)} connected sites and "
             f"their stated distances (mean residual {rres:.2f} km) - VERIFY on site")
        )

    # ---- 3c. effective coordinates ---------------------------------------
    for r in rows:
        lat = r.get("lat_fix", r["lat_raw"])
        lon = r.get("lon_fix", r["lon_raw"])
        if lat is None and r["row_ref"] in estimates:
            lat, lon = estimates[r["row_ref"]][0], estimates[r["row_ref"]][1]
        r["lat"], r["lon"] = lat, lon
        if r["row_ref"] in estimates:
            r["coord_status"] = "Estimated (trilateration)"
        elif "lat_fix" in r or "lon_fix" in r:
            r["coord_status"] = "Corrected"
        else:
            r["coord_status"] = "Source"

    # ---- 3d. assign CPE -> high site -------------------------------------
    hs_of = {}
    for hs_ref, cpe_refs in ASSIGNMENTS.items():
        for x in cpe_refs:
            hs_of[x] = hs_ref
    for r in rows:
        if r["klass"] == "High Site":
            r["hs_ref"] = r["row_ref"]
        else:
            r["hs_ref"] = hs_of.get(r["row_ref"])

    for r in rows:
        if r["hs_ref"] is None:
            # CPE listed with a distance but its serving high site is not in the sheet
            r["hs_name"] = ""
            r["calc_dist"] = None
            r["variance"] = None
            issues.append(
                (r["row_ref"], r["name"], "Medium", "High Site",
                 "no high site listed for this town in the source sheet",
                 "left unassigned - confirm which high site serves it")
            )
            continue
        hs = by_row[r["hs_ref"]]
        r["hs_name"] = hs["name_clean"]
        if r["klass"] == "High Site":
            r["calc_dist"] = 0.0
            r["variance"] = None
        else:
            d = haversine(hs["lat"], hs["lon"], r["lat"], r["lon"])
            r["calc_dist"] = d
            r["variance"] = (d - r["dist_raw"]) if r["dist_raw"] is not None else None

    # ---- 3e. flag distance outliers / duplicate coordinates --------------
    for r in rows:
        if (
            r["klass"] != "High Site"
            and r["variance"] is not None
            and abs(r["variance"]) > 1.0
            and by_row[r["hs_ref"]]["coord_status"] == "Source"
        ):
            issues.append(
                (r["row_ref"], r["name"], "Medium", "Distance",
                 f"sheet says {r['dist_raw']:.2f} km, coordinates give {r['calc_dist']:.2f} km",
                 "coordinates used for mapping - VERIFY")
            )

    # two CPEs of the same town whose stated distances look transposed
    cpes = [r for r in rows if r["klass"] != "High Site" and r["dist_raw"] and r["calc_dist"]]
    flagged = set()
    for a in cpes:
        for b in cpes:
            if a is b or a["hs_ref"] != b["hs_ref"]:
                continue
            if abs(a["calc_dist"] - b["dist_raw"]) < 0.15 and abs(b["calc_dist"] - a["dist_raw"]) < 0.15:
                if abs(a["calc_dist"] - a["dist_raw"]) > 1.0 and (a["row_ref"], b["row_ref"]) not in flagged:
                    flagged.add((a["row_ref"], b["row_ref"]))
                    flagged.add((b["row_ref"], a["row_ref"]))
                    issues.append(
                        (a["row_ref"], a["name"], "Medium", "Distance",
                         f"stated {a['dist_raw']:.2f} km but {a['calc_dist']:.2f} km by "
                         f"coordinates - looks transposed with row {b['row_ref']}",
                         "coordinates used for mapping - VERIFY")
                    )

    seen = {}
    for r in rows:
        if r["lat"] is None:
            continue
        key = (round(r["lat"], 6), round(r["lon"], 6))
        if key in seen:
            issues.append(
                (r["row_ref"], r["name"], "Medium", "Coordinates",
                 f"identical to row {seen[key]}", "plotted at same point - VERIFY")
            )
        else:
            seen[key] = r["row_ref"]

    # ---- 3f. coverage check ----------------------------------------------
    for r in rows:
        if r["klass"] == "High Site":
            r["in_radius"] = True
            r["coverage"] = "High site (5 km coverage)"
        elif r["calc_dist"] is None:
            r["in_radius"] = None
            r["coverage"] = "No high site listed"
        else:
            r["in_radius"] = r["calc_dist"] <= RADIUS_KM
            r["coverage"] = "Inside 5 km" if r["in_radius"] else "Outside 5 km"

    # keep one message per (row, field, severity) - the later, more specific one wins
    dedup = {}
    for it in issues:
        dedup[(it[0], it[3], it[2])] = it
    issues = list(dedup.values())
    issues.sort(key=lambda i: ({"High": 0, "Medium": 1, "Low": 2}[i[2]], i[0]))

    # ---- 3g. per-site extras used by the KML/KMZ and the workbook ---------
    for r in rows:
        r["_flags"] = [
            f"{it[3]}: {it[4]} &rarr; {it[5]}" for it in issues if it[0] == r["row_ref"]
        ]
        r["_remote"] = []
    for src_row, target, km in REMOTE_LINKS:
        src = by_row[src_row]
        hs = src if src["klass"] == "High Site" else by_row.get(src["hs_ref"])
        if hs is None:
            continue
        label = target + (f" ({km:g} km)" if km else "")
        hs["_remote"].append(label)
        issues.append(
            (src_row, target, "Info", "Remote link",
             f"remote link to {label} noted against this high site",
             "target coordinates not in the source sheet - shown as a note only")
        )
    for r in rows:
        if r["klass"] != "High Site":
            continue
        members = [c for c in rows if c["klass"] != "High Site" and c["hs_ref"] == r["row_ref"]]
        r["_members"] = members
        r["_connected"] = len(members)
        r["_inside"] = sum(1 for c in members if c["in_radius"])
        r["_outside"] = sum(1 for c in members if c["in_radius"] is False)
        r["_farthest"] = max((c["calc_dist"] for c in members if c["calc_dist"]), default=None)

    return rows, by_row, issues


# --------------------------------------------------------------------------
# 4. KML / KMZ
# --------------------------------------------------------------------------

# KML colours are aabbggrr
STYLES = """
  <Style id="hs">
    <IconStyle><color>ff1d1dff</color><scale>1.15</scale>
      <Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-circle.png</href></Icon>
    </IconStyle>
    <LabelStyle><color>ff000000</color><scale>0.85</scale></LabelStyle>
    <BalloonStyle><text>$[description]</text><bgColor>ffffffff</bgColor></BalloonStyle>
  </Style>
  <Style id="cpe">
    <IconStyle><color>fff08000</color><scale>0.75</scale>
      <Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href></Icon>
    </IconStyle>
    <LabelStyle><color>ff666666</color><scale>0.6</scale></LabelStyle>
    <BalloonStyle><text>$[description]</text><bgColor>ffffffff</bgColor></BalloonStyle>
  </Style>
  <Style id="cpe-off">
    <IconStyle><color>ff3333cc</color><scale>0.85</scale>
      <Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href></Icon>
    </IconStyle>
    <LabelStyle><color>ff2222aa</color><scale>0.7</scale></LabelStyle>
    <BalloonStyle><text>$[description]</text><bgColor>ffffffff</bgColor></BalloonStyle>
  </Style>
  <Style id="cover">
    <LineStyle><color>ff1d6bff</color><width>2</width></LineStyle>
    <PolyStyle><color>33006bff</color></PolyStyle>
  </Style>
  <Style id="link">
    <LineStyle><color>8c3c3c3c</color><width>1.6</width></LineStyle>
  </Style>
  <Style id="link-off">
    <LineStyle><color>ff1d1dc8</color><width>1.8</width></LineStyle>
  </Style>
"""


def esc(text):
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def description(r):
    rows = [
        ("Site", r["name_clean"]),
        ("Town", r["town"]),
        ("Type", r["klass"]),
        ("Bearer / link type", r["link"]),
    ]
    if r["klass"] == "High Site":
        rows.append(("Coverage radius", f"{RADIUS_KM:.0f} km"))
        rows.append(("Connected sites", str(r["_connected"])))
        if r.get("_remote"):
            rows.append(("Remote links (sheet note)", "<br>".join(r["_remote"])))
    else:
        rows.append(("Served by high site", r["hs_name"] or "not listed in source sheet"))
        if r["dist_raw"] is not None:
            rows.append(("Distance (source sheet)", f"{r['dist_raw']:.2f} km"))
        if r["calc_dist"] is not None:
            rows.append(("Distance (from coordinates)", f"{r['calc_dist']:.2f} km"))
        rows.append(("Coverage", r["coverage"]))
    rows.append(("Coordinates", f"{r['lat']:.6f}, {r['lon']:.6f}"))
    rows.append(("Coordinate status", r["coord_status"]))
    if r.get("_flags"):
        rows.append(("Data quality", "<br>".join(r["_flags"])))
    rows.append(("Source row", f"row {r['row_ref']} of the CURRENT TECHNOLOGY sheet"))
    body = "".join(
        f'<tr><td style="padding:2px 8px 2px 0;color:#666;white-space:nowrap">{esc(k)}</td>'
        f'<td style="padding:2px 0">{v}</td></tr>'
        for k, v in rows
    )
    return f'<![CDATA[<table style="font-family:Roboto,Arial,sans-serif;font-size:12px">{body}</table>]]>'


def placemark(r, style, prefix=""):
    return (
        "      <Placemark>\n"
        f"        <name>{esc(prefix + r['name_clean'])}</name>\n"
        f"        <styleUrl>#{style}</styleUrl>\n"
        f"        <description>{description(r)}</description>\n"
        "        <ExtendedData>\n"
        f'          <Data name="town"><value>{esc(r["town"])}</value></Data>\n'
        f'          <Data name="type"><value>{esc(r["klass"])}</value></Data>\n'
        f'          <Data name="link_type"><value>{esc(r["link"])}</value></Data>\n'
        f'          <Data name="high_site"><value>{esc(r["hs_name"] or "")}</value></Data>\n'
        f'          <Data name="coord_status"><value>{esc(r["coord_status"])}</value></Data>\n'
        "        </ExtendedData>\n"
        "        <Point><coordinates>"
        f"{r['lon']:.6f},{r['lat']:.6f},0</coordinates></Point>\n"
        "      </Placemark>\n"
    )


def coverage_placemark(hs):
    ring = circle_ring(hs["lat"], hs["lon"], RADIUS_KM)
    ring.append(ring[0])
    coords = " ".join(f"{lon:.6f},{lat:.6f},0" for lat, lon in ring)
    return (
        "      <Placemark>\n"
        f"        <name>{esc('5 km coverage - ' + hs['name_clean'])}</name>\n"
        "        <styleUrl>#cover</styleUrl>\n"
        f"        <description><![CDATA[<div style=\"font-family:Roboto,Arial,sans-serif;font-size:12px\">"
        f"<b>{RADIUS_KM:.0f} km</b> coverage radius around the high site "
        f"<b>{esc(hs['name_clean'])}</b> ({esc(hs['town'])}).<br>"
        f"Connected sites inside/outside this radius: {hs['_inside']} in / {hs['_outside']} out "
        f"of {hs['_connected']}.</div>]]></description>\n"
        "        <Polygon><tessellate>1</tessellate><altitudeMode>clampToGround</altitudeMode>\n"
        "          <outerBoundaryIs><LinearRing><coordinates>"
        f"{coords}</coordinates></LinearRing></outerBoundaryIs>\n"
        "        </Polygon>\n"
        "      </Placemark>\n"
    )


def link_placemark(cpe, hs):
    style = "link" if cpe["in_radius"] else "link-off"
    dist = f"{cpe['calc_dist']:.2f} km" if cpe["calc_dist"] else ""
    return (
        "      <Placemark>\n"
        f"        <name>{esc(hs['name_clean'] + ' -> ' + cpe['name_clean'])}</name>\n"
        f"        <styleUrl>#{style}</styleUrl>\n"
        f"        <description><![CDATA[{esc(hs['name_clean'])} &rarr; {esc(cpe['name_clean'])}"
        f" &middot; {esc(dist)} &middot; {esc(cpe['link'])} &middot; {esc(cpe['coverage'])}]]></description>\n"
        "        <LineString><tessellate>1</tessellate><altitudeMode>clampToGround</altitudeMode>\n"
        "          <coordinates>"
        f"{hs['lon']:.6f},{hs['lat']:.6f},0 {cpe['lon']:.6f},{cpe['lat']:.6f},0"
        "</coordinates></LineString>\n"
        "      </Placemark>\n"
    )


def kml_document(rows, hs_lookup):
    hs = [r for r in rows if r["klass"] == "High Site"]
    cpe = [r for r in rows if r["klass"] != "High Site"]
    n_out = sum(1 for r in cpe if r["in_radius"] is False)
    towns = sorted({r["town"] for r in rows})

    legend = f"""
    <h2 style="font-family:Roboto,Arial,sans-serif;margin:0 0 6px">North West site network &mdash; Current Technology</h2>
    <div style="font-family:Roboto,Arial,sans-serif;font-size:12px;line-height:1.5">
      <b>{len(rows)}</b> sites &middot; <b>{len(hs)}</b> high sites &middot; <b>{len(cpe)}</b> connected (client) sites<br>
      Every high site carries a <b>{RADIUS_KM:.0f} km</b> coverage radius. Each connected site is
      joined to the high site that serves it by a link line.<br><br>
      <b>Layers</b><br>
      01 &middot; High sites &mdash; red pins<br>
      02 &middot; 5 km coverage radius &mdash; orange rings<br>
      03 &middot; Connected sites &mdash; blue dots (darker red dots = outside the 5 km radius)<br>
      04 &middot; Links &mdash; grey lines inside 5 km, red lines outside 5 km<br>
      05 &middot; Sites with no high site in the source sheet<br><br>
      <b>Coverage:</b> {len(cpe) - n_out - sum(1 for r in cpe if r['calc_dist'] is None)}
      of {len(cpe)} connected sites fall inside their high site's 5 km radius;
      <b>{n_out}</b> fall outside.
      <br><br>Built from rows 40&ndash;145 of the &ldquo;CURRENT TECHNOLOGY&rdquo; sheet.
    </div>
    """

    out = [
        '<?xml version="1.0" encoding="UTF-8"?>\n',
        '<kml xmlns="http://www.opengis.net/kml/2.2">\n<Document>\n',
        "  <name>North West Site Network - Current Technology</name>\n",
        f"  <description><![CDATA[{legend}]]></description>\n",
        STYLES,
    ]

    def folder(name, open_flag=1):
        out.append(f'  <Folder>\n    <name>{esc(name)}</name>\n    <open>{open_flag}</open>\n')

    def close():
        out.append("  </Folder>\n")

    # 01 high sites
    folder(f"01 - High Sites ({len(hs)})")
    for town in towns:
        group = [r for r in hs if r["town"] == town]
        if not group:
            continue
        out.append(f'    <Folder><name>{esc(f"{town} ({len(group)})")}</name><open>0</open>\n')
        for r in sorted(group, key=lambda x: x["name_clean"]):
            out.append(placemark(r, "hs"))
        out.append("    </Folder>\n")
    close()

    # 02 coverage
    folder("02 - 5 km coverage radius")
    for town in towns:
        group = [r for r in hs if r["town"] == town]
        if not group:
            continue
        out.append(f'    <Folder><name>{esc(f"{town} ({len(group)})")}</name><open>0</open>\n')
        for r in sorted(group, key=lambda x: x["name_clean"]):
            out.append(coverage_placemark(r))
        out.append("    </Folder>\n")
    close()

    # 03 connected sites
    folder("03 - Connected sites")
    for town in towns:
        group = [r for r in cpe if r["town"] == town]
        if not group:
            continue
        out.append(f'    <Folder><name>{esc(f"{town} ({len(group)})")}</name><open>0</open>\n')
        for r in sorted(group, key=lambda x: x["name_clean"]):
            style = "cpe" if r["in_radius"] else "cpe-off"
            out.append(placemark(r, style))
        out.append("    </Folder>\n")
    close()

    # 04 links
    folder("04 - Links (high site to connected site)", 0)
    for town in towns:
        group = [r for r in cpe if r["town"] == town and r["hs_ref"]]
        if not group:
            continue
        out.append(f'    <Folder><name>{esc(f"{town} ({len(group)})")}</name><open>0</open>\n')
        for r in sorted(group, key=lambda x: x["name_clean"]):
            out.append(link_placemark(r, hs_lookup[r["hs_ref"]]))
        out.append("    </Folder>\n")
    close()

    # 05 unassigned
    orphans = [r for r in cpe if r["hs_ref"] is None]
    if orphans:
        folder("05 - No high site listed in source sheet")
        for town in sorted({r["town"] for r in orphans}):
            out.append(f'    <Folder><name>{esc(town)}</name><open>0</open>\n')
            for r in orphans:
                if r["town"] == town:
                    out.append(placemark(r, "cpe-off"))
            out.append("    </Folder>\n")
        close()

    out.append("</Document>\n</kml>\n")
    return "".join(out)


def write_kml(rows, path):
    hs_lookup = {r["row_ref"]: r for r in rows if r["klass"] == "High Site"}
    txt = kml_document(rows, hs_lookup)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(txt)
    return txt


# --------------------------------------------------------------------------
# 5. XLSX workbook
# --------------------------------------------------------------------------

NAVY = "1F3864"
MID_BLUE = "2E5C9A"
INK = "262626"
GREY = "6F6F6F"
LINE = "D6DCE5"
GREEN = "E2EFDA"
RED_FILL = "FCE4E4"
AMBER = "FFF2CC"

# colours carried over from the source sheet (strong = high site, tint = CPE)
TOWN_COLOURS = {
    "Delareyville": "FFFFFF",
    "Bloemhof": "2F5597",
    "Christiana": "C6E0B4",
    "Taung": "92D050",
    "Schweizer-Reneke": "F4B183",
    "Vryburg": "CCFFFF",
    "Ganyesa": "E5B8E5",
    "Setlagole": "943634",
    "Brits": "00FF00",
    "Lehurutse": "7030A0",
    "Itsoseng": "FFFF00",
    "Mahikeng": "92D050",
    "Lichtenburg": "2F5597",
    "Zeerust": "943634",
    "Mogwase": "CCFFFF",
    "Moses Kutane": "7030A0",
    "Rustenburg": "F4B183",
    "Reivello": "00FF00",
    "Potchefstroom": "7030A0",
}


def tint(hex_colour, amount=0.72):
    """Blend a colour towards white; amount = how much white to add."""
    r, g, b = (int(hex_colour[i : i + 2], 16) for i in (0, 2, 4))
    return "{:02X}{:02X}{:02X}".format(
        int(r + (255 - r) * amount), int(g + (255 - g) * amount), int(b + (255 - b) * amount)
    )


def text_colour(hex_colour):
    r, g, b = (int(hex_colour[i : i + 2], 16) for i in (0, 2, 4))
    return "FFFFFF" if (0.299 * r + 0.587 * g + 0.114 * b) < 140 else INK


THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def print_setup(ws, title, landscape=True, fit_width=1):
    """Make every sheet behave like a page of a report when printed / exported to PDF."""
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.55
    ws.page_margins.header = ws.page_margins.footer = 0.25
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = fit_width
    ws.page_setup.fitToHeight = 0
    ws.oddHeader.left.text = "North West site network - Current Technology"
    ws.oddHeader.left.size = 8
    ws.oddHeader.left.color = "6F6F6F"
    ws.oddHeader.right.text = title
    ws.oddHeader.right.size = 8
    ws.oddHeader.right.color = "6F6F6F"
    ws.oddFooter.left.text = "Built from the CURRENT TECHNOLOGY sheet, rows 40-145"
    ws.oddFooter.left.size = 8
    ws.oddFooter.left.color = "6F6F6F"
    ws.oddFooter.right.text = "Page &P of &N"
    ws.oddFooter.right.size = 8
    ws.oddFooter.right.color = "6F6F6F"


def style_header(ws, row, ncols, height=32):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, color="FFFFFF", size=10.5, name="Calibri")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = BOX
    ws.row_dimensions[row].height = height


def title_block(ws, title, subtitle, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(bold=True, size=16, color=NAVY, name="Calibri")
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 26
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=2 if False else 1, value=subtitle)
    c.font = Font(size=9.5, color=GREY, name="Calibri")
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 16


def maps_link(r):
    return f"https://www.google.com/maps/search/?api=1&query={r['lat']:.6f},{r['lon']:.6f}"


def write_xlsx(rows, issues, path):
    hs = [r for r in rows if r["klass"] == "High Site"]
    cpe = [r for r in rows if r["klass"] != "High Site"]
    towns = sorted({r["town"] for r in rows})
    inside = [r for r in cpe if r["in_radius"]]
    outside = [r for r in cpe if r["in_radius"] is False]
    orphan = [r for r in cpe if r["hs_ref"] is None]

    wb = Workbook()

    # ------------------------------------------------------------------ cover
    ws = wb.active
    ws.title = "Overview"
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.showGridLines = False
    width = 9
    for col in "ABCDEFGHI":
        ws.column_dimensions[col].width = 15
    title_block(
        ws,
        "North West Site Network  -  Improved Site Database",
        "Rebuilt from the 'CURRENT TECHNOLOGY' sheet (rows 40-145) - current technology site information database. "
        "Prepared " + datetime.now().strftime("%d %B %Y") + ".",
        width,
    )
    ws.row_dimensions[1].height = 30

    kpis = [
        ("Total sites", len(rows), "all records in the source sheet"),
        ("High sites", len(hs), "each drawn with a 5 km coverage radius"),
        ("Connected sites", len(cpe), "CPEs / client sites hanging off a high site"),
        ("Inside 5 km", len(inside), "connected sites within the coverage radius"),
        ("Outside 5 km", len(outside), "connected sites beyond the coverage radius"),
        ("Towns covered", len(towns), "district towns in the database"),
        ("High sites without coordinates", len([h for h in hs if h["coord_status"] != "Source"]),
         "estimated in this workbook - verify on site"),
        ("Sites with no high site", len(orphan), "Delareyville - no high site in the source sheet"),
    ]
    r0 = 4
    ws.cell(row=r0, column=1, value="HEADLINE NUMBERS").font = Font(bold=True, size=11, color=NAVY)
    for i, (label, value, note) in enumerate(kpis):
        rr = r0 + 1 + i
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=2)
        c = ws.cell(row=rr, column=1, value=label)
        c.font = Font(bold=True, size=10.5, color=INK)
        c.fill = PatternFill("solid", fgColor="EEF2F8")
        c.alignment = Alignment(vertical="center", indent=1)
        v = ws.cell(row=rr, column=3, value=value)
        v.font = Font(bold=True, size=12, color=NAVY)
        v.alignment = Alignment(horizontal="center", vertical="center")
        v.fill = PatternFill("solid", fgColor="EEF2F8")
        ws.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=9)
        n = ws.cell(row=rr, column=4, value=note)
        n.font = Font(size=9.5, color=GREY)
        n.alignment = Alignment(vertical="center", indent=1)
        for col in range(1, 10):
            ws.cell(row=rr, column=col).border = BOX
        ws.row_dimensions[rr].height = 18

    r1 = r0 + len(kpis) + 3
    ws.cell(row=r1, column=1, value="WHAT IS IN THIS WORKBOOK").font = Font(bold=True, size=11, color=NAVY)
    contents = [
        ("Sites", "Every site: name, town, type, coordinates, serving high site, both distances, 5 km check, flags"),
        ("High Sites", "The 20 high sites with connected-site counts and the 5 km coverage radius"),
        ("Towns", "Sites per town and coverage per town"),
        ("Remote Links", "Long-range links noted in the margin of the source sheet"),
        ("Data Quality", "Every fault found in the source sheet and what was done about it"),
        ("Source (as captured)", "The original rows 40-145, unchanged, for audit"),
    ]
    for i, (sheet, desc) in enumerate(contents):
        rr = r1 + 1 + i
        c = ws.cell(row=rr, column=1, value=sheet)
        c.font = Font(bold=True, size=10, color=MID_BLUE)
        c.alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=2)
        d = ws.cell(row=rr, column=3, value=desc)
        ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=9)
        d.font = Font(size=9.5, color=INK)
        d.alignment = Alignment(vertical="center", indent=1)

    r2 = r1 + len(contents) + 3
    ws.cell(row=r2, column=1, value="HOW THE MAPPING WAS DONE").font = Font(bold=True, size=11, color=NAVY)
    method = [
        "Every high site is drawn with a 5 km radius ring (the coverage requirement in the brief).",
        "Each connected site is linked by a line to the high site that serves it; lines that leave the 5 km ring are drawn in red.",
        "Distances were recalculated from the coordinates (great-circle) and compared with the distances in the sheet; the difference is shown in the Sites sheet.",
        "Two high sites have no coordinates in the source (Brits Library, Lethabong Maboloka Library). Both were fixed from their own connected sites and the distances stated in the sheet: the fit reproduces every stated distance to within 0.05 km (mean residual 0.03 km and 0.01 km), and the Lethabong fix lands ~1 km from Maboloka village. Both are flagged - confirm on site.",
        "Lehurutse Library had a positive latitude (+25.486423). It was corrected to -25.486423.",
        "Rustenburg Office of the Premier had a longitude of 26.181303, which plots ~108 km west of its own town; corrected to 27.181303 and flagged.",
        "Three Delareyville sites have distances but no high site anywhere in the sheet. They are left unassigned rather than invented.",
        "The companion map file is NW_Network_Sites.kmz (open it in Google Earth). It carries the same 106 sites, the 5 km rings and a link line from every connected site to its high site.",
    ]
    for i, m in enumerate(method):
        rr = r2 + 1 + i
        c = ws.cell(row=rr, column=1, value="-  " + m)
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=9)
        c.font = Font(size=9.5, color=INK)
        c.alignment = Alignment(vertical="top", wrap_text=True, indent=1)
        ws.row_dimensions[rr].height = 26

    r3 = r2 + len(method) + 3
    ws.cell(row=r3, column=1, value="LEGEND").font = Font(bold=True, size=11, color=NAVY)
    legend = [
        (NAVY, "FFFFFF", "High site row (strong town colour)"),
        ("F2F2F2", INK, "Connected site row (tinted town colour)"),
        (RED_FILL, "9C0006", "Outside the 5 km radius / data quality flag"),
        (GREEN, "375623", "Inside the 5 km radius"),
        (AMBER, "7F6000", "Coordinate estimated or corrected - verify on site"),
    ]
    for i, (fill, font_colour, label) in enumerate(legend):
        rr = r3 + 1 + i
        c = ws.cell(row=rr, column=1, value="")
        c.fill = PatternFill("solid", fgColor=fill)
        c.border = BOX
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=2)
        l = ws.cell(row=rr, column=3, value=label)
        ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=9)
        l.font = Font(size=9.5, color=font_colour)
        l.alignment = Alignment(vertical="center", indent=1)
    print_setup(ws, "Overview", landscape=False)

    # ------------------------------------------------------------------ sites
    ws = wb.create_sheet("Sites")
    ws.sheet_properties.tabColor = MID_BLUE
    headers = [
        ("No", 6),
        ("Site Name", 44),
        ("Town", 18),
        ("Classification", 13),
        ("Link Type (bearer)", 22),
        ("Latitude", 13),
        ("Longitude", 13),
        ("Serving High Site", 32),
        ("Distance to High Site -\nsource sheet (km)", 15),
        ("Distance to High Site -\nfrom coordinates (km)", 16),
        ("Variance\n(km)", 10),
        ("5 km coverage", 15),
        ("Coordinate status", 20),
        ("Data quality notes (click the flag)", 46),
        ("Source row", 8),
        ("Map", 10),
    ]
    title_block(
        ws,
        "Sites  -  all " + str(len(rows)) + " sites",
        "Filter or sort any column. Rows are coloured the way the source sheet colours each town (strong colour = high site).",
        len(headers),
    )
    for i, (h, w) in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
        ws.column_dimensions[ws.cell(row=4, column=i).column_letter].width = w
    style_header(ws, 4, len(headers), height=42)

    ordered = sorted(rows, key=lambda r: (r["town"], r["klass"] != "High Site", r["hs_name"], r["name_clean"]))
    r_i = 5
    for n, r in enumerate(ordered, start=1):
        base = TOWN_COLOURS.get(r["town"], "FFFFFF")
        fill = base if r["klass"] == "High Site" else tint(base)
        fontc = text_colour(base) if r["klass"] == "High Site" else INK
        vals = [
            n,
            r["name_clean"],
            r["town"],
            r["klass"],
            r["link"],
            round(r["lat"], 6),
            round(r["lon"], 6),
            r["hs_name"] or "-",
            r["dist_raw"],
            round(r["calc_dist"], 2) if r["calc_dist"] is not None else None,
            round(r["variance"], 2) if r["variance"] is not None else None,
            r["coverage"],
            r["coord_status"],
            "; ".join(re.sub("<[^>]+>", "", f) for f in r["_flags"]) or "",
            r["row_ref"],
            "Open",
        ]
        for c_i, v in enumerate(vals, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.border = BOX
            cell.font = Font(size=10, color=fontc, bold=(r["klass"] == "High Site"))
            cell.fill = PatternFill("solid", fgColor=fill)
            if c_i in (2, 3, 8, 14):
                cell.alignment = Alignment(vertical="center", indent=1, wrap_text=(c_i == 14))
            elif c_i in (6, 7):
                cell.number_format = "0.000000"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_i in (9, 10):
                cell.number_format = "0.00"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i == 11:
                cell.number_format = "+0.00;-0.00;0.00"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i == 16:
                cell.hyperlink = maps_link(r)
                cell.font = Font(size=9.5, color="0563C1", underline="single")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
        if r["coord_status"] != "Source":
            ws.cell(row=r_i, column=13).fill = PatternFill("solid", fgColor=AMBER)
        ws.row_dimensions[r_i].height = 15
        r_i += 1

    last = r_i - 1
    ws.auto_filter.ref = f"A4:P{last}"
    ws.freeze_panes = "C5"
    rng = f"L5:L{last}"
    ws.conditional_formatting.add(
        rng, CellIsRule(operator="equal", formula=['"Outside 5 km"'],
                        fill=PatternFill("solid", bgColor=RED_FILL),
                        font=Font(color="9C0006", bold=True))
    )
    ws.conditional_formatting.add(
        rng, CellIsRule(operator="equal", formula=['"Inside 5 km"'],
                        fill=PatternFill("solid", bgColor=GREEN),
                        font=Font(color="375623"))
    )
    ws.conditional_formatting.add(
        f"M5:M{last}", CellIsRule(operator="equal", formula=['"Estimated (trilateration)"'],
                                  fill=PatternFill("solid", bgColor=AMBER))
    )
    ws.print_title_rows = "4:4"
    print_setup(ws, "Sites")

    # -------------------------------------------------------------- high sites
    ws = wb.create_sheet("High Sites")
    ws.sheet_properties.tabColor = "C00000"
    headers = [
        ("No", 5),
        ("High Site", 42),
        ("Town", 18),
        ("Latitude", 13),
        ("Longitude", 13),
        ("Bearer", 22),
        ("Connected sites", 11),
        ("Inside 5 km", 10),
        ("Outside 5 km", 10),
        ("Farthest connected site (km)", 13),
        ("Coverage radius (km)", 11),
        ("Coordinate status", 20),
        ("Remote links noted in the sheet margin", 42),
        ("Map", 10),
    ]
    title_block(
        ws,
        "High sites  -  " + str(len(hs)) + " sites, each with a 5 km coverage radius",
        "Coverage radius is the ring drawn around every high site in the KMZ file.",
        len(headers),
    )
    for i, (h, w) in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
        ws.column_dimensions[ws.cell(row=4, column=i).column_letter].width = w
    style_header(ws, 4, len(headers), height=42)

    r_i = 5
    for n, r in enumerate(sorted(hs, key=lambda x: (x["town"], x["name_clean"])), start=1):
        base = TOWN_COLOURS.get(r["town"], "FFFFFF")
        vals = [
            n, r["name_clean"], r["town"], round(r["lat"], 6), round(r["lon"], 6), r["link"],
            r["_connected"], r["_inside"], r["_outside"],
            round(r["_farthest"], 2) if r["_farthest"] else None,
            RADIUS_KM, r["coord_status"], "; ".join(r["_remote"]), "Open",
        ]
        for c_i, v in enumerate(vals, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.border = BOX
            cell.font = Font(size=10, color=text_colour(base), bold=True)
            cell.fill = PatternFill("solid", fgColor=base)
            if c_i in (2, 3, 13):
                cell.alignment = Alignment(vertical="center", indent=1, wrap_text=(c_i == 13))
            elif c_i in (4, 5):
                cell.number_format = "0.000000"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c_i in (7, 8, 9, 11):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i == 10:
                cell.number_format = "0.00"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i == 14:
                cell.hyperlink = maps_link(r)
                cell.font = Font(size=9.5, color="0563C1", underline="single")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
        if r["coord_status"] != "Source":
            ws.cell(row=r_i, column=12).fill = PatternFill("solid", fgColor=AMBER)
        ws.row_dimensions[r_i].height = 15
        r_i += 1

    last = r_i - 1
    tot = r_i
    ws.cell(row=tot, column=1, value="")
    ws.merge_cells(start_row=tot, start_column=1, end_row=tot, end_column=6)
    t = ws.cell(row=tot, column=1, value="TOTAL / NETWORK")
    t.font = Font(bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    for c_i, formula in [(7, f"=SUM(G5:G{last})"), (8, f"=SUM(H5:H{last})"), (9, f"=SUM(I5:I{last})")]:
        c = ws.cell(row=tot, column=c_i, value=formula)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.auto_filter.ref = f"A4:N{last}"
    ws.freeze_panes = "C5"
    ws.print_title_rows = "4:4"
    print_setup(ws, "High Sites")

    # ------------------------------------------------------------------ towns
    ws = wb.create_sheet("Towns")
    ws.sheet_properties.tabColor = "70AD47"
    headers = [
        ("Town", 22), ("High sites", 11), ("Connected sites", 14), ("Total sites", 11),
        ("Inside 5 km", 11), ("Outside 5 km", 12), ("No high site", 12),
        ("Bearers used", 40),
    ]
    title_block(ws, "Sites and coverage per town", "Bearers are the transmission media recorded in the source sheet.", len(headers))
    for i, (h, w) in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
        ws.column_dimensions[ws.cell(row=4, column=i).column_letter].width = w
    style_header(ws, 4, len(headers))
    r_i = 5
    for town in towns:
        group = [r for r in rows if r["town"] == town]
        h = [r for r in group if r["klass"] == "High Site"]
        c = [r for r in group if r["klass"] != "High Site"]
        bearers = sorted({r["link"] for r in group})
        base = TOWN_COLOURS.get(town, "FFFFFF")
        vals = [
            town, len(h), len(c), len(group),
            sum(1 for x in c if x["in_radius"]),
            sum(1 for x in c if x["in_radius"] is False),
            sum(1 for x in c if x["hs_ref"] is None),
            ", ".join(bearers),
        ]
        for c_i, v in enumerate(vals, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.border = BOX
            cell.font = Font(size=10, color=text_colour(base) if c_i == 1 else INK, bold=(c_i == 1))
            cell.fill = PatternFill("solid", fgColor=base if c_i == 1 else "FFFFFF")
            cell.alignment = Alignment(
                vertical="center", indent=1 if c_i in (1, 8) else 0,
                horizontal="left" if c_i in (1, 8) else "center",
                wrap_text=(c_i == 8),
            )
        ws.row_dimensions[r_i].height = 16
        r_i += 1

    last = r_i - 1
    ws.auto_filter.ref = f"A4:H{last}"
    ws.freeze_panes = "B5"
    chart = BarChart()
    chart.title = "Connected sites per town"
    chart.y_axis.title = "Connected sites"
    chart.height, chart.width = 8, 20
    data = Reference(ws, min_col=3, min_row=4, max_row=last)
    cats = Reference(ws, min_col=1, min_row=5, max_row=last)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None
    chart.style = 10
    ws.add_chart(chart, "J4")
    print_setup(ws, "Towns")

    # ---------------------------------------------------------- remote links
    ws = wb.create_sheet("Remote Links")
    ws.sheet_properties.tabColor = "ED7D31"
    headers = [("High site", 40), ("Town", 18), ("Remote site (from sheet margin)", 34),
               ("Distance (km)", 12), ("Noted against source row", 16),
               ("Note", 46)]
    title_block(
        ws,
        "Remote links",
        "The source sheet lists these out-lying sites in the margin next to a high site. Their coordinates are not in the sheet, "
        "so they are recorded here (and in the high site's balloon in the KMZ) but not plotted.",
        len(headers),
    )
    for i, (h, w) in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
        ws.column_dimensions[ws.cell(row=4, column=i).column_letter].width = w
    style_header(ws, 4, len(headers))
    r_i = 5
    for src_row, target, km in REMOTE_LINKS:
        src = next(x for x in rows if x["row_ref"] == src_row)
        hsrec = src if src["klass"] == "High Site" else next(
            (x for x in rows if x["row_ref"] == src["hs_ref"]), None)
        vals = [
            hsrec["name_clean"] if hsrec else "-",
            src["town"], target, km, src_row,
            "long-range link" + ("" if km else " (distance not given)"),
        ]
        for c_i, v in enumerate(vals, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.border = BOX
            cell.font = Font(size=10)
            cell.alignment = Alignment(
                vertical="center", indent=1 if c_i in (1, 3, 6) else 0,
                horizontal="center" if c_i in (4, 5) else "left")
            if c_i == 4:
                cell.number_format = "0.00"
        ws.row_dimensions[r_i].height = 16
        r_i += 1
    print_setup(ws, "Remote links")

    # ---------------------------------------------------------- data quality
    ws = wb.create_sheet("Data Quality")
    ws.sheet_properties.tabColor = "C00000"
    headers = [("Severity", 10), ("Source row", 10), ("Site", 42), ("Field", 16),
               ("What is wrong", 62), ("What this workbook does", 56)]
    title_block(
        ws,
        "Data quality",
        "Everything found while rebuilding the sheet, and the action taken. Anything marked High should be confirmed on site.",
        len(headers),
    )
    for i, (h, w) in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
        ws.column_dimensions[ws.cell(row=4, column=i).column_letter].width = w
    style_header(ws, 4, len(headers))
    order = {"High": 0, "Medium": 1, "Low": 2, "Info": 3}
    r_i = 5
    for row_ref, site, sev, field, bad, action in sorted(issues, key=lambda i: (order[i[2]], i[0])):
        fill = {"High": RED_FILL, "Medium": AMBER, "Low": "FFFFFF", "Info": "EEF2F8"}[sev]
        for c_i, v in enumerate([sev, row_ref, site, field, bad, action], start=1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.border = BOX
            cell.font = Font(size=10, bold=(c_i == 1 and sev == "High"),
                             color="9C0006" if sev == "High" and c_i == 1 else INK)
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.alignment = Alignment(
                vertical="center", wrap_text=c_i in (5, 6, 3),
                horizontal="center" if c_i in (1, 2, 4) else "left",
                indent=1 if c_i in (3, 5, 6) else 0)
        ws.row_dimensions[r_i].height = 28
        r_i += 1
    ws.auto_filter.ref = f"A4:F{r_i - 1}"
    ws.freeze_panes = "A5"
    ws.print_title_rows = "4:4"
    print_setup(ws, "Data quality")

    # ------------------------------------------------------- source (verbatim)
    ws = wb.create_sheet("Source (as captured)")
    ws.sheet_properties.tabColor = "A6A6A6"
    headers = [("Source row", 11), ("Site Name", 44), ("Town", 18), ("Link Type", 24),
               ("Classification", 14), ("Latitude", 13), ("Longitude", 13), ("Distance (km)", 13)]
    title_block(
        ws,
        "Source sheet as captured",
        "Rows 40-145 of 'CURRENT TECHNOLOGY', exactly as they appear in the original (including the errors listed on the Data Quality sheet).",
        len(headers),
    )
    for i, (h, w) in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
        ws.column_dimensions[ws.cell(row=4, column=i).column_letter].width = w
    style_header(ws, 4, len(headers))
    r_i = 5
    for r in rows:
        vals = [r["row_ref"], r["name"], r["town"], r["link_raw"], r["klass"],
                r["lat_raw"], r["lon_raw"], r["dist_raw"]]
        for c_i, v in enumerate(vals, start=1):
            cell = ws.cell(row=r_i, column=c_i, value=v)
            cell.border = BOX
            cell.font = Font(size=10)
            cell.alignment = Alignment(vertical="center", horizontal="left" if c_i == 2 else "center",
                                       indent=1 if c_i == 2 else 0)
            if c_i in (6, 7):
                cell.number_format = "0.000000"
            if c_i == 8:
                cell.number_format = "0.00"
        ws.row_dimensions[r_i].height = 15
        r_i += 1
    ws.auto_filter.ref = f"A4:H{r_i - 1}"
    ws.freeze_panes = "B5"
    print_setup(ws, "Source rows")

    wb.active = 0
    wb.save(path)
    return path


# --------------------------------------------------------------------------
# 6. KMZ, CSV, download page
# --------------------------------------------------------------------------


def write_kmz(kml_text, path):
    doc = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<kml xmlns="http://www.opengis.net/kml/2.2">\n<Document>\n'
        "  <name>North West Site Network - Current Technology</name>\n"
        '</Document>\n</kml>\n'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("doc.kml", kml_text)
    return path


def write_zip(path, files):
    """One-click bundle: KMZ + Excel report + KML + CSV + notes."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for item in files:
            full = os.path.join(HERE, item)
            if os.path.exists(full):
                z.write(full, item)
    return path


def write_csv(rows, path):
    cols = [
        ("row_ref", "source_row"),
        ("name_clean", "site_name"),
        ("town", "town"),
        ("klass", "classification"),
        ("link", "link_type"),
        ("lat", "latitude"),
        ("lon", "longitude"),
        ("hs_name", "serving_high_site"),
        ("dist_raw", "distance_to_high_site_source_km"),
        ("calc_dist", "distance_to_high_site_calculated_km"),
        ("coverage", "five_km_coverage"),
        ("coord_status", "coordinate_status"),
        ("flags", "data_quality"),
    ]
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow([c[1] for c in cols])
        for r in sorted(rows, key=lambda x: (x["town"], x["klass"] != "High Site", x["name_clean"])):
            out = []
            for key, _ in cols:
                v = r.get(key)
                if key == "flags":
                    v = re.sub("<[^>]+>", "", "; ".join(r["_flags"]))
                if key in ("lat", "lon") and v is not None:
                    v = f"{v:.6f}"
                elif key in ("dist_raw", "calc_dist") and v is not None:
                    v = f"{v:.2f}"
                out.append("" if v is None else v)
            w.writerow(out)
    return path


def short_name(name, limit=24):
    name = name.split(" (")[0].split(" - ")[0].replace("&", "and").strip()
    if len(name) <= limit:
        return name
    cut = name[:limit].rsplit(" ", 1)[0]
    return cut.rstrip(",;") + "\u2026"


def svg_panel(rows, town):
    """One town detail panel: equirectangular in km, 5 km rings at true size."""
    group = [r for r in rows if r["town"] == town]
    hs = [r for r in group if r["klass"] == "High Site"]
    lat0 = sum(r["lat"] for r in group) / len(group)
    kx = 111.320 * math.cos(math.radians(lat0))
    ky = 110.574

    def xy(r):
        return (r["lon"] * kx, -r["lat"] * ky)

    pts = [xy(r) for r in group]
    pad = RADIUS_KM * 0.55 + 1.2
    x0 = min(p[0] for p in pts) - pad
    x1 = max(p[0] for p in pts) + pad
    y0 = min(p[1] for p in pts) - pad
    y1 = max(p[1] for p in pts) + pad
    # keep a sensible aspect so tidy little clusters do not become slivers
    min_span = RADIUS_KM * 2.6 if hs else RADIUS_KM * 1.2
    if x1 - x0 < min_span:
        grow = (min_span - (x1 - x0)) / 2
        x0, x1 = x0 - grow, x1 + grow
    if y1 - y0 < min_span:
        grow = (min_span - (y1 - y0)) / 2
        y0, y1 = y0 - grow, y1 + grow
    w, h = x1 - x0, y1 - y0
    unit = min(w, h)
    r_cpe = max(unit * 0.014, 0.12)
    r_hs = max(unit * 0.034, 0.30)
    fs = max(unit * 0.050, 0.45)

    # labels sit to the right of their high site: widen the frame until the longest
    # label fits, so nothing is ever clipped
    if hs:
        right = max(
            (xy(r)[0] + RADIUS_KM * 1.06 + len(short_name(r["name_clean"])) * fs * 0.62)
            for r in hs
        )
        x1 = max(x1, right + fs * 1.4)
    w = x1 - x0

    hs_by_ref = {r["row_ref"]: r for r in hs}
    out = [
        f'<svg viewBox="{x0:.2f} {y0:.2f} {w:.2f} {h:.2f}" preserveAspectRatio="xMidYMid meet" '
        f'role="img" aria-label="{esc(town)}: high sites with 5 km coverage and connected sites">',
        f'<rect x="{x0:.2f}" y="{y0:.2f}" width="{w:.2f}" height="{h:.2f}" fill="#f8fafd"/>',
    ]
    for r in group:
        if r["klass"] == "High Site" or r["hs_ref"] is None:
            continue
        a, b = xy(hs_by_ref[r["hs_ref"]]), xy(r)
        cls = "pl" if r["in_radius"] else "pl-out"
        out.append(
            f'<line class="{cls}" x1="{a[0]:.2f}" y1="{a[1]:.2f}" x2="{b[0]:.2f}" y2="{b[1]:.2f}"/>'
        )
    for r in hs:
        x, y = xy(r)
        out.append(f'<circle class="pring" cx="{x:.2f}" cy="{y:.2f}" r="{RADIUS_KM:.2f}"/>')
    for r in group:
        if r["klass"] == "High Site":
            continue
        x, y = xy(r)
        cls = "pcpe" if r["in_radius"] else "pcpe-out"
        tip = f'{r["name_clean"]} - {r["coverage"]}'
        if r["calc_dist"]:
            tip += f' ({r["calc_dist"]:.2f} km)'
        out.append(
            f'<circle class="{cls}" cx="{x:.2f}" cy="{y:.2f}" r="{r_cpe:.2f}">'
            f'<title>{esc(tip)}</title></circle>'
        )
    for r in hs:
        x, y = xy(r)
        out.append(
            f'<rect class="phs" x="{x - r_hs:.2f}" y="{y - r_hs:.2f}" '
            f'width="{2 * r_hs:.2f}" height="{2 * r_hs:.2f}">'
            f'<title>{esc(r["name_clean"])} - 5 km ring, {r["_connected"]} connected sites</title></rect>'
        )
        out.append(
            f'<text class="plbl" x="{x + RADIUS_KM * 1.06:.2f}" y="{y + fs * 0.36:.2f}" '
            f'font-size="{fs:.2f}">{esc(short_name(r["name_clean"]))}</text>'
        )
    out.append("</svg>")
    return "".join(out)


def write_index(rows, issues, path):
    hs = [r for r in rows if r["klass"] == "High Site"]
    cpe = [r for r in rows if r["klass"] != "High Site"]
    inside = [r for r in cpe if r["in_radius"]]
    outside = [r for r in cpe if r["in_radius"] is False]
    orphan = [r for r in cpe if r["hs_ref"] is None]
    high = [i for i in issues if i[2] == "High"]
    towns = sorted({r["town"] for r in rows})

    def size(name):
        p = os.path.join(HERE, name)
        kb = os.path.getsize(p) / 1024
        return f"{kb/1024:.1f} MB" if kb > 1024 else f"{kb:.0f} KB"

    def hs_rows():
        out = []
        for r in sorted(hs, key=lambda x: (x["town"], x["name_clean"])):
            dots = "".join(
                '<i class="dot"></i>' for _ in range(min(r["_connected"], 12))
            )
            out.append(
                f"<tr><td>{esc(r['town'])}</td><td><b>{esc(r['name_clean'])}</b></td>"
                f"<td>{esc(r['link'])}</td><td class='n'>{r['_connected']}</td>"
                f"<td class='n'>{r['_inside']}</td><td class='n'>{r['_outside']}</td>"
                f"<td class='n'>{r['_farthest']:.2f}</td>"
                f"<td><code>{r['lat']:.5f}, {r['lon']:.5f}</code></td></tr>"
            )
        return "".join(out)

    def out_rows():
        rows_html = []
        for r in sorted(outside, key=lambda x: -x["calc_dist"]):
            rows_html.append(
                f"<tr><td><b>{esc(r['name_clean'])}</b></td><td>{esc(r['town'])}</td>"
                f"<td>{esc(r['hs_name'])}</td><td class='n'>{r['calc_dist']:.2f}</td>"
                f"<td class='n'>{r['calc_dist'] - RADIUS_KM:+.2f}</td></tr>"
            )
        return "".join(rows_html)

    def town_panels():
        out = []
        for town in sorted(towns):
            group = [r for r in rows if r["town"] == town]
            n_hs = sum(1 for r in group if r["klass"] == "High Site")
            n_cpe = len(group) - n_hs
            n_out = sum(1 for r in group if r["klass"] != "High Site" and r["in_radius"] is False)
            bits = [f"{n_cpe} connected"]
            if n_hs:
                bits.append(f"{n_hs} high site" + ("s" if n_hs > 1 else ""))
                bits.append(f"{RADIUS_KM:.0f} km ring")
            else:
                bits.append("no high site in the sheet")
            if n_out:
                bits.append(f"<b class='warn'>{n_out} outside</b>")
            out.append(
                f'    <figure class="tile">\n'
                f'      <figcaption><b>{esc(town)}</b><span>{" &middot; ".join(bits)}</span></figcaption>\n'
                f'      {svg_panel(rows, town)}\n'
                f'    </figure>'
            )
        return "\n".join(out)

    panels = town_panels()

    def issue_rows():
        lo = {"High": 0, "Medium": 1, "Low": 2, "Info": 3}
        out = []
        for row_ref, site, sev, field, bad, action in sorted(issues, key=lambda i: (lo[i[2]], i[0]))[:12]:
            out.append(
                f"<tr><td><span class='tag {sev.lower()}'>{sev}</span></td>"
                f"<td>{esc(site)}</td><td>{esc(field)}</td><td>{esc(bad)}</td></tr>"
            )
        return "".join(out)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>North West site network - KMZ and improved spreadsheet</title>
<style>
  :root {{ --navy:#1f3864; --blue:#2e5c9a; --ink:#22272e; --grey:#6f7883; --line:#e2e7ee; --bg:#f6f8fb; }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:var(--bg); color:var(--ink);
         font:16px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif; }}
  .wrap {{ max-width:1120px; margin:0 auto; padding:0 22px 70px; }}
  header {{ background:linear-gradient(160deg,#16294a,#1f3864 55%,#2e5c9a); color:#fff; padding:44px 0 38px; }}
  header h1 {{ margin:0 0 8px; font-size:30px; letter-spacing:-.4px; }}
  header p {{ margin:0; color:#c9d6e8; max-width:70ch; }}
  .eyebrow {{ text-transform:uppercase; letter-spacing:2px; font-size:11px; color:#8fb0dc; margin-bottom:12px; }}
  h2 {{ font-size:19px; margin:44px 0 12px; color:var(--navy); letter-spacing:-.2px; }}
  h2 span {{ font-weight:400; color:var(--grey); font-size:14px; }}
  .kpis {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(148px,1fr)); gap:12px; margin:-26px 0 0; }}
  .kpi {{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:14px 16px;
          box-shadow:0 6px 20px rgba(20,40,70,.07); }}
  .kpi b {{ display:block; font-size:26px; color:var(--navy); line-height:1.15; }}
  .kpi span {{ font-size:12px; color:var(--grey); }}
  .band {{ margin-top:26px; background:#fff; border:1px solid var(--line); border-radius:14px;
           padding:18px 20px; display:flex; flex-wrap:wrap; gap:16px 24px; align-items:center;
           justify-content:space-between; box-shadow:0 6px 22px rgba(20,40,70,.06); }}
  .band-txt b {{ display:block; font-size:16px; color:var(--navy); }}
  .band-txt span {{ font-size:12.5px; color:var(--grey); }}
  .band-btns {{ display:flex; flex-wrap:wrap; gap:10px; }}
  .big {{ display:flex; flex-direction:column; gap:1px; text-decoration:none; border-radius:10px;
          padding:10px 18px; border:1px solid var(--navy); background:#fff; color:var(--navy);
          min-width:158px; }}
  .big:hover {{ background:#eef2f8; }}
  .big.primary {{ background:var(--navy); border-color:var(--navy); color:#fff; }}
  .big.primary:hover {{ background:#2e5c9a; }}
  .big .t {{ font-weight:700; font-size:14px; }}
  .big .s {{ font-size:11px; opacity:.75; }}
  .cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(250px,1fr)); gap:16px; }}
  .card {{ background:#fff; border:1px solid var(--line); border-radius:14px; padding:20px;
           display:flex; flex-direction:column; gap:6px; }}
  .card .fmt {{ font-size:11px; font-weight:700; letter-spacing:1.4px; text-transform:uppercase; color:#8a94a3; }}
  .card h3 {{ margin:0; font-size:17px; color:var(--navy); }}
  .card p {{ margin:0; font-size:13.5px; color:#4b5563; flex:1; }}
  .btn {{ display:inline-block; margin-top:10px; text-align:center; background:var(--navy); color:#fff;
          text-decoration:none; padding:10px 16px; border-radius:8px; font-weight:600; font-size:14px; }}
  .btn:hover {{ background:#2e5c9a; }}
  .meta {{ font-size:12px; color:#8a94a3; }}
  .panel {{ background:#fff; border:1px solid var(--line); border-radius:14px; padding:10px; overflow:hidden; }}
  .atlas {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(300px,1fr)); gap:14px; margin-top:12px;
            max-height:820px; overflow:auto; padding-right:4px; }}
  .tile {{ margin:0; background:#fff; border:1px solid var(--line); border-radius:12px; padding:10px 10px 6px;
           overflow:hidden; }}
  .tile figcaption {{ display:flex; flex-direction:column; gap:1px; margin:0 2px 6px; }}
  .tile figcaption b {{ font-size:14px; color:var(--navy); }}
  .tile figcaption span {{ font-size:11.5px; color:var(--grey); }}
  .tile .warn {{ color:#a30000; }}
  .tile svg {{ width:100%; height:auto; aspect-ratio:1.25; display:block; border-radius:8px; }}
  .pring {{ fill:rgba(255,140,0,.13); stroke:#ff8c00; stroke-width:.7; vector-effect:non-scaling-stroke; }}
  .pl {{ stroke:#9aa5b1; stroke-width:1; vector-effect:non-scaling-stroke; }}
  .pl-out {{ stroke:#c00000; stroke-width:1.3; stroke-dasharray:2.4 1.8; vector-effect:non-scaling-stroke; }}
  .pcpe {{ fill:#0b5cab; stroke:#fff; stroke-width:.8; vector-effect:non-scaling-stroke; }}
  .pcpe-out {{ fill:#c00000; stroke:#fff; stroke-width:.8; vector-effect:non-scaling-stroke; }}
  .phs {{ fill:#c00000; stroke:#fff; stroke-width:1.1; vector-effect:non-scaling-stroke; }}
  .plbl {{ font-family:-apple-system,"Segoe UI",Roboto,Arial,sans-serif; font-weight:600; fill:#1f3864;
           paint-order:stroke; stroke:#fff; stroke-width:.5px; stroke-linejoin:round; }}
  table {{ width:100%; border-collapse:collapse; font-size:13.5px; background:#fff; }}
  th {{ background:var(--navy); color:#fff; text-align:left; font-size:11.5px; letter-spacing:.6px;
        text-transform:uppercase; padding:9px 10px; position:sticky; top:0; }}
  td {{ padding:8px 10px; border-top:1px solid var(--line); vertical-align:top; }}
  tr:nth-child(even) td {{ background:#fbfcfe; }}
  td.n {{ text-align:right; font-variant-numeric:tabular-nums; }}
  code {{ font-size:12px; color:#3b475a; }}
  .scroll {{ max-height:430px; overflow:auto; border:1px solid var(--line); border-radius:12px; }}
  .legend {{ display:flex; flex-wrap:wrap; gap:18px; font-size:12.5px; color:#4b5563; margin:12px 2px 0; }}
  .legend i {{ display:inline-block; width:11px; height:11px; border-radius:50%; margin-right:6px;
               vertical-align:-1px; border:1px solid #fff; box-shadow:0 0 0 1px #ccd; }}
  .legend .sq {{ border-radius:2px; }}
  .tag {{ font-size:10.5px; font-weight:700; padding:2px 7px; border-radius:999px; text-transform:uppercase;
          letter-spacing:.5px; }}
  .tag.high {{ background:#fde2e2; color:#a30000; }}
  .tag.medium {{ background:#fff1cc; color:#7a5b00; }}
  .tag.low {{ background:#e9edf3; color:#4b5563; }}
  .tag.info {{ background:#e7eefb; color:#2e5c9a; }}
  ul.notes {{ font-size:14px; color:#3b475a; padding-left:20px; }}
  ul.notes li {{ margin:7px 0; }}
  footer {{ margin-top:54px; padding-top:18px; border-top:1px solid var(--line); font-size:12.5px; color:#8a94a3; }}
</style>
</head>
<body>
<header>
  <div class="wrap">
    <div class="eyebrow">North West province &middot; current technology site database</div>
    <h1>Site network map and improved spreadsheet</h1>
    <p>All {len(rows)} sites from the &ldquo;CURRENT TECHNOLOGY&rdquo; sheet, with every high site drawn at its
       {RADIUS_KM:.0f} km coverage radius and every connected site linked to the high site that serves it.</p>
  </div>
</header>

<div class="wrap">
  <div class="kpis">
    <div class="kpi"><b>{len(rows)}</b><span>sites in total</span></div>
    <div class="kpi"><b>{len(hs)}</b><span>high sites</span></div>
    <div class="kpi"><b>{len(cpe)}</b><span>connected sites</span></div>
    <div class="kpi"><b>{RADIUS_KM:.0f} km</b><span>coverage radius</span></div>
    <div class="kpi"><b>{len(inside)}</b><span>inside the radius</span></div>
    <div class="kpi"><b>{len(outside)}</b><span>outside the radius</span></div>
    <div class="kpi"><b>{len(high)}</b><span>items to verify</span></div>
  </div>

  <div class="band">
    <div class="band-txt">
      <b>Download the results</b>
      <span>Everything below is built from the same {len(rows)} sites &mdash; one map, one workbook, no differences between them.</span>
    </div>
    <div class="band-btns">
      <a class="big primary" href="NW_Site_Network_Pack.zip" download>
        <span class="t">Download both</span>
        <span class="s">ZIP &middot; {size('NW_Site_Network_Pack.zip')}</span>
      </a>
      <a class="big" href="NW_Network_Sites.kmz" download>
        <span class="t">Google Earth map</span>
        <span class="s">KMZ &middot; {size('NW_Network_Sites.kmz')}</span>
      </a>
      <a class="big" href="NW_Network_Site_Database.xlsx" download>
        <span class="t">Excel report</span>
        <span class="s">XLSX &middot; {size('NW_Network_Site_Database.xlsx')}</span>
      </a>
    </div>
  </div>

  <h2>Downloads <span>&mdash; both files cover exactly the same {len(rows)} sites</span></h2>
  <div class="cards">
    <div class="card">
      <div class="fmt">KMZ &middot; {size('NW_Network_Sites.kmz')}</div>
      <h3>Google Earth map</h3>
      <p>Opens in Google Earth. 20 high sites with a 5 km radius ring, 86 connected sites,
         a link line from every connected site to its high site, and layers you can switch off.</p>
      <a class="btn" href="NW_Network_Sites.kmz" download>Download KMZ</a>
      <div class="meta">Google Earth Pro / web / mobile</div>
    </div>
    <div class="card">
      <div class="fmt">XLSX &middot; {size('NW_Network_Site_Database.xlsx')}</div>
      <h3>Improved spreadsheet</h3>
      <p>7 sheets: overview, all sites with both distances and the 5 km check, high sites,
         towns, remote links, data quality and the untouched source rows for audit.</p>
      <a class="btn" href="NW_Network_Site_Database.xlsx" download>Download Excel</a>
      <div class="meta">Excel / LibreOffice / Google Sheets</div>
    </div>
    <div class="card">
      <div class="fmt">KML + CSV &middot; {size('NW_Network_Sites.kml')}</div>
      <h3>Plain KML or CSV</h3>
      <p>Use the KML for Google My Maps or any GIS tool. The CSV is the cleaned data set
         for import into your own systems.</p>
      <a class="btn" href="NW_Network_Sites.kml" download>Download KML</a>
      <div class="meta" style="margin-top:4px">
        <a href="nw_sites_clean.csv" download style="color:#2e5c9a">or the cleaned CSV</a>
      </div>
    </div>
  </div>

  <h2>Coverage per town <span>&mdash; same geometry as the KMZ, one panel per town</span></h2>
  <div class="legend">
    <span><i class="sq" style="background:#c00000"></i>High site</span>
    <span><i style="background:rgba(255,140,0,.18);border:1px solid #ff8c00"></i>{RADIUS_KM:.0f} km coverage radius</span>
    <span><i style="background:#0b5cab"></i>Connected site inside the radius</span>
    <span><i style="background:#c00000"></i>Connected site outside the radius</span>
    <span><i style="background:#9aa5b1"></i>Link inside the radius</span>
    <span><i style="background:#c00000;border-radius:0"></i>Link outside the radius</span>
  </div>
  <div class="atlas">
{panels}
  </div>

  <h2>High sites and their coverage</h2>
  <div class="scroll">
    <table>
      <thead><tr><th>Town</th><th>High site</th><th>Bearer</th><th>Connected</th>
      <th>In 5 km</th><th>Out</th><th>Farthest (km)</th><th>Coordinates</th></tr></thead>
      <tbody>{hs_rows()}</tbody>
    </table>
  </div>

  <h2>Connected sites outside the {RADIUS_KM:.0f} km radius <span>&mdash; {len(outside)} of {len(cpe)}</span></h2>
  <table>
    <thead><tr><th>Site</th><th>Town</th><th>High site</th><th>Distance (km)</th><th>Beyond radius</th></tr></thead>
    <tbody>{out_rows()}</tbody>
  </table>

  <h2>What had to be fixed</h2>
  <table>
    <thead><tr><th>Severity</th><th>Site</th><th>Field</th><th>Finding</th></tr></thead>
    <tbody>{issue_rows()}</tbody>
  </table>
  <p class="meta">All {len(issues)} findings are listed on the <b>Data Quality</b> sheet of the workbook.</p>

  <h2>Method</h2>
  <ul class="notes">
    <li><b>Coverage:</b> every high site carries a {RADIUS_KM:.0f} km radius ring; a connected site is reported as
        inside or outside that ring using its great-circle distance from the high site it serves.</li>
    <li><b>Connectivity:</b> each connected site is joined to its high site with a link line, in the KMZ and on the map above.</li>
    <li><b>Distances were recalculated</b> from the coordinates and compared with the sheet. Nine sites come out
        more than 1 km different from the figure in the sheet; the differences are listed on the Data Quality sheet.</li>
    <li><b>Two high sites had no coordinates</b> in the sheet (Brits Library, Lethabong Maboloka Library). They are
        positioned by trilateration from their connected sites and the sheet's own distances, and flagged
        &ldquo;estimated&rdquo; in every file.</li>
    <li><b>Three Delareyville sites</b> have a distance but no high site anywhere in the sheet, so they are left
        unassigned rather than guessed.</li>
    <li>Nothing in the source data was deleted: the original rows 40&ndash;145 are kept verbatim on the last sheet.</li>
  </ul>

  <footer>Built from the screenshot of the &ldquo;CURRENT TECHNOLOGY&rdquo; sheet, rows 40&ndash;145
  ({len(rows)} records). Geometry: WGS84, great-circle distances, {RADIUS_KM:.0f} km rings.
  Static page &mdash; no tracking, no external requests.</footer>
</div>
</body>
</html>
"""
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)
    return path


# --------------------------------------------------------------------------
# 7. main
# --------------------------------------------------------------------------


def main():
    rows, by_row, issues = build()
    kml_path = os.path.join(HERE, "NW_Network_Sites.kml")
    kmz_path = os.path.join(HERE, "NW_Network_Sites.kmz")
    xlsx_path = os.path.join(HERE, "NW_Network_Site_Database.xlsx")
    csv_path = os.path.join(HERE, "nw_sites_clean.csv")
    html_path = os.path.join(HERE, "index.html")

    zip_path = os.path.join(HERE, "NW_Site_Network_Pack.zip")

    write_xlsx(rows, issues, xlsx_path)          # first: index.html reads file sizes
    kml_text = write_kml(rows, kml_path)
    write_kmz(kml_text, kmz_path)
    write_csv(rows, csv_path)
    write_zip(
        zip_path,
        [
            "NW_Network_Sites.kmz",
            "NW_Network_Site_Database.xlsx",
            "NW_Network_Sites.kml",
            "nw_sites_clean.csv",
            "README.md",
        ],
    )
    write_index(rows, issues, html_path)

    for p in (kml_path, kmz_path, xlsx_path, csv_path, zip_path, html_path):
        print(f"{os.path.getsize(p)/1024:8.1f} KB  {os.path.relpath(p, os.path.dirname(HERE))}")


if __name__ == "__main__":
    main()
