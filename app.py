"""
GEOLOCATOR  v7
──────────────
Changes from v6:
  • Renamed to GeoLocator
  • Output preserves ALL original columns — geocoding columns added on the right
  • Review panel shows original row data + free-text Notes field
  • Skipped batches are excluded entirely (not in review, not in download)
  • Lighter, readable Excel colour scheme
"""

import io
import time
import urllib.parse

import streamlit as st
import pandas as pd
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="GeoLocator",
    page_icon="📍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.block-container { padding-top: 1rem !important; max-width: 1200px; }
.review-box { background:#f8fafc; border:1px solid #e2e8f0;
              border-radius:8px; padding:1rem 1.2rem; margin-bottom:0.8rem; }
.orig-data { background:#f0f9ff; border:1px solid #bae6fd;
             border-radius:6px; padding:0.7rem 1rem; margin-bottom:0.6rem;
             font-size:0.82rem; }
#MainMenu{visibility:hidden;} footer{visibility:hidden;}
</style>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ═════════════════════════════════════════════════════════════════════════════
def ss(k, v):
    if k not in st.session_state:
        st.session_state[k] = v

ss("phase",        1)
ss("batch_idx",    0)
ss("all_results",  [])   # only includes processed rows (skipped excluded)
ss("review_queue", [])
ss("reviewed",     [])
ss("review_idx",   0)
ss("api_key",      "")
ss("col_config", {
    "sheet":     "RMI raw",
    "facility":  "Facility Name",
    "group":     "Smelter Look-up (*)",
    "city":      "City",
    "country":   "Country",
    "use_group": True,
})

BATCH_SIZE = 200


# ═════════════════════════════════════════════════════════════════════════════
#  GOOGLE MAPS GEOCODING
# ═════════════════════════════════════════════════════════════════════════════
LOCATION_TYPE_CONF = {
    "ROOFTOP":            ("High",   98),
    "RANGE_INTERPOLATED": ("High",   85),
    "GEOMETRIC_CENTER":   ("Medium", 65),
    "APPROXIMATE":        ("Low",    35),
}

FACILITY_RESULT_TYPES = {
    "establishment", "point_of_interest", "premise", "subpremise",
    "natural_feature", "airport", "park", "place_of_worship",
    "street_address", "route",
}

CITY_ONLY_TYPES = {
    "locality", "administrative_area_level_1", "administrative_area_level_2",
    "administrative_area_level_3", "administrative_area_level_4",
    "country", "postal_code", "colloquial_area", "neighborhood",
    "sublocality", "sublocality_level_1",
}


def _google_geocode(query: str, api_key: str) -> dict:
    try:
        r = requests.get(
            "https://maps.googleapis.com/maps/api/geocode/json",
            params={"address": query, "key": api_key},
            timeout=10,
        )
        r.raise_for_status()
        return r.json()
    except requests.exceptions.RequestException as e:
        return {"status": "CONNECTION_ERROR", "error": str(e)}


def _is_facility_result(result: dict) -> bool:
    types    = set(result.get("types", []))
    loc_type = result.get("geometry", {}).get("location_type", "")
    if types and types.issubset(CITY_ONLY_TYPES):
        return False
    if loc_type in ("ROOFTOP", "RANGE_INTERPOLATED"):
        return True
    if types & FACILITY_RESULT_TYPES:
        return True
    if loc_type == "GEOMETRIC_CENTER" and not types.issubset(CITY_ONLY_TYPES):
        return True
    return False


def _country_ok(result: dict, expected_country: str) -> bool:
    if not expected_country or expected_country.strip().lower() in ("", "nan"):
        return True
    addr = result.get("formatted_address", "").lower()
    ec   = expected_country.lower().strip()
    aliases = {
        "usa": ["united states"], "us": ["united states"],
        "uk":  ["united kingdom"], "uae": ["united arab emirates"],
        "south korea": ["korea"], "korea": ["south korea"],
        "czechia": ["czech republic"], "czech republic": ["czechia"],
    }
    return any(v in addr for v in [ec] + aliases.get(ec, []))


def _parse_result(result: dict) -> dict:
    loc      = result["geometry"]["location"]
    loc_type = result["geometry"].get("location_type", "APPROXIMATE")
    conf, score = LOCATION_TYPE_CONF.get(loc_type, ("Low", 30))
    return {
        "status":           "OK",
        "latitude":         round(loc["lat"], 6),
        "longitude":        round(loc["lng"], 6),
        "matched_address":  result.get("formatted_address", ""),
        "confidence":       conf,
        "confidence_score": score,
        "location_type":    loc_type,
        "result_types":     ", ".join(result.get("types", [])),
    }


def geocode_name_google(name: str, city: str, country: str, api_key: str):
    if not name or name.strip().lower() in ("", "nan"):
        return None
    parts = [name.strip()]
    if city    and city.strip().lower()    not in ("", "nan"): parts.append(city.strip())
    if country and country.strip().lower() not in ("", "nan"): parts.append(country.strip())
    query = ", ".join(parts)
    data  = _google_geocode(query, api_key)
    if data.get("status") == "REQUEST_DENIED":
        st.error(f"❌ API key rejected: {data.get('error_message','')}")
        st.stop()
    if data.get("status") != "OK" or not data.get("results"):
        return None
    for result in data["results"]:
        if not _country_ok(result, country):
            continue
        if _is_facility_result(result):
            parsed = _parse_result(result)
            parsed["query_used"]  = query
            parsed["match_level"] = "name_match"
            return parsed
    return None


def geocode_facility_google(facility, group, city, country, use_group, api_key):
    result = geocode_name_google(facility, city, country, api_key)
    if result:
        result["matched_by"] = "Facility Name"
        return result
    if use_group and group and group.strip().lower() not in ("", "nan", facility.lower()):
        result = geocode_name_google(group, city, country, api_key)
        if result:
            result["matched_by"] = "Group Name"
            return result
    return {
        "status": "ZERO_RESULTS", "error": "No facility name match found",
        "latitude": None, "longitude": None, "matched_address": "",
        "confidence": "None", "confidence_score": 0,
        "matched_by": "—", "query_used": f"{facility}/{group}",
        "match_level": "none",
    }


def geocode_address_google(address: str, api_key: str) -> dict:
    data = _google_geocode(address, api_key)
    if data.get("status") != "OK" or not data.get("results"):
        return {"status": "ZERO_RESULTS"}
    result = data["results"][0]
    parsed = _parse_result(result)
    parsed["matched_address"] = result.get("formatted_address", "")
    return parsed


def validate_api_key(api_key: str) -> tuple:
    data = _google_geocode("London, United Kingdom", api_key)
    if data.get("status") == "OK":
        return True, ""
    if data.get("status") == "REQUEST_DENIED":
        return False, data.get("error_message", "Key rejected")
    return False, f"Status: {data.get('status')}"


# ═════════════════════════════════════════════════════════════════════════════
#  REVIEW QUEUE BUILDER
# ═════════════════════════════════════════════════════════════════════════════
def _build_review_queue():
    auto  = []
    queue = []
    for r in st.session_state.all_results:
        if r.get("confidence") == "High" and r.get("status") == "OK":
            r["final_status"] = "Auto-High"
            r["verified_by"]  = f"Google Maps High ({r.get('matched_by','name')})"
            auto.append(r)
        else:
            r.setdefault("final_status", "Pending Review")
            r["review_reason"] = (f"Confidence: {r.get('confidence','?')} "
                                  f"| Status: {r.get('status','?')}")
            queue.append(r)
    st.session_state.reviewed     = auto
    st.session_state.review_queue = queue


# ═════════════════════════════════════════════════════════════════════════════
#  SATELLITE MAP
# ═════════════════════════════════════════════════════════════════════════════
def sat_map(lat, lon, label, zoom=16):
    fac = str(label).replace("'", "\\'").replace('"', '\\"')
    if not lat and not lon:
        lat, lon, zoom = 20.0, 0.0, 2
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8"/>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box;}}
#map{{width:100%;height:390px;border-radius:6px;}}
#co{{font-family:monospace;font-size:0.85rem;padding:8px 10px;
     background:#e0f2fe;border:2px solid #0284c7;border-radius:6px;
     margin-top:6px;color:#0c4a6e;font-weight:600;}}
#hint{{font-size:0.72rem;color:#64748b;margin-top:4px;}}
</style></head><body>
<div id="map"></div>
<div id="co">📍 Lat: <span id="dlat">{lat}</span> &nbsp;|&nbsp; Lon: <span id="dlon">{lon}</span></div>
<div id="hint">Click or drag pin to reposition</div>
<script>
var map=L.map('map',{{zoomControl:true}}).setView([{lat},{lon}],{zoom});
L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{{z}}/{{y}}/{{x}}',
  {{attribution:'Esri',maxZoom:19}}).addTo(map);
L.tileLayer('https://services.arcgisonline.com/arcgis/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{{z}}/{{y}}/{{x}}',
  {{attribution:'Esri',maxZoom:19,opacity:0.85}}).addTo(map);
var mk=L.marker([{lat},{lon}],{{draggable:true}}).addTo(map);
mk.bindPopup('<b>{fac}</b>').openPopup();
function upd(a,o){{
  a=Math.round(a*1e6)/1e6; o=Math.round(o*1e6)/1e6;
  document.getElementById('dlat').textContent=a;
  document.getElementById('dlon').textContent=o;
}}
mk.on('dragend',function(e){{var p=e.target.getLatLng();upd(p.lat,p.lng);}});
map.on('click',function(e){{mk.setLatLng(e.latlng);upd(e.latlng.lat,e.latlng.lng);}});
window.addEventListener('message',function(e){{
  if(e.data&&e.data.type==='set_pin'){{
    var ll=L.latLng(e.data.lat,e.data.lon);
    mk.setLatLng(ll);map.setView(ll,17);upd(e.data.lat,e.data.lon);
  }}
}});
</script></body></html>"""


# ═════════════════════════════════════════════════════════════════════════════
#  EXCEL OUTPUT  —  preserves all original columns, lighter colours
# ═════════════════════════════════════════════════════════════════════════════
_T  = Side(style="thin", color="cbd5e1")
_B  = Border(left=_T, right=_T, top=_T, bottom=_T)
_HF = PatternFill("solid", fgColor="1e3a5f")   # header: dark blue

# Light, readable status fills
_GF = PatternFill("solid", fgColor="dcfce7")   # light green  — verified
_YF = PatternFill("solid", fgColor="fef9c3")   # light yellow — medium
_OF = PatternFill("solid", fgColor="ffedd5")   # light orange — low/pending
_RF = PatternFill("solid", fgColor="fee2e2")   # light red    — no match/rejected
_NF = PatternFill("solid", fgColor="f1f5f9")   # light grey   — original cols

STATUS_FILL = {
    "Auto-High":    _GF, "Verified":  _GF, "Corrected": _GF,
    "Auto-Medium":  _YF,
    "Rejected":     _RF, "No Match":  _RF, "Needs Review": _RF,
}

# New geocoding columns to append to the right of original data
GEO_COLS = [
    "Latitude", "Longitude", "Matched Address",
    "Geo Status", "Matched By", "Confidence", "Score /100",
    "Location Type", "Verified By", "Notes",
]


def _hdr_cell(c, text, fill=None):
    c.value = text
    c.font  = Font(name="Calibri", bold=True,
                   color="FFFFFF" if fill == _HF else "1e3a5f", size=10)
    c.fill  = fill or _HF
    c.border = _B
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _aw(ws):
    for col in ws.columns:
        l = get_column_letter(col[0].column)
        w = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[l].width = min(w + 4, 55)


def build_excel(rows: list, orig_columns: list) -> bytes:
    """
    Build output workbook.
    Each row contains the original data dict (_orig) plus geocoding fields.
    orig_columns = ordered list of original column names from the source sheet.
    """
    wb = openpyxl.Workbook()

    # ── Sheet 1: All geocoded results ─────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Geocoded Results"
    ws1.freeze_panes = "A2"

    all_cols = orig_columns + GEO_COLS
    # Write header
    for ci, col in enumerate(all_cols, 1):
        c = ws1.cell(1, ci)
        # Original cols get a lighter header tint
        if col in orig_columns:
            c.value = col
            c.font  = Font(name="Calibri", bold=True, color="1e3a5f", size=10)
            c.fill  = PatternFill("solid", fgColor="dbeafe")
            c.border = _B
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        else:
            _hdr_cell(c, col)
    ws1.row_dimensions[1].height = 22

    for r in rows:
        status = r.get("final_status", "Needs Review")
        fill   = STATUS_FILL.get(status, _RF)
        orig   = r.get("_orig", {})

        row_vals = []
        # Original columns first
        for col in orig_columns:
            row_vals.append(orig.get(col, ""))
        # Geocoding columns
        row_vals += [
            r.get("latitude"),    r.get("longitude"),
            r.get("matched_address"),
            status,
            r.get("matched_by", "—"),
            r.get("confidence"),  r.get("confidence_score"),
            r.get("location_type", ""),
            r.get("verified_by", ""),
            r.get("notes", ""),
        ]

        ws1.append(row_vals)
        ri = ws1.max_row

        for ci, col in enumerate(all_cols, 1):
            c = ws1.cell(ri, ci)
            c.font = Font(name="Calibri", size=10, color="111827")
            # Original cols: no status fill, just light stripe
            c.fill = _NF if col in orig_columns else fill
            c.border = _B
            c.alignment = Alignment(vertical="center")

        # Lat/Lon columns
        lat_ci = len(orig_columns) + 1
        lon_ci = len(orig_columns) + 2
        ws1.cell(ri, lat_ci).number_format = "0.000000"
        ws1.cell(ri, lon_ci).number_format = "0.000000"

    _aw(ws1)

    # ── Sheet 2: Needs Manual Review ──────────────────────────────────────────
    ws2 = wb.create_sheet("Needs Manual Review")
    ws2.freeze_panes = "A2"

    review_cols = orig_columns + [
        "Issue", "Latitude (fill manually)", "Longitude (fill manually)", "Notes"
    ]
    for ci, col in enumerate(review_cols, 1):
        c = ws2.cell(1, ci)
        if col in orig_columns:
            c.value = col
            c.font  = Font(name="Calibri", bold=True, color="1e3a5f", size=10)
            c.fill  = PatternFill("solid", fgColor="dbeafe")
            c.border = _B
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        else:
            _hdr_cell(c, col)
    ws2.row_dimensions[1].height = 22

    needs = [r for r in rows if r.get("final_status") in
             ("Needs Review", "Rejected", "No Match")]
    if needs:
        for r in needs:
            orig = r.get("_orig", {})
            row_vals = [orig.get(col, "") for col in orig_columns]
            row_vals += [
                r.get("review_reason", r.get("error", "No match")),
                "", "",
                r.get("notes", ""),
            ]
            ws2.append(row_vals)
            ri = ws2.max_row
            for ci in range(1, len(review_cols)+1):
                c = ws2.cell(ri, ci)
                c.font = Font(name="Calibri", size=10, color="111827")
                c.fill = _NF if review_cols[ci-1] in orig_columns else _RF
                c.border = _B
                c.alignment = Alignment(vertical="center")
    else:
        ws2.append(["✅  All rows have coordinates — nothing needs manual review!"])
        ws2.cell(ws2.max_row, 1).font = Font(name="Calibri", bold=True,
                                              color="166534", size=11)
    _aw(ws2)

    # ── Sheet 3: Summary ──────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Summary")
    total   = len(rows)
    vfd     = sum(1 for r in rows if r.get("final_status") in
                  ("Auto-High","Verified","Corrected"))
    manual  = sum(1 for r in rows if r.get("final_status") in
                  ("Needs Review","Rejected","No Match"))
    rooftop = sum(1 for r in rows if r.get("location_type") == "ROOFTOP")
    by_grp  = sum(1 for r in rows if r.get("matched_by") == "Group Name")

    for a, b in [
        ("GEOLOCATOR — RUN SUMMARY",    ""),
        ("",                             ""),
        ("Total rows processed",         total),
        ("Coordinates assigned ✅",      vfd),
        ("  ROOFTOP precision",          rooftop),
        ("  Matched via Group Name",     by_grp),
        ("Need manual work 📋",          manual),
        ("",                             ""),
        ("Coverage",   f"{vfd/total*100:.1f}%" if total else "N/A"),
        ("",                             ""),
        ("Geocoding source",             "Google Maps Geocoding API"),
        ("Name-first rule",              "City/country fallback disabled"),
        ("Skipped batches",              "Excluded from this file"),
    ]:
        ws3.append([a, b])
        ws3.cell(ws3.max_row, 1).font = Font(name="Calibri", bold=True, size=10)
        ws3.cell(ws3.max_row, 2).font = Font(name="Calibri", size=10)
    ws3.column_dimensions["A"].width = 35
    ws3.column_dimensions["B"].width = 32

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ═════════════════════════════════════════════════════════════════════════════
#  UI
# ═════════════════════════════════════════════════════════════════════════════
st.title("📍 GeoLocator")
st.caption("Google Maps · Batch processing · Name-first · Satellite review")
st.divider()


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 1 — CONFIG + UPLOAD
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.phase == 1:

    st.subheader("Step 1 — Google Maps API Key")
    st.info("Your key is only held in memory while the app runs — never saved to disk.")

    api_input = st.text_input(
        "Paste your Google Maps API key",
        type="password",
        value=st.session_state.api_key,
        placeholder="AIzaSy...",
    )
    if api_input and api_input != st.session_state.api_key:
        with st.spinner("Validating…"):
            ok, err = validate_api_key(api_input)
        if ok:
            st.session_state.api_key = api_input
            st.success("✅ API key valid!")
        else:
            st.error(f"❌ {err} — make sure the Geocoding API is enabled.")

    if not st.session_state.api_key:
        st.markdown("""
        **How to get a free key:**
        1. Go to **console.cloud.google.com**
        2. Create a project → search **Geocoding API** → **Enable**
        3. **Credentials** → **+ Create Credentials** → **API Key**
        4. Add a billing account (won't be charged at this scale)
        """)
        st.stop()

    st.divider()
    st.subheader("Step 2 — Column settings")
    cfg = st.session_state.col_config

    col_l, col_r = st.columns(2, gap="large")
    with col_l:
        cfg["sheet"]    = st.text_input("Sheet name",           value=cfg["sheet"])
        cfg["facility"] = st.text_input("Facility name column", value=cfg["facility"])
        cfg["city"]     = st.text_input("City column",          value=cfg["city"])
        cfg["country"]  = st.text_input("Country column",       value=cfg["country"])
    with col_r:
        cfg["use_group"] = st.toggle("Use secondary / group name column",
                                      value=cfg["use_group"])
        if cfg["use_group"]:
            cfg["group"] = st.text_input("Secondary name column", value=cfg["group"])
        st.info(
            "**Name-first:** coordinates only assigned when a facility name "
            "is found on the map — no city/country fallback.\n\n"
            "**Skipped batches** are excluded from review and the final download."
        )
    st.session_state.col_config = cfg

    st.divider()
    st.subheader("Step 3 — Upload your Excel file")
    uploaded = st.file_uploader("Choose .xlsx file", type=["xlsx", "xlsm"])

    if uploaded:
        try:
            all_sheets = pd.read_excel(uploaded, sheet_name=None)
            sname = cfg["sheet"]
            if sname not in all_sheets:
                st.error(f"Sheet '{sname}' not found. Available: {list(all_sheets.keys())}")
            else:
                df = all_sheets[sname]
                required = [cfg["facility"], cfg["city"], cfg["country"]]
                missing  = [c for c in required if c not in df.columns]
                group_missing = (cfg["use_group"] and cfg.get("group") and
                                 cfg["group"] not in df.columns)
                if missing:
                    st.error(f"Columns not found: {missing}\nIn sheet: {list(df.columns)}")
                else:
                    if group_missing:
                        st.warning(f"Group column '{cfg['group']}' not found — will be skipped.")
                    n         = len(df)
                    n_batches = -(-n // BATCH_SIZE)
                    est_min   = round(n * 0.1 / 60, 1)
                    st.success(f"✅ **{n} rows** · {n_batches} batch(es) of {BATCH_SIZE}")
                    st.caption(f"Estimated time if all batches run: ~{est_min} min")
                    with st.expander("Preview first 5 rows"):
                        st.dataframe(df.head(5), use_container_width=True)

                    if st.button("▶  Start", type="primary"):
                        st.session_state.upload_df      = df
                        st.session_state.upload_name    = uploaded.name
                        st.session_state.orig_columns   = list(df.columns)
                        st.session_state.group_exists   = not group_missing
                        st.session_state.batch_idx      = 0
                        st.session_state.all_results    = []
                        st.session_state.phase          = 2
                        st.rerun()
            uploaded.seek(0)
        except Exception as e:
            st.error(f"Could not read file: {e}")


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 2 — BATCH GEOCODING
# ─────────────────────────────────────────────────────────────────────────────
elif st.session_state.phase == 2:

    df          = st.session_state.upload_df
    cfg         = st.session_state.col_config
    api_key     = st.session_state.api_key
    orig_cols   = st.session_state.orig_columns
    use_group   = cfg["use_group"] and st.session_state.get("group_exists", False)
    group_col   = cfg.get("group", "") if use_group else ""
    total_rows  = len(df)
    n_batches   = -(-total_rows // BATCH_SIZE)
    batch_idx   = st.session_state.batch_idx
    done_rows   = len(st.session_state.all_results)

    st.subheader(f"Geocoding — Batch {batch_idx+1} of {n_batches}")

    overall_pct = int(done_rows / total_rows * 100) if total_rows else 0
    st.progress(overall_pct,
                text=f"{done_rows} / {total_rows} rows processed ({overall_pct}%)")

    if batch_idx >= n_batches:
        st.success("✅ All batches done!")
        _build_review_queue()
        if st.button("▶  Go to Satellite Review", type="primary"):
            st.session_state.phase      = 3
            st.session_state.review_idx = 0
            st.rerun()
        st.stop()

    # Running stats
    if st.session_state.all_results:
        hi      = sum(1 for r in st.session_state.all_results if r.get("confidence") == "High")
        no_m    = sum(1 for r in st.session_state.all_results if r.get("status") == "ZERO_RESULTS")
        rooftop = sum(1 for r in st.session_state.all_results if r.get("location_type") == "ROOFTOP")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Processed",       done_rows)
        c2.metric("High confidence", hi)
        c3.metric("Rooftop",         rooftop)
        c4.metric("No match",        no_m)

    start_i  = batch_idx * BATCH_SIZE
    end_i    = min(start_i + BATCH_SIZE, total_rows)
    batch_df = df.iloc[start_i:end_i]

    st.info(
        f"**Batch {batch_idx+1}:** rows {start_i+1}–{end_i} "
        f"(~{round(len(batch_df)*0.1/60,1)} min)  ·  "
        f"Skipping a batch **permanently excludes** those rows from review and download."
    )

    col_run, col_skip, col_dl = st.columns([2, 1, 1])
    with col_run:
        run_batch = st.button(f"▶  Run Batch {batch_idx+1}",
                              type="primary", use_container_width=True)
    with col_skip:
        skip_batch = st.button("⏭  Skip & exclude", use_container_width=True,
                               help="These rows will NOT appear in review or the final download.")
    with col_dl:
        if st.session_state.all_results:
            dl_bytes = build_excel(
                st.session_state.all_results,
                st.session_state.orig_columns,
            )
            st.download_button(
                "⬇️ Save progress", data=dl_bytes,
                file_name="geolocator_progress.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    if skip_batch:
        # Simply advance batch index — do NOT add rows to all_results
        st.session_state.batch_idx += 1
        st.rerun()

    if run_batch:
        pbar       = st.progress(0)
        status_txt = st.empty()

        seen = {
            f"{r.get('_orig',{}).get(cfg['facility'],'').lower()}|"
            f"{r.get('_orig',{}).get(cfg['country'],'').lower()}": r
            for r in st.session_state.all_results
            if r.get("status") == "OK"
        }

        batch_results = []
        for j, (_, row) in enumerate(batch_df.iterrows()):
            # Store the full original row
            orig_data = {col: row.get(col, "") for col in orig_cols}

            facility = str(row.get(cfg["facility"], "") or "").strip()
            group    = str(row.get(group_col, "")       or "").strip() if group_col else ""
            city     = str(row.get(cfg["city"],    "")  or "").strip()
            country  = str(row.get(cfg["country"], "")  or "").strip()

            if not facility and not group and not city and not country:
                pbar.progress(int((j+1) / len(batch_df) * 100))
                continue

            cache_key = f"{facility.lower()}|{country.lower()}"
            status_txt.markdown(f"🔍 **{j+1}/{len(batch_df)}** — `{facility or group or city}`")

            if cache_key in seen:
                cached = seen[cache_key]
                result = {
                    **cached,
                    "facility": facility, "group": group,
                    "city": city, "country": country,
                    "_orig": orig_data, "duplicate": True,
                }
            else:
                geo = geocode_facility_google(
                    facility, group, city, country, use_group, api_key
                )
                result = {
                    "facility": facility, "group": group,
                    "city": city, "country": country,
                    "_orig": orig_data,
                    **geo,
                }
                if geo["status"] == "OK":
                    seen[cache_key] = geo

            batch_results.append(result)
            pbar.progress(int((j+1) / len(batch_df) * 100))

        pbar.progress(100)
        status_txt.markdown(f"✅ Batch {batch_idx+1} done!")

        st.session_state.all_results.extend(batch_results)
        st.session_state.batch_idx += 1

        matched  = sum(1 for r in batch_results if r.get("status") == "OK")
        rooftops = sum(1 for r in batch_results if r.get("location_type") == "ROOFTOP")
        next_b   = batch_idx + 2
        cont_msg = (f"Click **Run Batch {next_b}** to continue."
                    if end_i < total_rows else "All batches done!")
        st.success(
            f"Batch {batch_idx+1}: **{matched}/{len(batch_results)}** matched "
            f"({rooftops} rooftop). {cont_msg}"
        )
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 3 — SATELLITE REVIEW
# ─────────────────────────────────────────────────────────────────────────────
elif st.session_state.phase == 3:

    if not st.session_state.review_queue and not st.session_state.reviewed:
        _build_review_queue()

    queue   = st.session_state.review_queue
    idx     = st.session_state.review_idx
    total_q = len(queue)

    if idx >= total_q:
        st.session_state.phase = 4
        st.rerun()

    row = queue[idx]
    pct = int(idx / total_q * 100) if total_q else 100

    st.subheader(f"Satellite Review — {idx} of {total_q} done")
    st.progress(pct)

    no_match = row.get("status") == "ZERO_RESULTS"
    map_key  = f"mk_{idx}"
    if map_key not in st.session_state:
        st.session_state[map_key] = {
            "lat": row.get("latitude")  or 0.0,
            "lon": row.get("longitude") or 0.0,
        }
    cur_lat = st.session_state[map_key]["lat"]
    cur_lon = st.session_state[map_key]["lon"]

    col_map, col_ctrl = st.columns([1.6, 1], gap="large")

    with col_map:
        # Facility info
        group_line = (f"<br><small>Group: {row.get('group','')}</small>"
                      if row.get("group") else "")
        match_info = (
            "❌ No name match — use the tools on the right."
            if no_match else
            f"📍 Matched: <b>{row.get('matched_address','—')[:90]}</b><br>"
            f"<small>Confidence: {row.get('confidence','?')} · "
            f"Type: {row.get('location_type','?')} · "
            f"By: {row.get('matched_by','?')}</small>"
        )
        st.markdown(f"""
        <div class="review-box">
          <b style="font-size:1.05rem">{row.get('facility','—')}</b>
          {group_line}
          <br><small style="color:#64748b">
            {row.get('city','')}{', ' if row.get('city') and row.get('country') else ''}
            {row.get('country','')}
          </small>
          <hr style="margin:0.5rem 0;border-color:#e2e8f0">
          {match_info}
        </div>""", unsafe_allow_html=True)

        # Show original data fields
        orig = row.get("_orig", {})
        if orig:
            cfg = st.session_state.col_config
            skip_cols = {cfg.get("facility",""), cfg.get("group",""),
                         cfg.get("city",""), cfg.get("country","")}
            extra = {k: v for k, v in orig.items()
                     if k not in skip_cols and v and str(v).strip() not in ("", "nan")}
            if extra:
                fields_html = "  ·  ".join(
                    f"<b>{k}:</b> {v}" for k, v in list(extra.items())[:8]
                )
                st.markdown(
                    f'<div class="orig-data">📋 {fields_html}</div>',
                    unsafe_allow_html=True,
                )

        zoom = 16 if (cur_lat or cur_lon) else 2
        st.components.v1.html(
            sat_map(cur_lat, cur_lon, row.get("facility", ""), zoom=zoom),
            height=440, scrolling=False,
        )
        st.caption(
            "1️⃣ Move pin  →  2️⃣ Copy Lat & Lon from blue box  →  "
            "3️⃣ Paste into fields on right  →  4️⃣ Click ✔ Use these  →  5️⃣ Verify"
        )

    with col_ctrl:
        st.markdown("**Find the correct location**")

        # Address search
        with st.expander("🏠 Search by address", expanded=no_match):
            addr = st.text_input(
                "Address or place name",
                placeholder="e.g. Aurubis AG Hamburg Germany",
                key=f"addr_{idx}",
            )
            if st.button("Search", key=f"asrch_{idx}", use_container_width=True):
                if addr.strip():
                    with st.spinner("Searching…"):
                        geo = geocode_address_google(addr.strip(),
                                                     st.session_state.api_key)
                    if geo["status"] == "OK":
                        st.session_state[map_key] = {
                            "lat": geo["latitude"], "lon": geo["longitude"]
                        }
                        st.session_state[f"pin_raw_{idx}"] = (
                            f"{geo['latitude']},{geo['longitude']}"
                        )
                        st.success(f"✅ {geo.get('matched_address','')[:70]}")
                        st.rerun()
                    else:
                        st.error("Not found — try Google Maps below.")

        # Google Maps
        with st.expander("🌐 Open in Google Maps"):
            q = urllib.parse.quote(
                f"{row.get('facility','')} {row.get('city','')} {row.get('country','')}"
            )
            st.caption("Right-click facility → 'What's here?' for coordinates.")
            st.markdown(
                f'<a href="https://www.google.com/maps/search/{q}" target="_blank">'
                f'<button style="width:100%;padding:7px;background:#1a73e8;color:white;'
                f'border:none;border-radius:5px;cursor:pointer;">🗺️ Open Google Maps'
                f'</button></a>',
                unsafe_allow_html=True,
            )

        # Coordinate entry
        with st.expander("📐 Use pin / enter coordinates", expanded=True):
            st.caption(
                "Copy Lat & Lon from the **blue box under the map**, paste here, "
                "then click ✔ Use these."
            )
            nl = st.number_input("Latitude",  value=float(cur_lat or 0.0),
                                  format="%.6f", key=f"nl_{idx}")
            no = st.number_input("Longitude", value=float(cur_lon or 0.0),
                                  format="%.6f", key=f"no_{idx}")
            if st.button("✔  Use these coordinates", key=f"chk_{idx}",
                         use_container_width=True, type="primary"):
                st.session_state[map_key] = {"lat": nl, "lon": no}
                st.session_state[f"pin_raw_{idx}"] = f"{nl},{no}"
                st.rerun()

        # Notes field
        st.markdown("**Notes** *(optional)*")
        note = st.text_area(
            "Add a note for this row",
            value=row.get("notes", ""),
            placeholder="e.g. Confirmed via company website · Pin placed on main building",
            key=f"note_{idx}",
            label_visibility="collapsed",
            height=80,
        )

        st.divider()
        st.markdown("**Your decision**")

        def _save(status, by):
            row["latitude"]     = st.session_state[map_key]["lat"]
            row["longitude"]    = st.session_state[map_key]["lon"]
            row["final_status"] = status
            row["verified_by"]  = by
            row["notes"]        = note
            st.session_state.reviewed.append(row)
            st.session_state.review_idx += 1
            if map_key in st.session_state:
                del st.session_state[map_key]
            st.rerun()

        if st.button("✅ Verify — looks correct", type="primary",
                     use_container_width=True, key=f"v_{idx}"):
            if not (st.session_state[map_key]["lat"] or
                    st.session_state[map_key]["lon"]):
                st.error("No coordinates set yet.")
            else:
                _save("Verified", "Human (satellite review)")

        if st.button("✏️ Save corrected coordinates",
                     use_container_width=True, key=f"c_{idx}"):
            if not (st.session_state[map_key]["lat"] or
                    st.session_state[map_key]["lon"]):
                st.error("No coordinates set yet.")
            else:
                _save("Corrected", "Human (corrected)")

        if st.button("❌ Reject — can't find it",
                     use_container_width=True, key=f"r_{idx}"):
            row["latitude"]      = None
            row["longitude"]     = None
            row["final_status"]  = "Rejected"
            row["review_reason"] = "Rejected in satellite review"
            row["notes"]         = note
            st.session_state.reviewed.append(row)
            st.session_state.review_idx += 1
            if map_key in st.session_state:
                del st.session_state[map_key]
            st.rerun()

        st.divider()
        if st.session_state.reviewed:
            dl = build_excel(
                st.session_state.reviewed,
                st.session_state.orig_columns,
            )
            st.download_button(
                "⬇️ Save progress", data=dl,
                file_name="geolocator_review_progress.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        st.caption(f"{idx} / {total_q} reviewed")

        if st.button("⏭ Skip rest → download",
                     use_container_width=True, key="skip_rev"):
            for r in queue[idx:]:
                r["final_status"]  = "Needs Review"
                r["review_reason"] = "Skipped"
                st.session_state.reviewed.append(r)
            st.session_state.phase = 4
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 4 — DOWNLOAD
# ─────────────────────────────────────────────────────────────────────────────
elif st.session_state.phase == 4:

    all_rows = st.session_state.reviewed
    total    = len(all_rows)
    vfd      = sum(1 for r in all_rows if r.get("final_status") in
                   ("Auto-High","Verified","Corrected"))
    manual   = sum(1 for r in all_rows if r.get("final_status") in
                   ("Needs Review","Rejected","No Match"))
    rooftop  = sum(1 for r in all_rows if r.get("location_type") == "ROOFTOP")

    st.subheader("✅ Done — Download your results")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total rows",            total)
    c2.metric("Have coordinates ✅",   vfd)
    c3.metric("Rooftop precision 🎯",  rooftop)
    c4.metric("Need manual lookup 📋", manual)

    st.divider()

    # Results preview
    df_show = pd.DataFrame(all_rows)
    show_cols = (st.session_state.orig_columns +
                 ["latitude","longitude","matched_address",
                  "final_status","confidence","location_type","notes"])
    show_cols = [c for c in show_cols if c in df_show.columns]
    st.dataframe(df_show[show_cols], use_container_width=True, height=360)

    st.divider()
    excel_bytes = build_excel(all_rows, st.session_state.orig_columns)
    base = getattr(st.session_state, "upload_name", "data").rsplit(".", 1)[0]

    st.download_button(
        "⬇️  Download GeoLocator Results",
        data=excel_bytes,
        file_name=f"{base}_GeoLocator.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
    st.caption(
        "3 sheets: **Geocoded Results** (all original columns + geo columns)  ·  "
        "**Needs Manual Review**  ·  **Summary**"
    )

    if st.button("🔄 Start over", use_container_width=True):
        for k in ["phase","batch_idx","all_results","review_queue","reviewed",
                  "review_idx","upload_df","upload_name","orig_columns","group_exists"]:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

st.divider()
st.caption("GeoLocator · Google Maps Geocoding API · Esri Satellite · 100% free")
