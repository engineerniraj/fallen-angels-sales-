import streamlit as st
import anthropic
import base64
import json
import copy
import io
from openpyxl import load_workbook

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Fallen Angels · Sales Processor",
    page_icon="🎭",
    layout="centered",
)

# ── Styles ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 56px;
    letter-spacing: 6px;
    background: linear-gradient(135deg, #c8a96e, #e8c98e);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    text-align: center;
    line-height: 1;
    margin-bottom: 4px;
}
.subtitle {
    text-align: center;
    color: #888;
    font-size: 13px;
    letter-spacing: 3px;
    text-transform: uppercase;
    margin-bottom: 40px;
}
.step-label {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 20px;
    letter-spacing: 2px;
    color: #c8a96e;
    margin-bottom: 4px;
}
.result-card {
    background: #1a1a1a;
    border: 1px solid #2a2a2a;
    border-radius: 10px;
    padding: 16px 20px;
    margin: 8px 0;
}
.result-card .day { font-family: 'Bebas Neue', sans-serif; font-size: 20px; color: #c8a96e; letter-spacing: 2px; }
.result-card .meta { font-size: 13px; color: #888; margin-top: 2px; }
.result-card .total { font-size: 22px; font-weight: 600; color: #e8c98e; float: right; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="title">FALLEN ANGELS</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Sales Report Processor</div>', unsafe_allow_html=True)

# ── Row mapping (col I = retail units sold, Position #1 IN) ──────────────────
ROW_MAP = {
    "Poster": 24, "Magnet": 25, "LapelPin": 26, "Keychain": 28,
    "Mug": 29, "Tote": 30,
    "Logo_S": 73, "Logo_M": 74, "Logo_L": 75, "Logo_XL": 76, "Logo_XXL": 77,
    "Lapse_S": 79, "Lapse_M": 80, "Lapse_L": 81, "Lapse_XL": 82, "Lapse_XXL": 83,
    "Hoodie_S": 85, "Hoodie_M": 86, "Hoodie_L": 87, "Hoodie_XL": 88, "Hoodie_XXL": 89,
}

SHEET_LABELS = {
    "#1": "Tuesday", "#2": "Wed 2PM", "#3": "Wed 8PM", "#4": "Thursday",
    "#5": "Friday", "#6": "Sat 2PM", "#7": "Sat 8PM", "#8": "Sunday",
}

SYSTEM_PROMPT = """You are a sales data extraction assistant. Given an image or PDF of a handwritten
"Fallen Angels Sales Sheet", extract the number of RETAIL units sold for each item.

Retail units sold = the number in the RETAIL column on the right half of the sheet
(it equals START quantity minus END quantity for each row).

Return ONLY valid JSON with no markdown or explanation:
{
  "seller": "name or unknown",
  "day": "day and time e.g. Saturday 8PM",
  "date": "date e.g. May 9",
  "total_retail": 620,
  "items": {
    "Poster": 0, "Magnet": 0, "LapelPin": 0, "Keychain": 0,
    "Mug": 0, "Tote": 0,
    "Logo_S": 0, "Logo_M": 0, "Logo_L": 0, "Logo_XL": 0, "Logo_XXL": 0,
    "Lapse_S": 0, "Lapse_M": 0, "Lapse_L": 0, "Lapse_XL": 0, "Lapse_XXL": 0,
    "Hoodie_S": 0, "Hoodie_M": 0, "Hoodie_L": 0, "Hoodie_XL": 0, "Hoodie_XXL": 0
  }
}"""


def match_sheet(day_str: str, used: set) -> int:
    """Return 0-based sheet index from day string."""
    d = day_str.lower()
    candidates = []
    if "tuesday" in d:                                      candidates = [0]
    elif "wed" in d and ("2" in d or "12" in d):           candidates = [1]
    elif "wed" in d and ("7" in d or "8" in d):            candidates = [2]
    elif "wed" in d:                                        candidates = [1, 2]
    elif "thu" in d:                                        candidates = [3]
    elif "fri" in d:                                        candidates = [4]
    elif "sat" in d and ("2" in d or "12" in d):           candidates = [5]
    elif "sat" in d and ("7" in d or "8" in d):            candidates = [6]
    elif "sat" in d:                                        candidates = [5, 6]
    elif "sun" in d:                                        candidates = [7]

    for c in candidates:
        if c not in used:
            return c
    # fallback: first unused slot
    for i in range(8):
        if i not in used:
            return i
    return -1


def extract_invoice(client: anthropic.Anthropic, file_bytes: bytes, media_type: str) -> dict:
    if media_type == "application/pdf":
        content_block = {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf",
                       "data": base64.b64encode(file_bytes).decode()},
        }
    else:
        content_block = {
            "type": "image",
            "source": {"type": "base64", "media_type": media_type,
                       "data": base64.b64encode(file_bytes).decode()},
        }

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": [
            content_block,
            {"type": "text", "text": "Extract the retail sales data. Return only JSON."},
        ]}],
    )
    raw = response.content[0].text.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(raw)


def write_to_excel(wb, results: list) -> io.BytesIO:
    used_sheets = set()
    summary = []

    for res in results:
        idx = match_sheet(res.get("day", ""), used_sheets)
        if idx == -1:
            st.warning(f"⚠️ No sheet slot for: {res.get('_filename', '?')}")
            continue
        used_sheets.add(idx)

        sheet_name = f"#{idx + 1}"
        if sheet_name not in wb.sheetnames:
            st.warning(f"Sheet {sheet_name} not found in workbook.")
            continue

        ws = wb[sheet_name]
        units_written = 0
        for key, val in (res.get("items") or {}).items():
            row = ROW_MAP.get(key)
            if row and val:
                ws.cell(row=row, column=9, value=int(val))
                units_written += val

        summary.append({
            "sheet": sheet_name,
            "label": SHEET_LABELS.get(sheet_name, sheet_name),
            "day": res.get("day", "?"),
            "seller": res.get("seller", "?"),
            "total": res.get("total_retail"),
            "units": units_written,
            "filename": res.get("_filename", ""),
        })

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, summary


# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown('<div class="step-label">① Master Excel File</div>', unsafe_allow_html=True)
excel_file = st.file_uploader(
    "Upload the blank Mastersheet .xlsx",
    type=["xlsx"],
    label_visibility="collapsed",
)

st.markdown('<div class="step-label" style="margin-top:24px">② Invoice Images / PDFs</div>', unsafe_allow_html=True)
invoice_files = st.file_uploader(
    "Upload all handwritten sales sheets",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if invoice_files:
    st.caption(f"✅ {len(invoice_files)} invoice{'s' if len(invoice_files) > 1 else ''} ready")

st.markdown('<div class="step-label" style="margin-top:24px">③ Process & Download</div>', unsafe_allow_html=True)

ready = excel_file is not None and len(invoice_files) > 0
go = st.button("🎭  Process Sales Reports", disabled=not ready, use_container_width=True, type="primary")

if go:
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        st.error("❌ ANTHROPIC_API_KEY not set in Streamlit secrets. See deployment guide.")
        st.stop()

    client = anthropic.Anthropic(api_key=api_key)

    wb_bytes = excel_file.read()
    wb = load_workbook(io.BytesIO(wb_bytes))

    results = []
    progress = st.progress(0, text="Starting…")
    log = st.empty()

    for i, inv in enumerate(invoice_files):
        pct = int((i / len(invoice_files)) * 80)
        progress.progress(pct, text=f"Reading {inv.name}…")
        try:
            ext = inv.name.rsplit(".", 1)[-1].lower()
            mime = "application/pdf" if ext == "pdf" else ("image/png" if ext == "png" else "image/jpeg")
            data = extract_invoice(client, inv.read(), mime)
            data["_filename"] = inv.name
            results.append(data)
            log.success(f"✓ {inv.name}  →  {data.get('day','?')} | {data.get('seller','?')} | ${data.get('total_retail','?')}")
        except Exception as e:
            log.error(f"✗ {inv.name}: {e}")

    progress.progress(90, text="Writing to Excel…")
    out_bytes, summary = write_to_excel(wb, results)
    progress.progress(100, text="Done!")

    st.success(f"✅ Done! {len(summary)} sheet(s) updated.")

    # Summary cards
    for s in summary:
        total_str = f"${s['total']}" if s['total'] else "—"
        st.markdown(f"""
        <div class="result-card">
            <span class="total">{total_str}</span>
            <div class="day">{s['label']}  <span style='font-size:14px;color:#666'>({s['sheet']})</span></div>
            <div class="meta">👤 {s['seller']}  ·  {s['units']} units written  ·  {s['filename']}</div>
        </div>""", unsafe_allow_html=True)

    st.download_button(
        label="⬇️  Download Filled Mastersheet",
        data=out_bytes,
        file_name="Fallen_Angels_Mastersheet_Filled.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
