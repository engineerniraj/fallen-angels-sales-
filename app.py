import io
from copy import copy

import openpyxl
import streamlit as st

st.set_page_config(page_title="Fallen Angels Sales Processor", page_icon="📊", layout="wide")

COLUMN_I = 9

SHOWS = [
    "Tuesday",
    "Wednesday 2 PM",
    "Wednesday 8 PM",
    "Thursday",
    "Friday",
    "Saturday 2 PM",
    "Saturday 8 PM",
    "Sunday",
]

PRODUCTS = [
    {"label": "Poster", "keywords": ["poster"]},
    {"label": "Magnet", "keywords": ["magnet"]},
    {"label": "Lapel Pin", "keywords": ["lapel", "pin"]},
    {"label": "Keychain", "keywords": ["keychain", "key chain"]},
    {"label": "Mug", "keywords": ["mug"]},
    {"label": "Tote", "keywords": ["tote"]},
    {"label": "Logo Tee S", "keywords": ["logo", "small", " s"]},
    {"label": "Logo Tee M", "keywords": ["logo", "medium", " m"]},
    {"label": "Logo Tee L", "keywords": ["logo", "large", " l"]},
    {"label": "Logo Tee XL", "keywords": ["logo", "xl"]},
    {"label": "Lapse Tee S", "keywords": ["lapse", "small", " s"]},
    {"label": "Lapse Tee M", "keywords": ["lapse", "medium", " m"]},
    {"label": "Lapse Tee L", "keywords": ["lapse", "large", " l"]},
    {"label": "Lapse Tee XL", "keywords": ["lapse", "xl"]},
    {"label": "Lapse Tee XXL", "keywords": ["lapse", "xxl", "2xl"]},
    {"label": "Hoodie S", "keywords": ["hoodie", "small", " s"]},
    {"label": "Hoodie M", "keywords": ["hoodie", "medium", " m"]},
    {"label": "Hoodie L", "keywords": ["hoodie", "large", " l"]},
    {"label": "Hoodie XL", "keywords": ["hoodie", "xl"]},
]


def normalize(value):
    return str(value or "").strip().lower()


def find_product_row(ws, product):
    keywords = product["keywords"]
    best_row = None
    best_score = 0
    for row in range(1, ws.max_row + 1):
        text = " ".join(normalize(ws.cell(row=row, column=col).value) for col in range(1, 9))
        if not text:
            continue
        score = sum(1 for keyword in keywords if keyword in text)
        if score > best_score:
            best_score = score
            best_row = row
    return best_row if best_score else None


def copy_cell_style(source, target):
    if source.has_style:
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy(source.protection)


def write_retail_values(workbook, sheet_entries):
    summary = []
    for index, show_name in enumerate(SHOWS):
        if index >= len(workbook.worksheets):
            continue
        ws = workbook.worksheets[index]
        entries = sheet_entries.get(show_name, {})
        entered_count = 0
        missing = []
        for product in PRODUCTS:
            qty = int(entries.get(product["label"], 0) or 0)
            if qty == 0:
                continue
            row = find_product_row(ws, product)
            if row:
                cell = ws.cell(row=row, column=COLUMN_I)
                copy_cell_style(ws.cell(row=row, column=COLUMN_I + 1), cell)
                cell.value = qty
                entered_count += 1
            else:
                missing.append(product["label"])
        summary.append({"show": show_name, "worksheet": ws.title, "items": entered_count, "missing": missing})
    return summary


def build_download(master_file, sheet_entries):
    workbook = openpyxl.load_workbook(master_file)
    summary = write_retail_values(workbook, sheet_entries)
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output, summary


def main():
    st.title("Fallen Angels Sales Report Processor")
    st.caption("No API key needed. Staff enter the retail quantities from each invoice, then download the finished Excel file.")

    with st.expander("How to use", expanded=True):
        st.markdown("""
        1. Upload the master Excel file.
        2. For each show, type the **RETAIL quantity sold** from the invoice.
        3. Click **Create completed Excel file**.
        4. Download the ready master file.
        """)

    master_file = st.file_uploader("Upload master Excel file", type=["xlsx"])
    st.divider()

    sheet_entries = {}
    tabs = st.tabs(SHOWS)
    for show_name, tab in zip(SHOWS, tabs):
        with tab:
            st.subheader(show_name)
            cols = st.columns(3)
            values = {}
            for idx, product in enumerate(PRODUCTS):
                with cols[idx % 3]:
                    values[product["label"]] = st.number_input(
                        product["label"],
                        min_value=0,
                        max_value=999,
                        value=0,
                        step=1,
                        key=f"{show_name}-{product['label']}",
                    )
            sheet_entries[show_name] = values

    st.divider()
    if st.button("Create completed Excel file", type="primary", disabled=master_file is None):
        try:
            output, summary = build_download(master_file, sheet_entries)
            st.success("Done. Download your completed master file below.")
            st.download_button(
                "Download completed Excel file",
                data=output,
                file_name="Fallen_Angels_Mastersheet_Completed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.write("Update summary")
            st.dataframe(summary, use_container_width=True)
        except Exception as exc:
            st.error(f"Could not create the file: {exc}")


if __name__ == "__main__":
    main()
