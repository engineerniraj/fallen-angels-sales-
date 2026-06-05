import hmac
import io
import re
import zipfile
from copy import copy

import openpyxl
import pytesseract
import streamlit as st
from PIL import Image
from pdf2image import convert_from_bytes
from pillow_heif import register_heif_opener

register_heif_opener()

st.set_page_config(page_title="Niraj Excel Tools", page_icon="📊", layout="wide")


def check_password():
    if st.session_state.get("authenticated"):
        return True

    st.title("Niraj Excel Tools")
    st.subheader("Login required")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login", type="primary"):
        valid_username = st.secrets.get("APP_USERNAME", "")
        valid_password = st.secrets.get("APP_PASSWORD", "")

        username_ok = hmac.compare_digest(username, valid_username)
        password_ok = hmac.compare_digest(password, valid_password)

        if username_ok and password_ok:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect username or password")

    return False


def logout_button():
    with st.sidebar:
        if st.button("Logout"):
            st.session_state["authenticated"] = False
            st.rerun()


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


def images_from_upload(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        return convert_from_bytes(file_bytes, dpi=220)
    image = Image.open(io.BytesIO(file_bytes))
    return [image]


def extract_text_from_invoice(uploaded_file):
    pages = images_from_upload(uploaded_file)
    text_parts = []
    for page in pages:
        image = page.convert("RGB")
        text = pytesseract.image_to_string(image, config="--psm 6")
        text_parts.append(text)
    return "\n".join(text_parts)


def detect_show_from_filename(filename):
    name = filename.lower()
    if "tuesday" in name:
        return "Tuesday"
    if "wednesday" in name and ("2" in name or "2pm" in name or "2 pm" in name):
        return "Wednesday 2 PM"
    if "wednesday" in name and ("8" in name or "8pm" in name or "8 pm" in name):
        return "Wednesday 8 PM"
    if "thursday" in name:
        return "Thursday"
    if "friday" in name:
        return "Friday"
    if "saturday" in name and ("2" in name or "2pm" in name or "2 pm" in name):
        return "Saturday 2 PM"
    if "saturday" in name and ("8" in name or "8pm" in name or "8 pm" in name):
        return "Saturday 8 PM"
    if "sunday" in name:
        return "Sunday"
    return None


def parse_qty_from_line(line):
    numbers = [int(match) for match in re.findall(r"\b\d+\b", line)]
    if len(numbers) >= 2:
        return max(numbers[0] - numbers[1], 0)
    if len(numbers) == 1:
        return numbers[0]
    return 0


def extract_entries_from_text(text):
    entries = {product["label"]: 0 for product in PRODUCTS}
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for line in lines:
        lowered = normalize(line)
        for product in PRODUCTS:
            if any(keyword.strip() and keyword.strip() in lowered for keyword in product["keywords"]):
                qty = parse_qty_from_line(line)
                if qty:
                    entries[product["label"]] += qty
    return entries


def automated_ocr_processor():
    st.header("Automated Invoice OCR to Excel")
    st.caption("Free OCR version: upload master Excel + invoice images/PDFs, then download the filled Excel file. No AI API key needed.")
    st.warning("This uses free OCR, not AI. It works best with clear printed invoices. Handwriting, blurry photos, or unusual layouts may need checking.")

    master_file = st.file_uploader("Upload master Excel file", type=["xlsx"], key="ocr-master")
    invoice_files = st.file_uploader(
        "Upload invoice images/PDFs",
        type=["jpg", "jpeg", "png", "heic", "heif", "pdf"],
        accept_multiple_files=True,
        key="ocr-invoices",
    )

    if not master_file or not invoice_files:
        st.info("Upload the master Excel file and all invoice files to begin.")
        return

    if st.button("Extract invoices and create Excel", type="primary"):
        sheet_entries = {show: {product["label"]: 0 for product in PRODUCTS} for show in SHOWS}
        extracted_rows = []
        with st.spinner("Reading invoices with OCR..."):
            for invoice in invoice_files:
                show_name = detect_show_from_filename(invoice.name)
                text = extract_text_from_invoice(invoice)
                entries = extract_entries_from_text(text)
                if show_name:
                    sheet_entries[show_name] = entries
                extracted_rows.append({
                    "file": invoice.name,
                    "detected_show": show_name or "Not detected from filename",
                    "extracted_text_preview": text[:800],
                    "entries": {k: v for k, v in entries.items() if v},
                })

        try:
            output, summary = build_download(master_file, sheet_entries)
            st.success("OCR complete. Download your completed master file below.")
            st.download_button(
                "Download completed Excel file",
                data=output,
                file_name="Fallen_Angels_Mastersheet_OCR_Completed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.subheader("Detected invoice data")
            st.dataframe(extracted_rows, use_container_width=True)
            st.subheader("Excel update summary")
            st.dataframe(summary, use_container_width=True)
        except Exception as exc:
            st.error(f"Could not create the Excel file: {exc}")


def fallen_angels_processor():
    st.header("Fallen Angels Sales Report Processor")
    st.caption("No API key needed. Staff enter the retail quantities from each invoice, then download the finished Excel file.")

    with st.expander("How to use", expanded=True):
        st.markdown("""
        1. Upload the master Excel file.
        2. For each show, type the **RETAIL quantity sold** from the invoice.
        3. Click **Create completed Excel file**.
        4. Download the ready master file.
        """)

    master_file = st.file_uploader("Upload master Excel file", type=["xlsx"], key="fallen-angels-master")
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


def simple_excel_column_updater():
    st.header("Simple Excel Column Updater")
    st.caption("Manual tool for future tasks: upload an Excel file, choose a sheet/column/cell range, and fill one value down the range.")
    st.info("This is a starter task. Tell me your next exact workflow and I can customize it inside this same app.")

    excel_file = st.file_uploader("Upload Excel file", type=["xlsx"], key="simple-excel-file")
    sheet_name = st.text_input("Sheet name", value="Sheet1")
    column_letter = st.text_input("Column letter to update", value="I", max_chars=3)
    start_row = st.number_input("Start row", min_value=1, value=2, step=1)
    end_row = st.number_input("End row", min_value=1, value=10, step=1)
    value = st.text_input("Value to write", value="")

    if st.button("Create updated Excel", disabled=excel_file is None):
        try:
            wb = openpyxl.load_workbook(excel_file)
            if sheet_name not in wb.sheetnames:
                st.error(f"Sheet '{sheet_name}' not found. Available sheets: {', '.join(wb.sheetnames)}")
                return
            ws = wb[sheet_name]
            for row in range(int(start_row), int(end_row) + 1):
                ws[f"{column_letter.upper()}{row}"] = value
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            st.success("Updated file is ready.")
            st.download_button(
                "Download updated Excel",
                data=output,
                file_name="Updated_Excel_File.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"Could not update the file: {exc}")


def heic_to_jpg_converter():
    st.header("HEIC to JPG Converter")
    st.caption("Convert iPhone HEIC/HEIF photos to JPG. No API key needed; conversion happens inside the Streamlit app.")

    uploaded_files = st.file_uploader(
        "Upload HEIC/HEIF images",
        type=["heic", "heif"],
        accept_multiple_files=True,
        key="heic-files",
    )

    quality = st.slider("JPG quality", min_value=60, max_value=100, value=90, step=5)

    if not uploaded_files:
        st.info("Upload one or more .heic or .heif files to convert them into JPG.")
        return

    converted_files = []
    for uploaded_file in uploaded_files:
        try:
            image = Image.open(uploaded_file)
            if image.mode not in ("RGB", "L"):
                image = image.convert("RGB")

            output = io.BytesIO()
            image.save(output, format="JPEG", quality=quality, optimize=True)
            output.seek(0)

            base_name = uploaded_file.name.rsplit(".", 1)[0]
            jpg_name = f"{base_name}.jpg"
            converted_files.append((jpg_name, output.getvalue()))

            st.success(f"Converted: {uploaded_file.name} → {jpg_name}")
            st.image(output.getvalue(), caption=jpg_name, use_container_width=True)
            st.download_button(
                f"Download {jpg_name}",
                data=output.getvalue(),
                file_name=jpg_name,
                mime="image/jpeg",
                key=f"download-{jpg_name}",
            )
        except Exception as exc:
            st.error(f"Could not convert {uploaded_file.name}: {exc}")

    if len(converted_files) > 1:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for file_name, file_bytes in converted_files:
                zip_file.writestr(file_name, file_bytes)
        zip_buffer.seek(0)
        st.download_button(
            "Download all JPG files as ZIP",
            data=zip_buffer,
            file_name="converted_jpg_files.zip",
            mime="application/zip",
        )


def main():
    if not check_password():
        return

    logout_button()
    st.title("Niraj Excel Tools")
    st.caption("One free no-API app for manual and fixed Excel tasks.")

    task = st.sidebar.selectbox(
        "Choose task",
        [
            "Automated invoice OCR to Excel",
            "Fallen Angels retail updater",
            "Simple Excel column updater",
            "HEIC to JPG Converter",
        ],
    )

    st.sidebar.info("No paid AI API key is used. Automated OCR is free but may need checking for messy invoices.")

    if task == "Automated invoice OCR to Excel":
        automated_ocr_processor()
    elif task == "Fallen Angels retail updater":
        fallen_angels_processor()
    elif task == "Simple Excel column updater":
        simple_excel_column_updater()
    elif task == "HEIC to JPG Converter":
        heic_to_jpg_converter()


if __name__ == "__main__":
    main()
