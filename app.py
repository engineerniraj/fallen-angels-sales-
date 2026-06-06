import streamlit as st
import google.generativeai as genai
from PIL import Image
import pandas as pd
import json
import io
import os
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(
    page_title="Handwritten Invoice Extractor",
    page_icon="🧾",
    layout="wide"
)

st.title("Handwritten Invoice Data Extractor")
st.write("Upload invoice images and extract structured data using Gemini API.")


def get_api_key():
    try:
        return st.secrets["GEMINI_API_KEY"]
    except Exception:
        return os.getenv("GEMINI_API_KEY")


api_key = get_api_key()

if not api_key:
    st.error("GEMINI_API_KEY is missing. Add it in Streamlit Secrets or your local .env file.")
    st.stop()

genai.configure(api_key=api_key)
model = genai.GenerativeModel("gemini-2.5-flash")


EXTRACTION_PROMPT = """
You are an expert invoice data extraction system.

Extract data from the invoice image, including handwritten and printed text.
Return only valid JSON. Do not include markdown, explanations, or extra text.

Use this exact JSON structure:
{
  "invoice_number": null,
  "invoice_date": null,
  "vendor_name": null,
  "vendor_address": null,
  "customer_name": null,
  "customer_address": null,
  "subtotal": null,
  "tax_amount": null,
  "discount": null,
  "total_amount": null,
  "currency": null,
  "payment_status": null,
  "line_items": [
    {
      "description": null,
      "quantity": null,
      "unit_price": null,
      "amount": null
    }
  ],
  "notes": null
}

Rules:
- If a value is unreadable or missing, use null.
- Keep amounts as numbers when possible.
- Do not guess unclear handwriting.
- For dates, use the format found on the invoice if uncertain.
- Include all visible line items.
"""


def clean_json_response(text):
    text = text.strip()
    text = text.replace("```json", "").replace("```", "").strip()

    start = text.find("{")
    end = text.rfind("}")

    if start != -1 and end != -1:
        text = text[start:end + 1]

    return text


def extract_invoice_data(uploaded_file):
    image = Image.open(uploaded_file).convert("RGB")

    response = model.generate_content([EXTRACTION_PROMPT, image])
    raw_text = response.text
    cleaned_text = clean_json_response(raw_text)

    try:
        data = json.loads(cleaned_text)
        return data, raw_text, None
    except json.JSONDecodeError as error:
        return None, raw_text, str(error)


def flatten_invoice_data(filename, data):
    line_items = data.get("line_items") or []

    if not line_items:
        line_items = [{
            "description": None,
            "quantity": None,
            "unit_price": None,
            "amount": None
        }]

    rows = []
    for item in line_items:
        rows.append({
            "file_name": filename,
            "invoice_number": data.get("invoice_number"),
            "invoice_date": data.get("invoice_date"),
            "vendor_name": data.get("vendor_name"),
            "vendor_address": data.get("vendor_address"),
            "customer_name": data.get("customer_name"),
            "customer_address": data.get("customer_address"),
            "subtotal": data.get("subtotal"),
            "tax_amount": data.get("tax_amount"),
            "discount": data.get("discount"),
            "total_amount": data.get("total_amount"),
            "currency": data.get("currency"),
            "payment_status": data.get("payment_status"),
            "item_description": item.get("description"),
            "item_quantity": item.get("quantity"),
            "item_unit_price": item.get("unit_price"),
            "item_amount": item.get("amount"),
            "notes": data.get("notes")
        })

    return rows


def dataframe_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted Invoices")
    output.seek(0)
    return output


uploaded_files = st.file_uploader(
    "Upload invoice image files",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True
)

if uploaded_files:
    all_rows = []
    extracted_json = {}
    failed_files = []

    if st.button("Extract Invoice Data"):
        progress_bar = st.progress(0)

        for index, uploaded_file in enumerate(uploaded_files):
            st.subheader(f"Processing: {uploaded_file.name}")

            col1, col2 = st.columns([1, 2])

            with col1:
                image = Image.open(uploaded_file).convert("RGB")
                st.image(image, caption=uploaded_file.name, use_container_width=True)

            with col2:
                try:
                    uploaded_file.seek(0)
                    data, raw_text, error = extract_invoice_data(uploaded_file)

                    if error:
                        st.error(f"Could not parse JSON for {uploaded_file.name}: {error}")
                        st.text_area("Raw Gemini Output", raw_text, height=220)
                        failed_files.append(uploaded_file.name)
                    else:
                        st.success("Extraction completed")
                        st.json(data)
                        extracted_json[uploaded_file.name] = data
                        all_rows.extend(flatten_invoice_data(uploaded_file.name, data))

                except Exception as error:
                    st.error(f"Could not extract {uploaded_file.name}: {error}")
                    failed_files.append(uploaded_file.name)

            progress_bar.progress((index + 1) / len(uploaded_files))

        if all_rows:
            st.divider()
            st.header("Extracted Data Table")

            df = pd.DataFrame(all_rows)
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

            csv_data = edited_df.to_csv(index=False).encode("utf-8")
            excel_data = dataframe_to_excel(edited_df)
            json_data = json.dumps(extracted_json, indent=2).encode("utf-8")

            col_csv, col_excel, col_json = st.columns(3)

            with col_csv:
                st.download_button(
                    "Download CSV",
                    data=csv_data,
                    file_name="extracted_invoices.csv",
                    mime="text/csv"
                )

            with col_excel:
                st.download_button(
                    "Download Excel",
                    data=excel_data,
                    file_name="extracted_invoices.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col_json:
                st.download_button(
                    "Download JSON",
                    data=json_data,
                    file_name="extracted_invoices.json",
                    mime="application/json"
                )

        if failed_files:
            st.warning("Some files could not be processed:")
            for failed_file in failed_files:
                st.write(f"- {failed_file}")
else:
    st.info("Upload one or more invoice images to begin.")

st.sidebar.header("Setup")
st.sidebar.write("Required Streamlit secret:")
st.sidebar.code('GEMINI_API_KEY = "your_api_key_here"', language="toml")
st.sidebar.write("Current model: gemini-2.5-flash")
