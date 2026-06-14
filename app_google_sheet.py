import re
from pathlib import Path
import tempfile

import requests
import streamlit as st
from openpyxl import load_workbook

from converter_faktur_coretax_v2_2 import (
    read_sheet,
    build_xml,
    SHEET_FAKTUR,
    SHEET_DETAIL,
    FAKTUR_HEADER_ROW,
    DETAIL_HEADER_ROW,
)


# ============================
# GOOGLE SHEET HELPERS
# ============================

def extract_google_sheet_id(url_or_id: str) -> str:
    """Extract Google Spreadsheet ID from full URL or return the ID as-is."""
    text = (url_or_id or "").strip()
    if not text:
        raise ValueError("URL/ID Google Sheet masih kosong.")

    patterns = [
        r"/spreadsheets/d/([a-zA-Z0-9-_]+)",
        r"id=([a-zA-Z0-9-_]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)

    # If user pasted only the spreadsheet ID
    if re.fullmatch(r"[a-zA-Z0-9-_]+", text):
        return text

    raise ValueError("Format URL Google Sheet tidak dikenali.")


def download_google_sheet_as_xlsx(sheet_url_or_id: str) -> Path:
    """Download a Google Sheet as XLSX and return temporary file path.

    Requirement: the Google Sheet must be accessible with the link
    or published/available for the app environment.
    """
    sheet_id = extract_google_sheet_id(sheet_url_or_id)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    response = requests.get(export_url, timeout=60)
    response.raise_for_status()

    content_type = response.headers.get("content-type", "").lower()
    if "text/html" in content_type and b"<html" in response.content[:500].lower():
        raise PermissionError(
            "Google Sheet tidak bisa diakses. Pastikan sharing diset ke 'Anyone with the link can view'."
        )

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(response.content)
    tmp.close()
    return Path(tmp.name)


def convert_xlsx_to_xml_bytes(xlsx_path: Path) -> bytes:
    wb = load_workbook(xlsx_path, data_only=True)

    faktur = read_sheet(wb[SHEET_FAKTUR], FAKTUR_HEADER_ROW)
    detail = read_sheet(wb[SHEET_DETAIL], DETAIL_HEADER_ROW)

    if not faktur:
        raise ValueError("Sheet Faktur kosong atau header/baris data tidak sesuai.")
    if not detail:
        raise ValueError("Sheet DetailFaktur kosong atau header/baris data tidak sesuai.")

    xml_tree = build_xml(faktur, detail)
    xml_output_path = xlsx_path.with_suffix(".xml")
    xml_tree.write(xml_output_path, encoding="utf-8", xml_declaration=True)

    return xml_output_path.read_bytes()


# ============================
# STREAMLIT UI
# ============================

st.title("🚀 Converter Faktur CoreTax Online")
st.write("Ambil data langsung dari Google Sheet atau upload file Excel template CoreTax, lalu konversi menjadi XML.")

mode = st.radio(
    "Pilih sumber data",
    ["Google Sheet", "Upload Excel"],
    horizontal=True,
)

if mode == "Google Sheet":
    sheet_url = st.text_input("Masukkan URL Google Sheet atau Spreadsheet ID")

    st.caption(
        "Catatan: untuk metode sederhana ini, Google Sheet perlu diset 'Anyone with the link can view'."
    )

    if st.button("Generate XML dari Google Sheet"):
        try:
            with st.spinner("Mengambil Google Sheet dan membuat XML..."):
                xlsx_path = download_google_sheet_as_xlsx(sheet_url)
                xml_bytes = convert_xlsx_to_xml_bytes(xlsx_path)

            st.success("XML berhasil dibuat dari Google Sheet!")
            st.download_button(
                label="⬇ Download XML",
                data=xml_bytes,
                file_name="output.xml",
                mime="application/xml",
            )
        except Exception as e:
            st.error(f"Terjadi error saat memproses Google Sheet: {e}")

else:
    uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded_file.read())
                temp_path = Path(tmp.name)

            with st.spinner("Memproses file Excel..."):
                xml_bytes = convert_xlsx_to_xml_bytes(temp_path)

            st.success("XML berhasil dibuat dari file Excel!")
            st.download_button(
                label="⬇ Download XML",
                data=xml_bytes,
                file_name="output.xml",
                mime="application/xml",
            )
        except Exception as e:
            st.error(f"Terjadi error saat memproses file: {e}")
