import warnings
import io
from pathlib import Path

import streamlit as st

warnings.filterwarnings("ignore")

st.set_page_config(page_title="XLS → TXT Converter", page_icon="📄")
st.title("📄 XLS / XLSX → TXT Converter")
st.write("Upload an Excel file and download a clean, readable text file.")

uploaded = st.file_uploader("Choose an Excel file", type=["xls", "xlsx"])


def extract_data(df):
    import pandas as pd

    data_start = None
    for i, row in df.iterrows():
        col0_empty = pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == ""
        other_filled = any(
            not pd.isna(v) and str(v).strip() != "" for v in row.iloc[1:]
        )
        if col0_empty and other_filled:
            data_start = i
            break

    if data_start is None:
        col0 = df.iloc[:, 0].dropna()
        col0 = col0[col0.astype(str).str.strip() != ""]
        return col0.astype(str).tolist(), []

    real = df.loc[data_start:, df.columns[1:]].copy()
    real = real.dropna(how="all").dropna(axis=1, how="all")

    boilerplate = {"Created by EDGAR Online, Inc.", "Table Of Contents"}
    meta_lines = []
    for v in df.loc[: data_start - 1 if data_start > 0 else 0, df.columns[0]]:
        s = str(v).strip()
        if s and s not in boilerplate and str(v) != "nan":
            meta_lines.append(s)

    return meta_lines, real


def write_sheet(f, df, tabulate):
    import pandas as pd

    meta_lines, real = extract_data(df)

    for line in meta_lines:
        f.write(line + "\n")
    if meta_lines:
        f.write("\n")

    if isinstance(real, list):
        for line in real:
            f.write(line + "\n")
        return

    if real.empty:
        f.write("  (no data)\n")
        return

    real = real.fillna("").astype(str)
    filled = real.apply(lambda r: r.str.strip().astype(bool).sum(), axis=1)
    header_idx = filled.idxmax()
    headers = real.loc[header_idx].tolist()
    data = real.loc[header_idx + 1 :].copy()
    data = data[data.apply(lambda r: r.str.strip().any(), axis=1)]

    if data.empty:
        for v in headers:
            if v.strip():
                f.write(v.strip() + "\n")
        return

    data.columns = headers
    table = tabulate(data, headers="keys", tablefmt="simple", showindex=False)
    f.write(table + "\n")


def convert(file_bytes, filename):
    import pandas as pd
    from tabulate import tabulate

    ext = Path(filename).suffix.lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"

    xl = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
    sheet_names = xl.sheet_names

    output = io.StringIO()
    output.write("=" * 80 + "\n")
    output.write(f"Source file : {filename}\n")
    output.write(f"Sheets      : {len(sheet_names)}\n")
    output.write("=" * 80 + "\n")

    progress = st.progress(0, text="Starting…")

    for i, sheet_name in enumerate(sheet_names):
        progress.progress((i + 1) / len(sheet_names), text=f"Processing: {sheet_name}")

        df = pd.read_excel(xl, sheet_name=sheet_name, header=None, engine=engine)

        output.write(f"\n{'=' * 80}\n")
        output.write(f"  SHEET: {sheet_name}\n")
        output.write(f"{'=' * 80}\n\n")

        if df.empty:
            output.write("  (empty sheet)\n")
            continue

        write_sheet(output, df, tabulate)
        output.write("\n")

    progress.empty()
    return output.getvalue()


if uploaded:
    file_bytes = uploaded.read()
    st.success(f"✅ Loaded **{uploaded.name}**")

    if st.button("▶ Convert to TXT"):
        with st.spinner("Converting…"):
            try:
                result = convert(file_bytes, uploaded.name)
                out_name = Path(uploaded.name).stem + ".txt"

                st.download_button(
                    label="⬇️ Download TXT file",
                    data=result,
                    file_name=out_name,
                    mime="text/plain",
                )

                st.subheader("Preview (first 100 lines)")
                preview = "\n".join(result.splitlines()[:100])
                st.code(preview, language=None)

            except Exception as e:
                st.error(f"Error: {e}")
