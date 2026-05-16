import warnings
import io
from pathlib import Path

import streamlit as st

warnings.filterwarnings("ignore")

st.set_page_config(page_title="File → TXT Converter", page_icon="📄")
st.title("📄 File → TXT Converter")
st.write("Upload one or more Excel (.xls / .xlsx) or PDF files and download a clean text file for each.")

uploaded_files = st.file_uploader(
    "Choose files",
    type=["xls", "xlsx", "pdf"],
    accept_multiple_files=True
)


# ── Excel helpers ─────────────────────────────────────────────────────────────

def extract_data_excel(df):
    import pandas as pd
    data_start = None
    for i, row in df.iterrows():
        col0_empty = pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == ""
        other_filled = any(not pd.isna(v) and str(v).strip() != "" for v in row.iloc[1:])
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


def write_excel_sheet(f, df, tabulate):
    import pandas as pd
    meta_lines, real = extract_data_excel(df)

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
    data = real.loc[header_idx + 1:].copy()
    data = data[data.apply(lambda r: r.str.strip().any(), axis=1)]

    if data.empty:
        for v in headers:
            if v.strip():
                f.write(v.strip() + "\n")
        return

    data.columns = headers
    table = tabulate(data, headers="keys", tablefmt="simple", showindex=False)
    f.write(table + "\n")


def convert_excel(file_bytes, filename, status_callback):
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

    for i, sheet_name in enumerate(sheet_names):
        status_callback(f"[{filename}] Sheet {i+1}/{len(sheet_names)}: {sheet_name}")
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None, engine=engine)

        output.write(f"\n{'=' * 80}\n")
        output.write(f"  SHEET: {sheet_name}\n")
        output.write(f"{'=' * 80}\n\n")

        if df.empty:
            output.write("  (empty sheet)\n")
            continue

        write_excel_sheet(output, df, tabulate)
        output.write("\n")

    return output.getvalue()


# ── PDF helpers ───────────────────────────────────────────────────────────────

def convert_pdf(file_bytes, filename, status_callback):
    import pdfplumber
    from tabulate import tabulate

    output = io.StringIO()
    output.write("=" * 80 + "\n")
    output.write(f"Source file : {filename}\n")
    output.write("=" * 80 + "\n")

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        total = len(pdf.pages)
        output.write(f"Pages       : {total}\n\n")

        for i, page in enumerate(pdf.pages):
            status_callback(f"[{filename}] Page {i+1}/{total}")

            output.write(f"\n{'=' * 80}\n")
            output.write(f"  PAGE {i+1}\n")
            output.write(f"{'=' * 80}\n\n")

            # Extract tables first
            tables = page.extract_tables()
            table_texts = set()

            if tables:
                for t in tables:
                    if not t:
                        continue
                    # Use first row as header if it looks like one
                    headers = [str(c).strip() if c else "" for c in t[0]]
                    rows = [[str(cell).strip() if cell else "" for cell in row] for row in t[1:]]
                    rows = [r for r in rows if any(c for c in r)]

                    if rows:
                        try:
                            import pandas as pd
                            df = pd.DataFrame(rows, columns=headers)
                            table_str = tabulate(df, headers="keys", tablefmt="simple", showindex=False)
                            output.write(table_str + "\n\n")
                            # Track cell text so we don't double-print
                            for row in t:
                                for cell in row:
                                    if cell:
                                        table_texts.add(str(cell).strip())
                        except Exception:
                            for row in t:
                                output.write("  ".join(str(c) if c else "" for c in row) + "\n")
                            output.write("\n")

            # Extract remaining text (skip what's already in tables)
            text = page.extract_text(x_tolerance=3, y_tolerance=3)
            if text:
                lines = text.splitlines()
                for line in lines:
                    stripped = line.strip()
                    if stripped and stripped not in table_texts:
                        output.write(stripped + "\n")
            output.write("\n")

    return output.getvalue()


# ── Main UI ───────────────────────────────────────────────────────────────────

if uploaded_files:
    file_count = len(uploaded_files)
    st.success(f"✅ {file_count} file{'s' if file_count > 1 else ''} loaded: {', '.join(f.name for f in uploaded_files)}")

    if st.button(f"▶ Convert {file_count} file{'s' if file_count > 1 else ''}"):
        status = st.empty()
        results = {}

        for uploaded in uploaded_files:
            file_bytes = uploaded.read()
            fname = uploaded.name
            ext = Path(fname).suffix.lower()

            try:
                if ext in [".xls", ".xlsx"]:
                    text = convert_excel(file_bytes, fname, lambda msg: status.info(msg))
                elif ext == ".pdf":
                    text = convert_pdf(file_bytes, fname, lambda msg: status.info(msg))
                else:
                    st.warning(f"Skipped unsupported file: {fname}")
                    continue

                results[fname] = text

            except Exception as e:
                st.error(f"Error converting {fname}: {e}")

        status.empty()

        if results:
            st.success(f"✅ Done! {len(results)} file{'s' if len(results) > 1 else ''} converted.")

            for fname, text in results.items():
                out_name = Path(fname).stem + ".txt"
                st.download_button(
                    label=f"⬇️ Download {out_name}",
                    data=text,
                    file_name=out_name,
                    mime="text/plain",
                    key=f"dl_{fname}"
                )

            # Preview the first result
            first_fname = list(results.keys())[0]
            st.subheader(f"Preview: {Path(first_fname).stem}.txt (first 100 lines)")
            preview = "\n".join(results[first_fname].splitlines()[:100])
            st.code(preview, language=None)
