"""
Streamlit App - Flatten & Visualisasi Laporan Fa Detail (16 Segmen)
----------------------------------------------------------------------
Upload file laporan hierarkis (.xlsx) -> dapatkan file flat (.xlsx / .csv) +
visualisasi realisasi anggaran (total, per Kegiatan, per Output/RO, per
Komponen, per Akun).

Kolom output final (setelah dipangkas & disingkat):
    kd_satker, satker, kd_program, program, kd_keg, kegiatan,
    kd_kro, kro, kd_ro, ro, kd_komponen, komponen,
    kd_subkomponen, subkomponen, kd_akun, akun, kd_item, item,
    pagu_anggaran, realisasi_anggaran

Cara jalankan lokal (di PyCharm):
    1. pip install -r requirements.txt
    2. streamlit run app.py

Cara deploy publik (gratis):
    1. Push folder ini (app.py + requirements.txt) ke sebuah repo GitHub.
    2. Buka https://share.streamlit.io -> New app -> pilih repo & file app.py.
    3. Deploy. Anda akan dapat URL publik seperti
       https://nama-app-anda.streamlit.app
"""

import io
import re

import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# Logika parsing (membaca hierarki dari laporan asli)
# ----------------------------------------------------------------------


def g(ws, r, c):
    return ws.cell(row=r, column=c).value


def parse_workbook(file_obj):
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb.active

    satker_kode = g(ws, 6, 15)
    satker_nama = g(ws, 6, 16)

    ctx = {
        "program_kode": None, "program_nama": None,
        "kegiatan_kode": None, "kegiatan_nama": None,
        "output_kode": None, "output_nama": None,          # level KRO
        "suboutput_kode": None, "suboutput_nama": None,    # level RO (Output)
        "komponen_kode": None, "komponen_nama": None,
        "subkomponen_kode": None, "subkomponen_nama": None,
        "akun_kode": None, "akun_nama": None,
    }

    rows_info = []

    for r in range(10, ws.max_row + 1):
        b = g(ws, r, 2)   # Program / Kegiatan
        c = g(ws, r, 3)   # KRO / RO(Output)
        e = g(ws, r, 5)   # Komponen
        f = g(ws, r, 6)   # SubKomponen
        h = g(ws, r, 8)   # Akun
        n = g(ws, r, 14)  # Item (kode + nama digabung)

        if b and str(b).startswith("*"):
            continue

        level = None
        if b is not None:
            level = "kegiatan" if "." in str(b) else "program"
        elif c is not None:
            level = "suboutput" if "." in str(c) else "output"
        elif e is not None:
            level = "komponen"
        elif f is not None:
            level = "subkomponen"
        elif h is not None:
            level = "akun"
        elif n is not None:
            level = "item"
        else:
            continue

        item_kode = None
        item_nama = None

        if level == "program":
            ctx["program_kode"] = b
            ctx["program_nama"] = g(ws, r, 4)
        elif level == "kegiatan":
            ctx["kegiatan_kode"] = b
            ctx["kegiatan_nama"] = g(ws, r, 9)
        elif level == "output":
            ctx["output_kode"] = c
            ctx["output_nama"] = g(ws, r, 7)
        elif level == "suboutput":
            ctx["suboutput_kode"] = c
            ctx["suboutput_nama"] = g(ws, r, 11)
        elif level == "komponen":
            ctx["komponen_kode"] = e
            ctx["komponen_nama"] = g(ws, r, 10)
        elif level == "subkomponen":
            ctx["subkomponen_kode"] = f
            ctx["subkomponen_nama"] = g(ws, r, 12)
        elif level == "akun":
            pagu_check = g(ws, r, 17)
            if (
                ctx["akun_kode"] == h
                and rows_info
                and rows_info[-1]["level"] == "akun"
                and rows_info[-1]["pagu"] == pagu_check
            ):
                ctx["akun_nama"] = (
                    str(ctx["akun_nama"]).rstrip() + " " + str(g(ws, r, 13)).strip()
                )
                rows_info[-1]["ctx"] = dict(ctx)
                continue
            ctx["akun_kode"] = h
            ctx["akun_nama"] = g(ws, r, 13)
        elif level == "item":
            item_full = n
            m = re.match(r"^\s*(\d+)\.\s*(.*)$", str(item_full))
            if m:
                item_kode, item_nama = m.group(1), m.group(2)
                is_continuation = False
            else:
                is_continuation = True
            if is_continuation and rows_info and rows_info[-1]["level"] == "item":
                rows_info[-1]["item_nama"] = (
                    str(rows_info[-1]["item_nama"]).rstrip() + " " + str(item_full).strip()
                )
                continue

        pagu = g(ws, r, 17)
        real_sd = g(ws, r, 26)

        rows_info.append({
            "row": r,
            "level": level,
            "ctx": dict(ctx),
            "item_kode": item_kode if level == "item" else None,
            "item_nama": item_nama if level == "item" else None,
            "pagu": pagu,
            "real_sd": real_sd,
        })

    leaf_rows = []
    n_total = len(rows_info)
    for i, ri in enumerate(rows_info):
        if ri["level"] == "item":
            leaf_rows.append(ri)
        elif ri["level"] == "akun":
            if i + 1 < n_total and rows_info[i + 1]["level"] == "item":
                continue
            leaf_rows.append(ri)

    header = {"satker_kode": satker_kode, "satker_nama": satker_nama}
    return leaf_rows, header


# ----------------------------------------------------------------------
# Definisi kolom output final (nama kolom pendek)
# ----------------------------------------------------------------------

OUTPUT_COLUMNS = [
    ("kd_satker", lambda ri, h: h["satker_kode"]),
    ("satker", lambda ri, h: h["satker_nama"]),
    ("kd_program", lambda ri, h: ri["ctx"]["program_kode"]),
    ("program", lambda ri, h: ri["ctx"]["program_nama"]),
    ("kd_keg", lambda ri, h: ri["ctx"]["kegiatan_kode"]),
    ("kegiatan", lambda ri, h: ri["ctx"]["kegiatan_nama"]),
    ("kd_kro", lambda ri, h: ri["ctx"]["output_kode"]),
    ("kro", lambda ri, h: ri["ctx"]["output_nama"]),
    ("kd_ro", lambda ri, h: ri["ctx"]["suboutput_kode"]),
    ("ro", lambda ri, h: ri["ctx"]["suboutput_nama"]),
    ("kd_komponen", lambda ri, h: ri["ctx"]["komponen_kode"]),
    ("komponen", lambda ri, h: ri["ctx"]["komponen_nama"]),
    ("kd_subkomponen", lambda ri, h: ri["ctx"]["subkomponen_kode"]),
    ("subkomponen", lambda ri, h: ri["ctx"]["subkomponen_nama"]),
    ("kd_akun", lambda ri, h: ri["ctx"]["akun_kode"]),
    ("akun", lambda ri, h: ri["ctx"]["akun_nama"]),
    ("kd_item", lambda ri, h: ri["item_kode"]),
    ("item", lambda ri, h: ri["item_nama"]),
    ("pagu_anggaran", lambda ri, h: ri["pagu"]),
    ("realisasi_anggaran", lambda ri, h: ri["real_sd"]),
]

MONEY_COLS = {"pagu_anggaran", "realisasi_anggaran"}

COLUMN_WIDTHS = {
    "kd_satker": 10, "satker": 30, "kd_program": 10, "program": 32,
    "kd_keg": 10, "kegiatan": 32, "kd_kro": 10, "kro": 30,
    "kd_ro": 10, "ro": 36, "kd_komponen": 10, "komponen": 32,
    "kd_subkomponen": 12, "subkomponen": 42, "kd_akun": 10, "akun": 28,
    "kd_item": 10, "item": 45, "pagu_anggaran": 16, "realisasi_anggaran": 18,
}


def leaf_rows_to_dataframe(leaf_rows, header):
    """Ubah leaf_rows jadi pandas DataFrame flat dengan nama kolom final
    (dipakai untuk export CSV & visualisasi)."""
    records = []
    for ri in leaf_rows:
        row = {name: fn(ri, header) for name, fn in OUTPUT_COLUMNS}
        row["pagu_anggaran"] = row["pagu_anggaran"] or 0
        row["realisasi_anggaran"] = row["realisasi_anggaran"] or 0
        records.append(row)
    return pd.DataFrame(records, columns=[name for name, _ in OUTPUT_COLUMNS])


def build_output_bytes(leaf_rows, header):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Flat"

    font_name = "Arial"
    header_font = Font(name=font_name, bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    cell_font = Font(name=font_name, size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for j, (name, _fn) in enumerate(OUTPUT_COLUMNS, start=1):
        cell = ws.cell(row=1, column=j, value=name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    money_col_idx = [j for j, (name, _fn) in enumerate(OUTPUT_COLUMNS, start=1) if name in MONEY_COLS]

    for i, ri in enumerate(leaf_rows, start=2):
        for j, (_name, fn) in enumerate(OUTPUT_COLUMNS, start=1):
            val = fn(ri, header)
            cell = ws.cell(row=i, column=j, value=val)
            cell.font = cell_font
            cell.border = border
            if j in money_col_idx:
                cell.number_format = "#,##0"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_COLUMNS))}{len(leaf_rows) + 1}"

    for j, (name, _fn) in enumerate(OUTPUT_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(j)].width = COLUMN_WIDTHS.get(name, 14)

    ws.sheet_view.showGridLines = False

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ----------------------------------------------------------------------
# Helper visualisasi
# ----------------------------------------------------------------------

def grouped_bar(df, group_col, top_n=None):
    """Bar chart pagu_anggaran vs realisasi_anggaran, dikelompokkan per
    group_col, diurutkan berdasarkan realisasi terbesar. top_n membatasi
    jumlah kategori yang ditampilkan (sisanya digabung 'Lainnya')."""
    agg = (
        df.groupby(group_col, dropna=False)[["pagu_anggaran", "realisasi_anggaran"]]
        .sum()
        .reset_index()
    )
    agg[group_col] = agg[group_col].fillna("(Tidak diketahui)")
    agg = agg.sort_values("realisasi_anggaran", ascending=False)

    if top_n and len(agg) > top_n:
        head = agg.iloc[:top_n]
        rest = agg.iloc[top_n:][["pagu_anggaran", "realisasi_anggaran"]].sum()
        rest_row = pd.DataFrame([{
            group_col: f"Lainnya ({len(agg) - top_n} item)",
            "pagu_anggaran": rest["pagu_anggaran"],
            "realisasi_anggaran": rest["realisasi_anggaran"],
        }])
        agg = pd.concat([head, rest_row], ignore_index=True)

    agg["persentase_realisasi"] = (
        agg["realisasi_anggaran"] / agg["pagu_anggaran"].replace(0, pd.NA) * 100
    ).fillna(0)

    melted = agg.melt(
        id_vars=[group_col, "persentase_realisasi"],
        value_vars=["pagu_anggaran", "realisasi_anggaran"],
        var_name="Jenis", value_name="Nilai",
    )
    melted["Jenis"] = melted["Jenis"].map({
        "pagu_anggaran": "Pagu Anggaran", "realisasi_anggaran": "Realisasi Anggaran",
    })

    fig = px.bar(
        melted, x="Nilai", y=group_col, color="Jenis", orientation="h",
        barmode="group", text_auto=".2s",
        color_discrete_map={"Pagu Anggaran": "#94A3B8", "Realisasi Anggaran": "#1F4E78"},
    )
    fig.update_layout(
        yaxis={"categoryorder": "total ascending"},
        legend_title_text="", xaxis_title="Nilai (Rp)", yaxis_title="",
        height=max(320, 40 * len(agg)),
        margin=dict(l=10, r=10, t=30, b=10),
    )
    return fig, agg


# ----------------------------------------------------------------------
# UI Streamlit
# ----------------------------------------------------------------------

st.set_page_config(page_title="Flatten & Visualisasi Laporan 16 Segmen", page_icon="📊", layout="wide")

st.title("📊 Flatten & Visualisasi Laporan Fa Detail (16 Segmen)")
st.write(
    "Upload file laporan ketersediaan dana yang masih berbentuk hierarki "
    "(Program > Kegiatan > KRO > RO/Output > Komponen > SubKomponen > Akun > Item). "
    "Aplikasi akan menghasilkan tabel flat siap unduh (Excel & CSV), sekaligus visualisasi realisasi anggarannya."
)

uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        with st.spinner("Memproses file..."):
            file_bytes = io.BytesIO(uploaded_file.getvalue())
            leaf_rows, header = parse_workbook(file_bytes)
            output_buffer = build_output_bytes(leaf_rows, header)
            df = leaf_rows_to_dataframe(leaf_rows, header)

        st.success(f"Berhasil! {len(leaf_rows)} baris data flat dihasilkan.")

        total_pagu = df["pagu_anggaran"].sum()
        total_real = df["realisasi_anggaran"].sum()
        total_sisa = total_pagu - total_real
        pct_overall = (total_real / total_pagu * 100) if total_pagu else 0

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Pagu Anggaran", f"Rp {total_pagu:,.0f}")
        col2.metric("Total Realisasi Anggaran", f"Rp {total_real:,.0f}")
        col3.metric("Sisa Anggaran", f"Rp {total_sisa:,.0f}")
        col4.metric("% Realisasi", f"{pct_overall:.2f}%")
        st.progress(min(pct_overall / 100, 1.0))
        st.caption(
            "Cocokkan Total Pagu & Realisasi di atas dengan baris 'JUMLAH SELURUHNYA' pada "
            "laporan asli untuk memastikan tidak ada data yang hilang atau terhitung dobel."
        )

        base_name = uploaded_file.name.rsplit(".", 1)[0]
        csv_bytes = df.to_csv(index=False, sep=";").encode("utf-8-sig")

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="⬇️ Unduh file hasil (Excel .xlsx)",
                data=output_buffer,
                file_name=f"{base_name}_FLAT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col_dl2:
            st.download_button(
                label="⬇️ Unduh file hasil (CSV .csv)",
                data=csv_bytes,
                file_name=f"{base_name}_FLAT.csv",
                mime="text/csv",
                use_container_width=True,
            )
        st.caption(
            "File CSV menggunakan pemisah titik-koma (;) agar kolom tidak pecah saat dibuka "
            "langsung di Excel versi Indonesia (yang memakai koma sebagai pemisah desimal)."
        )

        st.divider()
        st.subheader("📈 Visualisasi Realisasi Anggaran")

        tab_keg, tab_ro, tab_komp, tab_akun = st.tabs(
            ["Per Kegiatan", "Per Output (RO)", "Per Komponen", "Per Akun"]
        )

        with tab_keg:
            fig, agg = grouped_bar(df, "kegiatan")
            st.plotly_chart(fig, use_container_width=True)
            with st.expander("Lihat tabel"):
                st.dataframe(agg, use_container_width=True, hide_index=True)

        with tab_ro:
            fig, agg = grouped_bar(df, "ro")
            st.plotly_chart(fig, use_container_width=True)
            with st.expander("Lihat tabel"):
                st.dataframe(agg, use_container_width=True, hide_index=True)

        with tab_komp:
            fig, agg = grouped_bar(df, "komponen", top_n=15)
            st.plotly_chart(fig, use_container_width=True)
            with st.expander("Lihat tabel"):
                st.dataframe(agg, use_container_width=True, hide_index=True)

        with tab_akun:
            fig, agg = grouped_bar(df, "akun", top_n=15)
            st.plotly_chart(fig, use_container_width=True)
            with st.expander("Lihat tabel"):
                st.dataframe(agg, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("🔍 Pratinjau Data Flat")
        st.dataframe(df.head(50), use_container_width=True, hide_index=True)

    except Exception as exc:
        st.error(
            "Terjadi kesalahan saat memproses file. Pastikan format file sesuai dengan "
            "'Laporan Fa Detail (16 Segmen)' dari SAKTI/SPAN."
        )
        st.exception(exc)
else:
    st.info("Silakan upload file Excel untuk mulai memproses.")
