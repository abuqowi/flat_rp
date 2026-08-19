"""
Streamlit App - Flatten & Visualisasi Laporan Fa Detail (16 Segmen)
----------------------------------------------------------------------
Upload file laporan hierarkis (.xlsx) -> dapatkan file flat (.xlsx) + visualisasi
realisasi anggaran (total, per Kegiatan, per Output, per Komponen).

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
# Logika parsing & flattening (sama seperti versi sebelumnya)
# ----------------------------------------------------------------------


def g(ws, r, c):
    return ws.cell(row=r, column=c).value


def parse_workbook(file_obj):
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb.active

    kementerian_kode = g(ws, 4, 15)
    kementerian_nama = g(ws, 4, 16)
    unit_kode = g(ws, 5, 15)
    unit_nama = g(ws, 5, 16)
    satker_kode = g(ws, 6, 15)
    satker_nama = g(ws, 6, 16)
    periode = g(ws, 3, 1)

    ctx = {
        "program_kode": None, "program_nama": None,
        "kegiatan_kode": None, "kegiatan_nama": None,
        "output_kode": None, "output_nama": None,
        "suboutput_kode": None, "suboutput_nama": None,
        "komponen_kode": None, "komponen_nama": None,
        "subkomponen_kode": None, "subkomponen_nama": None,
        "akun_kode": None, "akun_nama": None,
    }

    rows_info = []

    for r in range(10, ws.max_row + 1):
        b = g(ws, r, 2)   # Program / Kegiatan
        c = g(ws, r, 3)   # Output(KRO) / SubOutput(Output)
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
        lock = g(ws, r, 19)
        real_lalu = g(ws, r, 23)
        real_ini = g(ws, r, 24)
        real_sd = g(ws, r, 26)
        pct = g(ws, r, 29)
        sisa = g(ws, r, 31)

        rows_info.append({
            "row": r,
            "level": level,
            "ctx": dict(ctx),
            "item_kode": item_kode if level == "item" else None,
            "item_nama": item_nama if level == "item" else None,
            "pagu": pagu, "lock": lock, "real_lalu": real_lalu,
            "real_ini": real_ini, "real_sd": real_sd, "pct": pct, "sisa": sisa,
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

    header = {
        "kementerian_kode": kementerian_kode, "kementerian_nama": kementerian_nama,
        "unit_kode": unit_kode, "unit_nama": unit_nama,
        "satker_kode": satker_kode, "satker_nama": satker_nama,
        "periode": periode,
    }
    return leaf_rows, header


def leaf_rows_to_dataframe(leaf_rows, header):
    """Ubah leaf_rows jadi pandas DataFrame flat (dipakai untuk export & visualisasi)."""
    records = []
    for ri in leaf_rows:
        c = ri["ctx"]
        records.append({
            "Kementerian Kode": header["kementerian_kode"],
            "Kementerian Nama": header["kementerian_nama"],
            "Unit Organisasi Kode": header["unit_kode"],
            "Unit Organisasi Nama": header["unit_nama"],
            "Satker Kode": header["satker_kode"],
            "Satker Nama": header["satker_nama"],
            "Periode": header["periode"],
            "Program Kode": c["program_kode"],
            "Program Nama": c["program_nama"],
            "Kegiatan Kode": c["kegiatan_kode"],
            "Kegiatan Nama": c["kegiatan_nama"],
            "KRO Kode": c["output_kode"],
            "KRO Nama": c["output_nama"],
            "Output Kode": c["suboutput_kode"],
            "Output Nama": c["suboutput_nama"],
            "Komponen Kode": c["komponen_kode"],
            "Komponen Nama": c["komponen_nama"],
            "SubKomponen Kode": c["subkomponen_kode"],
            "SubKomponen Nama": c["subkomponen_nama"],
            "Akun Kode": c["akun_kode"],
            "Akun Nama": c["akun_nama"],
            "Item Kode": ri["item_kode"],
            "Item Nama": ri["item_nama"],
            "Pagu Revisi": ri["pagu"] or 0,
            "Lock Pagu": ri["lock"] or 0,
            "Realisasi Periode Lalu": ri["real_lalu"] or 0,
            "Realisasi Periode Ini": ri["real_ini"] or 0,
            "Realisasi s.d. Periode": ri["real_sd"] or 0,
            "Persentase": ri["pct"] or 0,
            "Sisa Anggaran": ri["sisa"] or 0,
        })
    return pd.DataFrame(records)


def build_output_bytes(leaf_rows, header):
    columns = [
        ("Kementerian Kode", lambda ri: header["kementerian_kode"]),
        ("Kementerian Nama", lambda ri: header["kementerian_nama"]),
        ("Unit Organisasi Kode", lambda ri: header["unit_kode"]),
        ("Unit Organisasi Nama", lambda ri: header["unit_nama"]),
        ("Satker Kode", lambda ri: header["satker_kode"]),
        ("Satker Nama", lambda ri: header["satker_nama"]),
        ("Periode", lambda ri: header["periode"]),
        ("Program Kode", lambda ri: ri["ctx"]["program_kode"]),
        ("Program Nama", lambda ri: ri["ctx"]["program_nama"]),
        ("Kegiatan Kode", lambda ri: ri["ctx"]["kegiatan_kode"]),
        ("Kegiatan Nama", lambda ri: ri["ctx"]["kegiatan_nama"]),
        ("KRO Kode", lambda ri: ri["ctx"]["output_kode"]),
        ("KRO Nama", lambda ri: ri["ctx"]["output_nama"]),
        ("Output Kode", lambda ri: ri["ctx"]["suboutput_kode"]),
        ("Output Nama", lambda ri: ri["ctx"]["suboutput_nama"]),
        ("Komponen Kode", lambda ri: ri["ctx"]["komponen_kode"]),
        ("Komponen Nama", lambda ri: ri["ctx"]["komponen_nama"]),
        ("SubKomponen Kode", lambda ri: ri["ctx"]["subkomponen_kode"]),
        ("SubKomponen Nama", lambda ri: ri["ctx"]["subkomponen_nama"]),
        ("Akun Kode", lambda ri: ri["ctx"]["akun_kode"]),
        ("Akun Nama", lambda ri: ri["ctx"]["akun_nama"]),
        ("Item Kode", lambda ri: ri["item_kode"]),
        ("Item Nama", lambda ri: ri["item_nama"]),
        ("Pagu Revisi", lambda ri: ri["pagu"]),
        ("Lock Pagu", lambda ri: ri["lock"]),
        ("Realisasi Periode Lalu", lambda ri: ri["real_lalu"]),
        ("Realisasi Periode Ini", lambda ri: ri["real_ini"]),
        ("Realisasi s.d. Periode", lambda ri: ri["real_sd"]),
        ("Persentase", lambda ri: ri["pct"]),
        ("Sisa Anggaran", lambda ri: ri["sisa"]),
    ]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Flat"

    font_name = "Arial"
    header_font = Font(name=font_name, bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    cell_font = Font(name=font_name, size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for j, (name, _fn) in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=j, value=name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    name_to_idx = {name: j for j, (name, _fn) in enumerate(columns, start=1)}
    money_cols = [
        name_to_idx[n] for n in [
            "Pagu Revisi", "Lock Pagu", "Realisasi Periode Lalu",
            "Realisasi Periode Ini", "Realisasi s.d. Periode", "Sisa Anggaran",
        ]
    ]
    pct_col = name_to_idx["Persentase"]

    for i, ri in enumerate(leaf_rows, start=2):
        for j, (_name, fn) in enumerate(columns, start=1):
            val = fn(ri)
            cell = ws.cell(row=i, column=j, value=val)
            cell.font = cell_font
            cell.border = border
            if j in money_cols:
                cell.number_format = "#,##0"
            elif j == pct_col:
                cell.number_format = "0.00%"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{len(leaf_rows) + 1}"

    widths = {
        "Kementerian Kode": 12, "Kementerian Nama": 22, "Unit Organisasi Kode": 10,
        "Unit Organisasi Nama": 20, "Satker Kode": 10, "Satker Nama": 28, "Periode": 14,
        "Program Kode": 10, "Program Nama": 32, "Kegiatan Kode": 10, "Kegiatan Nama": 32,
        "KRO Kode": 10, "KRO Nama": 28, "Output Kode": 10, "Output Nama": 36,
        "Komponen Kode": 10, "Komponen Nama": 32, "SubKomponen Kode": 10,
        "SubKomponen Nama": 42, "Akun Kode": 10, "Akun Nama": 28, "Item Kode": 10,
        "Item Nama": 45, "Pagu Revisi": 16, "Lock Pagu": 12,
        "Realisasi Periode Lalu": 16, "Realisasi Periode Ini": 16,
        "Realisasi s.d. Periode": 16, "Persentase": 10, "Sisa Anggaran": 16,
    }
    for name, j in name_to_idx.items():
        ws.column_dimensions[get_column_letter(j)].width = widths.get(name, 14)

    ws.sheet_view.showGridLines = False

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ----------------------------------------------------------------------
# Helper visualisasi
# ----------------------------------------------------------------------

def grouped_bar(df, group_col, top_n=None):
    """Bar chart Pagu vs Realisasi, dikelompokkan per group_col, diurutkan
    berdasarkan Realisasi terbesar. top_n membatasi jumlah kategori yang
    ditampilkan (sisanya digabung sebagai 'Lainnya') agar chart tetap
    terbaca kalau kategorinya banyak (mis. per SubKomponen/Item)."""
    agg = (
        df.groupby(group_col, dropna=False)[["Pagu Revisi", "Realisasi s.d. Periode"]]
        .sum()
        .reset_index()
    )
    agg[group_col] = agg[group_col].fillna("(Tidak diketahui)")
    agg = agg.sort_values("Realisasi s.d. Periode", ascending=False)

    if top_n and len(agg) > top_n:
        head = agg.iloc[:top_n]
        rest = agg.iloc[top_n:][["Pagu Revisi", "Realisasi s.d. Periode"]].sum()
        rest_row = pd.DataFrame([{
            group_col: f"Lainnya ({len(agg) - top_n} item)",
            "Pagu Revisi": rest["Pagu Revisi"],
            "Realisasi s.d. Periode": rest["Realisasi s.d. Periode"],
        }])
        agg = pd.concat([head, rest_row], ignore_index=True)

    agg["Persentase Realisasi"] = (
        agg["Realisasi s.d. Periode"] / agg["Pagu Revisi"].replace(0, pd.NA) * 100
    ).fillna(0)

    melted = agg.melt(
        id_vars=[group_col, "Persentase Realisasi"],
        value_vars=["Pagu Revisi", "Realisasi s.d. Periode"],
        var_name="Jenis", value_name="Nilai",
    )

    fig = px.bar(
        melted, x="Nilai", y=group_col, color="Jenis", orientation="h",
        barmode="group", text_auto=".2s",
        color_discrete_map={"Pagu Revisi": "#94A3B8", "Realisasi s.d. Periode": "#1F4E78"},
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
    "(Program > Kegiatan > KRO > Output > Komponen > SubKomponen > Akun > Item). "
    "Aplikasi akan menghasilkan tabel flat siap unduh, sekaligus visualisasi realisasi anggarannya."
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

        total_pagu = df["Pagu Revisi"].sum()
        total_real = df["Realisasi s.d. Periode"].sum()
        total_sisa = df["Sisa Anggaran"].sum()
        pct_overall = (total_real / total_pagu * 100) if total_pagu else 0

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Pagu Revisi", f"Rp {total_pagu:,.0f}")
        col2.metric("Total Realisasi s.d. Periode", f"Rp {total_real:,.0f}")
        col3.metric("Sisa Anggaran", f"Rp {total_sisa:,.0f}")
        col4.metric("% Realisasi", f"{pct_overall:.2f}%")
        st.progress(min(pct_overall / 100, 1.0))
        st.caption(
            "Cocokkan angka di atas dengan baris 'JUMLAH SELURUHNYA' pada laporan asli "
            "untuk memastikan tidak ada data yang hilang atau terhitung dobel."
        )

        output_name = uploaded_file.name.rsplit(".", 1)[0] + "_FLAT.xlsx"
        csv_name = uploaded_file.name.rsplit(".", 1)[0] + "_FLAT.csv"
        # utf-8-sig supaya karakter dan koma ribuan terbaca benar saat dibuka di Excel
        csv_bytes = df.to_csv(index=False, sep=";").encode("utf-8-sig")

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="⬇️ Unduh file hasil (Excel .xlsx)",
                data=output_buffer,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col_dl2:
            st.download_button(
                label="⬇️ Unduh file hasil (CSV .csv)",
                data=csv_bytes,
                file_name=csv_name,
                mime="text/csv",
                use_container_width=True,
            )
        st.caption(
            "File CSV menggunakan pemisah titik-koma (;) agar kolom tidak pecah saat dibuka "
            "langsung di Excel versi Indonesia (yang memakai koma sebagai pemisah desimal)."
        )

        st.divider()
        st.subheader("📈 Visualisasi Realisasi Anggaran")

        tab_keg, tab_out, tab_komp, tab_akun = st.tabs(
            ["Per Kegiatan", "Per Output", "Per Komponen", "Per Akun"]
        )

        with tab_keg:
            fig, agg = grouped_bar(df, "Kegiatan Nama")
            st.plotly_chart(fig, use_container_width=True)
            with st.expander("Lihat tabel"):
                st.dataframe(agg, use_container_width=True, hide_index=True)

        with tab_out:
            fig, agg = grouped_bar(df, "Output Nama")
            st.plotly_chart(fig, use_container_width=True)
            with st.expander("Lihat tabel"):
                st.dataframe(agg, use_container_width=True, hide_index=True)

        with tab_komp:
            fig, agg = grouped_bar(df, "Komponen Nama", top_n=15)
            st.plotly_chart(fig, use_container_width=True)
            with st.expander("Lihat tabel"):
                st.dataframe(agg, use_container_width=True, hide_index=True)

        with tab_akun:
            fig, agg = grouped_bar(df, "Akun Nama", top_n=15)
            st.plotly_chart(fig, use_container_width=True)
            with st.expander("Lihat tabel"):
                st.dataframe(agg, use_container_width=True, hide_index=True)

    except Exception as exc:
        st.error(
            "Terjadi kesalahan saat memproses file. Pastikan format file sesuai dengan "
            "'Laporan Fa Detail (16 Segmen)' dari SAKTI/SPAN."
        )
        st.exception(exc)
else:
    st.info("Silakan upload file Excel untuk mulai memproses.")
