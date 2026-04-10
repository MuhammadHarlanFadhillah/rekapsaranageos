import streamlit as st
import pandas as pd
from datetime import datetime
import io
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


CANONICAL_COLUMNS = [
    "NO",
    "NO_HULL",
    "TYPE_CAR",
    "DRIVER",
    "FROM",
    "DESTINATION",
    "ACTIVITIES",
    "Date_Departure",
    "Date_Arrival",
    "Distance_Start",
    "Distance_Finish",
    "Distance_Total",
    "BBM,_L",
    "Time_Departure",
    "Time_Arrival",
    "TOTAL_TIME",
]


def normalize_sheet_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize sheet data so app always works with the expected columns."""
    if df is None or df.empty:
        return pd.DataFrame(columns=CANONICAL_COLUMNS)

    work = df.copy()
    work = work.dropna(how="all")

    header_tokens = {
        "NO",
        "NO_HULL",
        "TYPE_CAR",
        "DRIVER",
        "FROM",
        "DESTINATION",
        "ACTIVITIES",
        "DEPARTURE",
        "ARRIVAL",
        "START",
        "FINISH",
        "TOTAL",
    }

    header_row_idx = None
    max_scan = min(len(work), 15)
    for idx in range(max_scan):
        values = {
            str(v).strip().upper()
            for v in work.iloc[idx].tolist()
            if pd.notna(v) and str(v).strip() and str(v).strip().lower() != "none"
        }
        if len(values & header_tokens) >= 6:
            header_row_idx = idx
            break

    if header_row_idx is not None:
        work = work.iloc[header_row_idx + 1 :].reset_index(drop=True)

    work = work.iloc[:, : len(CANONICAL_COLUMNS)].copy()
    while work.shape[1] < len(CANONICAL_COLUMNS):
        work[f"_pad_{work.shape[1]}"] = pd.NA

    work.columns = CANONICAL_COLUMNS
    work = work.dropna(how="all")

    # Kolom NO idealnya integer agar urutan jelas saat tambah data.
    work["NO"] = pd.to_numeric(work["NO"], errors="coerce").astype("Int64")
    return work


def stringify_for_editor(df: pd.DataFrame) -> pd.DataFrame:
    """Prepare dataframe so every editable field is safe for Streamlit editor."""
    editor_df = df.copy()
    for col in editor_df.columns:
        editor_df[col] = editor_df[col].where(editor_df[col].notna(), "")
    return editor_df


def filter_dataframe(df: pd.DataFrame, keyword: str) -> pd.DataFrame:
    """Simple global search across all columns."""
    if not keyword:
        return df
    mask = df.astype(str).apply(lambda s: s.str.contains(keyword, case=False, na=False)).any(axis=1)
    return df[mask]


def build_styled_excel(df: pd.DataFrame, sheet_name: str) -> bytes:
    """Create a styled Excel file similar to the spreadsheet layout."""
    export_df = df.copy()

    for col in ["Date_Departure", "Date_Arrival"]:
        parsed = pd.to_datetime(export_df[col], errors="coerce")
        export_df[col] = parsed.dt.strftime("%d/%m/%Y")
        export_df[col] = export_df[col].fillna("")

    for col in ["Time_Departure", "Time_Arrival", "TOTAL_TIME"]:
        export_df[col] = export_df[col].fillna("").astype(str)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=2, header=False)
        ws = writer.book[sheet_name]

        ws.merge_cells("H1:I1")
        ws.merge_cells("J1:L1")
        ws.merge_cells("N1:P1")

        ws["H1"] = "DATE"
        ws["J1"] = "Distance (KM)"
        ws["N1"] = "Time"

        header_labels = [
            "NO",
            "NO_HULL",
            "TYPE_CAR",
            "DRIVER",
            "FROM",
            "DESTINATION",
            "ACTIVITIES",
            "Departure",
            "Arrival",
            "Start",
            "Finish",
            "Total",
            "BBM,_L",
            "Departure",
            "Arrival",
            "TOTAL_TIME",
        ]
        for col_idx, label in enumerate(header_labels, start=1):
            ws.cell(row=2, column=col_idx, value=label)

        green_fill = PatternFill("solid", fgColor="D9EAD3")
        blue_fill = PatternFill("solid", fgColor="9FC5E8")
        yellow_fill = PatternFill("solid", fgColor="F1C232")

        header_font = Font(bold=True)
        center = Alignment(horizontal="center", vertical="center")
        left = Alignment(horizontal="left", vertical="center")
        thin = Side(style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        def fill_for_col(col_idx: int):
            if col_idx in (8, 9):
                return blue_fill
            if col_idx in (10, 11, 12, 13):
                return yellow_fill
            return green_fill

        ws.row_dimensions[1].height = 22
        ws.row_dimensions[2].height = 22

        for row_idx in (1, 2):
            for col_idx in range(1, 17):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.fill = fill_for_col(col_idx)
                cell.font = header_font
                cell.alignment = center
                cell.border = border

        widths = [7, 14, 12, 12, 12, 14, 28, 12, 12, 11, 11, 10, 9, 11, 11, 11]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = w

        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=16):
            for cell in row:
                cell.border = border
                if cell.column == 7:
                    cell.alignment = left
                else:
                    cell.alignment = center

        ws.freeze_panes = "A3"

    return buffer.getvalue()

try:
    from streamlit_gsheets import GSheetsConnection
except ImportError:
    GSheetsConnection = None

# --- UI STREAMLIT ---
st.set_page_config(page_title="Rekap Sarana Kendaraan", layout="wide")
st.title("🚗 Aplikasi Rekap Sarana Kendaraan")

# --- KONEKSI KE GOOGLE SHEETS ---
# Membuat objek koneksi
if GSheetsConnection is None:
    st.error("Dependency st-gsheets-connection belum terpasang. Jalankan install dependency dari requirements.txt.")
    st.stop()

try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as exc:
    st.error("Koneksi Google Sheets belum siap. Pastikan secret Streamlit sudah diisi dengan benar.")
    st.caption(str(exc))
    st.stop()

gsheets_spreadsheet = ""
try:
    gsheets_spreadsheet = st.secrets.get("connections", {}).get("gsheets", {}).get("spreadsheet", "")
except Exception:
    gsheets_spreadsheet = ""

# Fallback aman: spreadsheet bisa diisi dari sidebar jika belum ditetapkan di secrets.
target_spreadsheet = st.sidebar.text_input(
    "Spreadsheet (nama atau URL)",
    value=str(gsheets_spreadsheet).strip(),
    placeholder="Contoh: Rekap Sarana Kendaraan atau https://docs.google.com/spreadsheets/...",
)

if not str(target_spreadsheet).strip():
    st.error("Konfigurasi belum lengkap: isi nilai 'spreadsheet' di sidebar atau Streamlit secrets.")
    st.stop()

# Daftar bulan/sheet yang tersedia (harus sesuai dengan nama tab di Google Sheets)
sheet_names = [
    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
    "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
]

# Sidebar Menu
menu = st.sidebar.radio("Pilih Menu", ["Lihat Data", "Tambah Data", "Edit Data", "Hapus Data"])
selected_sheet = st.sidebar.selectbox("Pilih Bulan (Sheet)", sheet_names)

# Tarik data dari Google Sheets (cache diset rendah agar lumayan real-time)
df_raw = conn.read(spreadsheet=target_spreadsheet.strip(), worksheet=selected_sheet, ttl=5)
df_current = normalize_sheet_dataframe(df_raw)

if menu == "Lihat Data":
    display_df = df_current.rename(
        columns={
            "Date_Departure": "Departure Date",
            "Date_Arrival": "Arrival Date",
            "Distance_Start": "Start (KM)",
            "Distance_Finish": "Finish (KM)",
            "Distance_Total": "Total (KM)",
            "Time_Departure": "Departure Time",
            "Time_Arrival": "Arrival Time",
        }
    )

    search_keyword = st.text_input(
        "Cari data (NO, driver, hull, tujuan, aktivitas, dll)",
        placeholder="Ketik kata kunci pencarian...",
    ).strip()
    filtered_display_df = filter_dataframe(display_df, search_keyword)

    st.subheader(f"Data Rekap - {selected_sheet}")
    st.dataframe(filtered_display_df, use_container_width=True)

    excel_bytes = build_styled_excel(df_current, selected_sheet)
    st.download_button(
        label=f"📥 Download Data {selected_sheet} (Excel)",
        data=excel_bytes,
        file_name=f"Rekap_Kendaraan_{selected_sheet}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if menu == "Tambah Data":
    st.subheader(f"➕ Tambah Data Baru - {selected_sheet}")
    with st.form("form_tambah_data", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            no = st.number_input("NO", min_value=1, step=1)
            no_hull = st.text_input("NO_HULL")
            type_car = st.text_input("TYPE_CAR")
            driver = st.text_input("DRIVER")
            from_loc = st.text_input("FROM")
            destination = st.text_input("DESTINATION")
            activities = st.text_input("ACTIVITIES")
            
        with col2:
            st.markdown("**DATE & TIME**")
            date_dep = st.date_input("Date Departure")
            time_dep = st.time_input("Time Departure")
            date_arr = st.date_input("Date Arrival")
            time_arr = st.time_input("Time Arrival")
            
        with col3:
            st.markdown("**DISTANCE (KM)**")
            dist_start = st.number_input("Distance Start", min_value=0.0, step=0.1)
            dist_finish = st.number_input("Distance Finish", min_value=0.0, step=0.1)
            
        submit_btn = st.form_submit_button("Simpan Data")
        
        if submit_btn:
            # 1. Kalkulasi Jarak & BBM
            dist_total = dist_finish - dist_start
            bbm_l = dist_total / 10.0
            
            # 2. Kalkulasi Waktu
            datetime_dep = datetime.combine(date_dep, time_dep)
            datetime_arr = datetime.combine(date_arr, time_arr)
            time_diff = datetime_arr - datetime_dep
            
            total_seconds = int(time_diff.total_seconds())
            hours, remainder = divmod(total_seconds, 3600)
            minutes, _ = divmod(remainder, 60)
            time_total_str = f"{hours:02d}:{minutes:02d}"
            
            # 3. Validasi input
            if dist_finish < dist_start:
                st.error("Distance Finish tidak boleh lebih kecil dari Distance Start!")
            elif datetime_arr < datetime_dep:
                st.error("Waktu Arrival tidak boleh lebih awal dari Departure!")
            else:
                # 4. Siapkan data baru (Pastikan nama *key* sama persis dengan header di Google Sheets kamu)
                new_data = {
                    "NO": int(no),
                    "NO_HULL": no_hull,
                    "TYPE_CAR": type_car,
                    "DRIVER": driver,
                    "FROM": from_loc,
                    "DESTINATION": destination,
                    "ACTIVITIES": activities,
                    "Date_Departure": date_dep.strftime("%Y-%m-%d"),
                    "Date_Arrival": date_arr.strftime("%Y-%m-%d"),
                    "Distance_Start": dist_start,
                    "Distance_Finish": dist_finish,
                    "Distance_Total": dist_total,
                    "BBM,_L": bbm_l,
                    "Time_Departure": time_dep.strftime("%H:%M"),
                    "Time_Arrival": time_arr.strftime("%H:%M"),
                    "TOTAL_TIME": time_total_str
                }
                
                # Masukkan baris baru ke dataframe
                new_row_df = pd.DataFrame([new_data])
                updated_df = pd.concat([df_current, new_row_df], ignore_index=True)
                
                # Update ke Google Sheets
                conn.update(spreadsheet=target_spreadsheet.strip(), worksheet=selected_sheet, data=updated_df)
                
                st.success("✅ Data berhasil di-push ke Google Sheets!")
                
                # Bersihkan cache agar tabel langsung menampilkan data terbaru
                st.cache_data.clear()
                st.rerun()

if menu == "Edit Data":
    st.subheader(f"✏️ Edit Data - {selected_sheet}")
    st.caption("Edit langsung di tabel, lalu simpan perubahan.")

    editor_df = stringify_for_editor(df_current)
    edited_df = st.data_editor(
        editor_df,
        use_container_width=True,
        num_rows="fixed",
        key=f"editor_only_{selected_sheet}",
        column_config={"NO": st.column_config.NumberColumn("NO", step=1)},
    )

    if st.button("💾 Simpan Edit", type="primary"):
        rows_to_save = edited_df.copy()
        for col in CANONICAL_COLUMNS:
            if col not in rows_to_save.columns:
                rows_to_save[col] = ""
        rows_to_save = rows_to_save[CANONICAL_COLUMNS]

        rows_to_save["NO"] = pd.to_numeric(rows_to_save["NO"], errors="coerce").astype("Int64")
        for col in ["Distance_Start", "Distance_Finish", "Distance_Total", "BBM,_L"]:
            rows_to_save[col] = pd.to_numeric(rows_to_save[col], errors="coerce")

        conn.update(spreadsheet=target_spreadsheet.strip(), worksheet=selected_sheet, data=rows_to_save)
        st.success("✅ Perubahan edit berhasil disimpan ke Google Sheets!")
        st.cache_data.clear()
        st.rerun()

if menu == "Hapus Data":
    st.subheader(f"🗑️ Hapus Data - {selected_sheet}")
    st.caption("Centang baris yang ingin dihapus, lalu klik tombol hapus.")

    delete_df = stringify_for_editor(df_current)
    delete_df.insert(0, "HAPUS", False)

    delete_editor_df = st.data_editor(
        delete_df,
        use_container_width=True,
        num_rows="fixed",
        key=f"delete_only_{selected_sheet}",
        column_config={
            "HAPUS": st.column_config.CheckboxColumn("HAPUS", help="Centang untuk hapus baris", default=False)
        },
        disabled=CANONICAL_COLUMNS,
    )

    if st.button("🗑️ Hapus Baris Terpilih", type="primary"):
        rows_to_keep = delete_editor_df[delete_editor_df["HAPUS"] != True].copy()
        rows_to_keep = rows_to_keep.drop(columns=["HAPUS"])

        for col in CANONICAL_COLUMNS:
            if col not in rows_to_keep.columns:
                rows_to_keep[col] = ""
        rows_to_keep = rows_to_keep[CANONICAL_COLUMNS]

        rows_to_keep["NO"] = pd.to_numeric(rows_to_keep["NO"], errors="coerce").astype("Int64")
        for col in ["Distance_Start", "Distance_Finish", "Distance_Total", "BBM,_L"]:
            rows_to_keep[col] = pd.to_numeric(rows_to_keep[col], errors="coerce")

        conn.update(spreadsheet=target_spreadsheet.strip(), worksheet=selected_sheet, data=rows_to_keep)
        st.success("✅ Data terpilih berhasil dihapus dari Google Sheets!")
        st.cache_data.clear()
        st.rerun()