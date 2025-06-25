

import streamlit as st
import win32com.client
import pythoncom
import pyperclip
import datetime
import time # <-- SATU-SATUNYA PENAMBAHAN IMPORT

def run_full_sap_automation(start_date, end_date, status_code, shipping_point_list_str):
    """
    Pastikan Anda sudah LOGIN MySAP dengan posisi standby tanpa T-code.
    """
    pythoncom.CoInitialize()

    try:
        # Menyalin daftar Shipping Point ke clipboard
        if not shipping_point_list_str:
            return "Error: Daftar Shipping Point tidak boleh kosong."
        pyperclip.copy(shipping_point_list_str)

        # Menghubungkan ke SAP GUI
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        # --- Mulai Rangkaian Automasi ---

        # 1. Navigasi ke T-Code VL06F
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "vl06f"
        session.findById("wnd[0]").sendVKey(0) # Tekan Enter

        # 2. Mengisi Nilai-nilai yang Tetap (Hardcoded)
        session.findById("wnd[0]/usr/ctxtIT_VKORG-LOW").text = "1005"
        session.findById("wnd[0]/usr/ctxtIT_VKORG-HIGH").text = "2205"
        session.findById("wnd[0]/usr/ctxtIT_VTWEG-LOW").text = "10"
        session.findById("wnd[0]/usr/ctxtIT_VTWEG-HIGH").text = "20"
        session.findById("wnd[0]/usr/ctxtIT_SPART-LOW").text = "00"
        session.findById("wnd[0]/usr/ctxtIT_SPART-HIGH").text = "07"
        
        # 3. Mengisi Shipping Point (VSTEL) dari input pengguna via clipboard
        session.findById("wnd[0]/usr/btn%_IF_VSTEL_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press() # Upload from Clipboard
        session.findById("wnd[1]/tbar[0]/btn[8]").press() # Execute

        # 4. Mengisi Nilai dari Input Pengguna (Tanggal & Status)
        formatted_start_date = start_date.strftime("%d.%m.%Y")
        formatted_end_date = end_date.strftime("%d.%m.%Y")
        
        session.findById("wnd[0]/usr/ctxtIT_WADAT-LOW").text = formatted_start_date
        session.findById("wnd[0]/usr/ctxtIT_WADAT-HIGH").text = formatted_end_date
        session.findById("wnd[0]/usr/ctxtIT_WBSTK-LOW").text = status_code
        session.findById("wnd[0]/usr/ctxtIT_WBSTK-HIGH").text = status_code

        # 5. Mencentang Checkbox
        session.findById("wnd[0]/usr/chkIF_ITEM").selected = True
        session.findById("wnd[0]/usr/chkIF_ANZPO").selected = True
        session.findById("wnd[0]/usr/chkIF_SPD_A").selected = True

        # 6. Eksekusi Laporan Awal
        session.findById("wnd[0]/tbar[1]/btn[8]").press() # Tekan tombol Execute
        
        # --- PERBAIKAN DI SINI ---
        # Beri SAP waktu untuk memuat layar hasil sebelum melanjutkan
        time.sleep(2) 
        # ------------------------

        # 7. Langkah-langkah Pemrosesan Hasil Laporan
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr").verticalScrollbar.position = 744
        session.findById("wnd[1]/usr").verticalScrollbar.position = 1072
        session.findById("wnd[1]/usr/lbl[1,14]").setFocus()
        session.findById("wnd[1]").sendVKey(2) # F2 (double-click/change)
        session.findById("wnd[0]").sendVKey(43) # Ctrl+F11
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        return "Sukses: Seluruh rangkaian automasi SAP berhasil dijalankan."

    except pythoncom.com_error as e:
        error_message = str(e)
        if "The control could not be found by id" in error_message:
            return f"Error: Elemen tidak ditemukan. {e}. Pastikan Anda berada di layar SAP yang benar atau periksa kembali ID elemen menggunakan Script Recorder."
        else:
            return f"Terjadi error COM: {e}. Pastikan SAP GUI berjalan dan tidak ada dialog tak terduga yang muncul."
    except Exception as e:
        return f"Terjadi error yang tidak terduga: {e}"
    finally:
        pythoncom.CoUninitialize()


# --- Antarmuka Streamlit ---

st.set_page_config(page_title="SAP DO Downloader", layout="centered")

st.title("DOWNLOAD DO OUTSTANDING (VL06F)")
st.info(" Pastikan Anda sudah LOGIN MySAP dengan posisi standby tanpa T-code..")

# 1. Input Status DO
st.subheader("1. Pilih Status DO (WBSTK)")
status_options = {
    "A": "A = DO GANTUNG",
    "B": "B = DO PARSIAL",
    "C": "C = DO SUDAH GI"
}
selected_status_label = st.selectbox(
    "Pilih status DO",
    options=list(status_options.values()),
    label_visibility="collapsed"
)
status_code = [key for key, value in status_options.items() if value == selected_status_label][0]

# 2. Input Tanggal
st.subheader("2. Pilih Rentang Tanggal (WADAT)")
today = datetime.date.today()

date_selection = st.date_input(
    "Pilih tanggal mulai dan tanggal selesai:",
    value=(today, today),
    format="DD.MM.YYYY"
)

start_date, end_date = (None, None)

if isinstance(date_selection, (list, tuple)) and len(date_selection) == 2:
    start_date, end_date = date_selection

# 3. Input Shipping Point
st.subheader("3. Input Data Shipping Point (VSTEL)")
input_method = st.radio(
    "Pilih metode input:",
    ("Input Manual", "Upload File Teks (.txt)")
)

shipping_point_data = ""
if input_method == "Input Manual":
    shipping_point_data = st.text_area(
        "Masukkan satu kode Shipping Point per baris", 
        height=150, 
        placeholder="Contoh:\n216F\n215R\n215C\n255Q\n255R\n255C\n255F\n255G\n255H\n255"
    )
else:
    uploaded_file = st.file_uploader("Pilih file .txt", type=["txt"])
    if uploaded_file is not None:
        shipping_point_data = uploaded_file.getvalue().decode("utf-8")
        st.text_area("Pratinjau isi file:", shipping_point_data, height=150, disabled=True)

# Tombol untuk menjalankan
if st.button("Jalankan Automasi Lengkap", type="primary"):
    
    # Validasi Input Pengguna
    if not all([start_date, end_date]):
        st.warning("Mohon lengkapi range tanggal (pilih tanggal mulai dan selesai).")
    elif not shipping_point_data.strip():
        st.warning("Mohon masukkan atau upload data Shipping Point terlebih dahulu.")
    elif start_date > end_date:
        st.error("Tanggal mulai tidak boleh lebih besar dari tanggal selesai.")
    else:
        with st.spinner("Sedang memproses... Menjalankan rangkaian automasi di SAP."):
            result = run_full_sap_automation(start_date, end_date, status_code, shipping_point_data)
            if result.startswith("Sukses"):
                st.success(result)
            else:
                st.error(result)

st.markdown("---")
st.write("aripili-2025.")
