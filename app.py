import streamlit as st
import pandas as pd
import qrcode
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
import io
import os
import zipfile

# ====================== KONFIGURASI HALAMAN ======================
st.set_page_config(
    page_title="Tamima E-Voucher Generator",
    page_icon="🎫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ====================== STYLE CUSTOM ======================
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 1rem;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

# ====================== FUNGSI UTILITAS ======================

def load_logo(logo_option, uploaded_logo_file):
    """
    Memuat logo berdasarkan pilihan user (Default atau Upload Custom).
    Returns PIL Image object or None.
    """
    logo_image = None
    if logo_option == "Upload Custom":
        if uploaded_logo_file:
            try:
                logo_image = Image.open(uploaded_logo_file)
            except Exception:
                st.sidebar.error("Gagal memuat gambar logo. Pastikan format PNG/JPG.")
        else:
            st.sidebar.warning("Silakan upload file logo.")
    else:
        # Default Logo
        if os.path.exists("logo_tamima.png"):
            try:
                logo_image = Image.open("logo_tamima.png")
            except Exception:
                st.sidebar.warning("File logo_tamima.png ditemukan tapi rusak.")
        else:
            st.sidebar.warning("File 'logo_tamima.png' tidak ditemukan di folder proyek.")
    return logo_image

def generate_qr_bytes(data):
    """
    Generate QR Code dan return sebagai BytesIO (in-memory).
    """
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=2
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    
    buffer = io.BytesIO()
    img.save(buffer, format='PNG')
    buffer.seek(0)
    return buffer

def create_pdf_voucher(row, logo_image):
    """
    Membuat PDF voucher untuk satu baris data menggunakan ReportLab.
    """
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    # --- HEADER ---
    # Logo
    if logo_image:
        logo_reader = ImageReader(logo_image)
        # Posisi: Kiri Atas (x=50, y=height-100), Ukuran: 100x80
        c.drawImage(logo_reader, x=50, y=height - 100, width=100, height=80, preserveAspectRatio=True)
    
    # Judul
    c.setFont("Helvetica-Bold", 20)
    c.setFillColor(colors.HexColor("#1E3A8A"))  # Biru Tua
    c.drawString(160, height - 80, "E-VOUCHER TAMIMA")
    
    c.setFont("Helvetica", 12)
    c.setFillColor(colors.black)
    c.drawString(160, height - 100, "Hajj & Umrah Service")
    
    # Garis Pemisah Emas
    c.setStrokeColor(colors.HexColor("#D4AF37"))
    c.setLineWidth(2)
    c.line(50, height - 115, width - 50, height - 115)
    
    # --- BODY (DETAIL DATA) ---
    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(colors.black)
    start_y = height - 150
    line_height = 25
    
    # Mapping label dan data
    fields = [
        ("Order ID", str(row['Order_ID'])),
        ("Nama Tamu", str(row['Nama_Tamu'])),
        ("Properti", str(row['Properti'])),
        ("Layanan", str(row['Layanan'])),
        ("Check-in", str(row['Check_in'])),
        ("Check-out", str(row['Check_out'])),
        ("Status", str(row['Status']))
    ]
    
    current_y = start_y
    for label, value in fields:
        c.drawString(60, current_y, f"{label}:")
        c.setFont("Helvetica", 11)
        c.drawString(150, current_y, value)
        c.setFont("Helvetica-Bold", 11)
        current_y -= line_height
    
    # --- QR CODE ---
    # Format: OrderID:{Order_ID}|Guest:{Nama_Tamu}|Status:{Status}
    qr_content = f"OrderID:{row['Order_ID']}|Guest:{row['Nama_Tamu']}|Status:{row['Status']}"
    qr_buffer = generate_qr_bytes(qr_content)
    qr_reader = ImageReader(qr_buffer)
    
    # Posisi QR: Kanan Tengah
    qr_size = 120
    qr_x = width - 200
    qr_y = height - 250
    c.drawImage(qr_reader, x=qr_x, y=qr_y, width=qr_size, height=qr_size)
    
    c.setFont("Helvetica-Oblique", 9)
    c.drawCentredString(qr_x + (qr_size/2), qr_y - 15, "Scan untuk verifikasi")
    
    # --- FOOTER ---
    c.setStrokeColor(colors.lightgrey)
    c.setLineWidth(1)
    c.line(50, 80, width - 50, 80)
    
    c.setFont("Helvetica-Oblique", 9)
    c.setFillColor(colors.darkgrey)
    c.drawString(50, 60, "Syarat & Ketentuan:")
    c.drawString(50, 45, "1. Voucher berlaku 6 bulan sejak tanggal pembelian.")
    c.drawString(50, 30, "2. Tidak dapat digabung dengan promo lainnya.")
    c.drawString(50, 15, "3. Tunjukkan voucher ini saat check-in.")
    
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# ====================== APLIKASI UTAMA ======================

def main():
    # Judul Utama
    st.markdown('<p class="main-header">🎫 Tamima E-Voucher Generator</p>', unsafe_allow_html=True)
    st.write("Unggah file Excel untuk membuat voucher PDF otomatis dengan Logo & QR Code.")
    
    # --- SIDEBAR PENGATURAN ---
    st.sidebar.header("⚙️ Pengaturan")
    logo_option = st.sidebar.radio(
        "Pilih Sumber Logo:",
        ["Default", "Upload Custom"],
        index=0
    )
    
    uploaded_logo_file = None
    if logo_option == "Upload Custom":
        uploaded_logo_file = st.sidebar.file_uploader(
            "Upload Logo (PNG/JPG)", 
            type=["png", "jpg", "jpeg"],
            help="Ukuran disarankan: 500x500 px, background transparan."
        )
    
    logo_image = load_logo(logo_option, uploaded_logo_file)
    if logo_image:
        st.sidebar.success("Logo berhasil dimuat.")
    else:
        st.sidebar.info("Voucher akan dibuat tanpa logo.")
    
    st.sidebar.markdown("---")
    st.sidebar.write("© 2024 Tamima Hajj & Umrah")
    
    # --- MAIN CONTENT ---
    uploaded_file = st.file_uploader("📤 Upload File Excel (.xlsx)", type=["xlsx"])
    
    if uploaded_file is not None:
        try:
            # Baca Excel
            df = pd.read_excel(uploaded_file)
            
            # Validasi Kolom
            required_cols = ['Order_ID', 'Nama_Tamu', 'Properti', 'Layanan', 'Check_in', 'Check_out', 'Status']
            missing_cols = [col for col in required_cols if col not in df.columns]
            
            if missing_cols:
                st.error(f"❌ **Kolom Tidak Ditemukan:** {', '.join(missing_cols)}")
                st.info(f"Pastikan file Excel memiliki kolom berikut: {', '.join(required_cols)}")
                st.stop()
            
            # Preview Data
            st.success(f"✅ File berhasil dibaca! Total **{len(df)}** data voucher.")
            with st.expander("👀 Preview Data Excel"):
                st.dataframe(df, use_container_width=True)
            
            st.markdown("---")
            st.header("📥 Opsi Download")
            
            # Pilihan Mode Download
            download_mode = st.radio(
                "Pilih metode download:",
                ["Download Single Voucher", "Download Semua (ZIP)"],
                horizontal=True
            )
            
            if download_mode == "Download Single Voucher":
                # Pilih Order ID
                order_list = df['Order_ID'].astype(str).tolist()
                selected_order = st.selectbox(
                    "Pilih Order ID:",
                    options=order_list,
                    index=0
                )
                
                # Ambil data baris terpilih
                selected_row = df[df['Order_ID'].astype(str) == selected_order].iloc[0]
                
                # Generate PDF Preview (Opsional, bisa langsung download)
                if st.button("📄 Generate & Download Voucher Ini", type="primary", key="btn_single"):
                    with st.spinner("Sedang membuat PDF..."):
                        try:
                            pdf_buffer = create_pdf_voucher(selected_row, logo_image)
                            file_name = f"Voucher_{selected_row['Order_ID']}.pdf"
                            
                            st.download_button(
                                label="⬇️ Klik untuk Download PDF",
                                data=pdf_buffer,
                                file_name=file_name,
                                mime="application/pdf",
                                key="download_single_pdf"
                            )
                            st.success("PDF siap diunduh!")
                        except Exception as e:
                            st.error(f"Gagal membuat PDF: {str(e)}")
            
            else:  # Download Semua (ZIP)
                st.warning(f"⚠️ Anda akan membuat **{len(df)}** file PDF dalam satu ZIP. Proses mungkin memakan waktu.")
                
                if st.button("🚀 Generate & Download ZIP", type="primary", key="btn_zip"):
                    zip_buffer = io.BytesIO()
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    try:
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for idx, row in df.iterrows():
                                # Update Progress
                                progress = (idx + 1) / len(df)
                                progress_bar.progress(progress)
                                status_text.text(f"Memproses: {row['Nama_Tamu']} ({idx + 1}/{len(df)})")
                                
                                # Generate PDF
                                pdf_buffer = create_pdf_voucher(row, logo_image)
                                
                                # Nama file dalam ZIP
                                # Sanitasi nama file untuk keamanan
                                safe_name = str(row['Nama_Tamu']).replace('/', '_').replace('\\', '_')
                                zip_filename = f"Voucher_{row['Order_ID']}_{safe_name}.pdf"
                                
                                # Masukkan ke ZIP
                                zip_file.writestr(zip_filename, pdf_buffer.getvalue())
                        
                        zip_buffer.seek(0)
                        status_text.success("✅ Semua voucher berhasil dibuat!")
                        progress_bar.empty()
                        
                        st.download_button(
                            label="⬇️ Download File ZIP",
                            data=zip_buffer,
                            file_name="Tamima_Vouchers.zip",
                            mime="application/zip",
                            key="download_zip_file"
                        )
                        
                    except Exception as e:
                        st.error(f"Terjadi kesalahan saat membuat ZIP: {str(e)}")
                        progress_bar.empty()
        
        except Exception as e:
            st.error(f"❌ Gagal membaca file Excel: {str(e)}")
            st.info("Pastikan file bukan dalam format .csv dan tidak sedang dibuka di aplikasi lain.")

if __name__ == "__main__":
    main()