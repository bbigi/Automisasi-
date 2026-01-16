import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import argparse
import sys
from dotenv import load_dotenv

# Load configuration dari .env
load_dotenv()
SMTP_SERVER = os.getenv('SMTP_SERVER')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))
SENDER_EMAIL = os.getenv('SENDER_EMAIL')
SENDER_PASSWORD = os.getenv('SENDER_PASSWORD')

# ==========================================
# 1. MODUL OTOMASI PENGOLAHAN NILAI
# ==========================================
def proses_nilai(input_file):
    print(f"\n[1/2] Membaca data nilai dari: {input_file}...")
    
    try:
        # Deteksi format file (CSV atau Excel)
        if input_file.endswith('.csv'):
            df = pd.read_csv(input_file)
        else:
            df = pd.read_excel(input_file)

        # Cek kelengkapan kolom
        kolom_wajib = ['Nama', 'Tugas', 'UTS', 'UAS']
        if not all(col in df.columns for col in kolom_wajib):
            print(f"[ERROR] File wajib memiliki kolom: {kolom_wajib}")
            return

        # 1. Hitung Nilai Akhir (Sesuai Soal: Tugas 20%, UTS 40%, UAS 40%)
        # [cite: 29]
        df['Nilai_Akhir'] = (df['Tugas'] * 0.2) + (df['UTS'] * 0.4) + (df['UAS'] * 0.4)

        # 2. Tentukan Grade (Rule Buatan Sendiri) [cite: 30]
        def get_grade(score):
            if score > 80: return 'A'
            elif score > 70: return 'B'
            elif score > 55: return 'C'
            elif score > 40: return 'D'
            else: return 'E'

        df['Grade'] = df['Nilai_Akhir'].apply(get_grade)

        # 3. Tentukan Status Lulus/Gagal [cite: 27]
        # Asumsi: A, B, C Lulus. D, E Gagal.
        df['Status'] = df['Grade'].apply(lambda x: 'LULUS' if x in ['A', 'B', 'C'] else 'GAGAL')

        # 4. Generate Laporan [cite: 31]
        output_file = 'Hasil_Laporan_Nilai.xlsx'
        df.to_excel(output_file, index=False)
        abs_path = os.path.abspath(output_file)
        print(f"[SUKSES] Laporan nilai tersimpan di:")
        print(f"         {abs_path}")
        
    except FileNotFoundError:
        print("[ERROR] File tidak ditemukan. Pastikan nama file benar.")
    except Exception as e:
        print(f"[ERROR] Terjadi kesalahan: {e}")

# ==========================================
# 2. MODUL OTOMASI KEHADIRAN & EMAIL
# ==========================================
def kirim_email_real(penerima, nama, persentase):
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = penerima
        msg['Subject'] = 'PERINGATAN PRESENSI'
        
        body = f"Halo {nama},\n\nKehadiran Anda saat ini {persentase:.1f}% (kurang dari 75%).\nMohon segera perbaiki kehadiran Anda.\n\nTerima kasih."
        msg.attach(MIMEText(body, 'plain'))
        
        server.send_message(msg)
        server.quit()
        print(f"   >>> [EMAIL BERHASIL] Dikirim ke: {penerima}")
    except Exception as e:
        print(f"   >>> [EMAIL GAGAL] {penerima}: {str(e)}")

def proses_kehadiran(input_file):
    print(f"\n[2/2] Membaca data kehadiran dari: {input_file}...")
    
    try:
        if input_file.endswith('.xlsx'):
            df = pd.read_excel(input_file)
        else:
            df = pd.read_csv(input_file)    

        # Normalisasi nama kolom: hapus spasi, underscore, dan jadikan lowercase untuk perbandingan
        df_cols_normalized = {col.replace('_', '').replace(' ', '').lower(): col for col in df.columns}
        
        # Cari kolom yang dibutuhkan (fleksibel terhadap underscore/spasi)
        col_nama = next((col for norm, col in df_cols_normalized.items() if norm == 'nama'), None)
        col_email = next((col for norm, col in df_cols_normalized.items() if norm == 'email'), None)
        col_hadir = next((col for norm, col in df_cols_normalized.items() if norm == 'jumlahhadir'), None)
        col_total = next((col for norm, col in df_cols_normalized.items() if norm == 'totalpertemuan'), None)
        
        if not all([col_nama, col_email, col_hadir, col_total]):
            print(f"[ERROR] File wajib memiliki kolom: Nama, Email, JumlahHadir (atau Jumlah_Hadir), TotalPertemuan (atau Total_Pertemuan)")
            print(f"[INFO] Kolom yang ditemukan: {list(df.columns)}")
            return
        
        # Rename ke format standar untuk kemudahan
        df = df.rename(columns={col_nama: 'Nama', col_email: 'Email', col_hadir: 'JumlahHadir', col_total: 'TotalPertemuan'})

        # 1. Hitung Persentase [cite: 40]
        df['Persentase'] = (df['JumlahHadir'] / df['TotalPertemuan']) * 100

        # 2. Logika Peringatan (< 75%) [cite: 41-42]
        status_list = []
        for index, row in df.iterrows():
            if row['Persentase'] < 75:
                status_list.append("PERINGATAN")
                # Panggil fungsi kirim email
                kirim_email_real(row['Email'], row['Nama'], row['Persentase'])
            else:
                status_list.append("OK")

        df['Status_Kehadiran'] = status_list
        
        # 3. Generate Laporan
        output_file = 'Hasil_Laporan_Kehadiran.xlsx'
        df.to_excel(output_file, index=False)
        abs_path = os.path.abspath(output_file)
        print(f"[SUKSES] Laporan kehadiran tersimpan di:")
        print(f"         {abs_path}")

    except FileNotFoundError:
        print("[ERROR] File tidak ditemukan.")
    except Exception as e:
        print(f"[ERROR] Terjadi kesalahan: {e}")

# ==========================================
# MAIN PROGRAM (GANTI NAMA FILE DI SINI)
# ==========================================
def prompt_for_file(prompt_text):
    while True:
        p = input(prompt_text).strip('"')
        if p == '':
            return None
        if os.path.exists(p):
            return p
        print("File tidak ditemukan. Coba lagi atau tekan Enter untuk lewati.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Otomasi Pengolahan Nilai dan Kehadiran')
    parser.add_argument('file_nilai', nargs='?', help='Path ke file nilai (.csv atau .xlsx)')
    parser.add_argument('file_absensi', nargs='?', help='Path ke file absensi (.csv atau .xlsx)')
    args = parser.parse_args()

    # Jika argumen tidak diberikan, tanyakan interaktif di terminal
    file_nilai = args.file_nilai
    if not file_nilai:
        file_nilai = prompt_for_file('Masukkan path file nilai (.csv/.xlsx) atau tekan Enter untuk lewati: ')

    file_absensi = args.file_absensi
    if not file_absensi:
        file_absensi = prompt_for_file('Masukkan path file kehadiran (.csv/.xlsx) atau tekan Enter untuk lewati: ')

    if not file_nilai and not file_absensi:
        print('Tidak ada file diberikan. Keluar.')
        sys.exit(0)

    if file_nilai:
        proses_nilai(file_nilai)
    else:
        print('Lewati proses nilai.')

    if file_absensi:
        proses_kehadiran(file_absensi)
    else:
        print('Lewati proses kehadiran.')