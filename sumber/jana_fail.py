import random
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

OUT = "Latihan_Claude_IAB/"
rng = np.random.default_rng(2026)
random.seed(2026)

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def write_sheet(ws, headers, rows, widths=None):
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = BORDER
            if cell.row > 1:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    for i, h in enumerate(headers, 1):
        w = (widths[i - 1] if widths else max(12, min(45, len(str(h)) + 4)))
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------- A1 Profil Program ----------------
d = Document()
st = d.styles["Normal"]
st.font.name, st.font.size = "Arial", Pt(11)
t = d.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("INSTITUT AMINUDDIN BAKI GENTING HIGHLANDS\nPROFIL PROGRAM KEPIMPINAN PENGETUA BAHARU SIRI 3/2026")
r.bold = True
d.add_paragraph("(Dokumen latihan. Semua data adalah rekaan.)").alignment = WD_ALIGN_PARAGRAPH.CENTER

d.add_heading("1. Maklumat Program", level=2)
tbl = d.add_table(rows=0, cols=2)
tbl.style = "Table Grid"
for k, v in [
    ("Nama program", "Program Kepimpinan Pengetua Baharu (PKPB) Siri 3/2026"),
    ("Tarikh", "7 – 11 Disember 2026 (Isnin – Jumaat)"),
    ("Tempat", "Institut Aminuddin Baki Genting Highlands, Pahang"),
    ("Kumpulan sasaran", "40 pengetua dan guru besar yang dilantik dalam tempoh 3 tahun"),
    ("Mod", "Bersemuka, berasrama penuh"),
    ("Peruntukan", "RM48,000.00"),
    ("Ketua Urus Setia", "[NAMA PEGAWAI], Pusat Kepimpinan Sekolah"),
    ("Siri sebelum", "Siri 2/2026 (13 – 17 Julai 2026), 38 peserta"),
]:
    row = tbl.add_row().cells
    row[0].text, row[1].text = k, v
    row[0].paragraphs[0].runs[0].bold = True

d.add_heading("2. Latar Belakang", level=2)
d.add_paragraph(
    "Pengetua dan guru besar baharu memerlukan sokongan berstruktur dalam tahun-tahun awal "
    "perlantikan. Dapatan Siri 2/2026 menunjukkan keperluan tinggi dalam kepimpinan instruksional, "
    "pengurusan kewangan dan pengurusan hubungan dengan komuniti. Program ini direka sebagai "
    "intervensi lima hari yang menggabungkan kuliah, lawatan sekolah, perbincangan kumpulan dan "
    "pembentangan pelan tindakan.")

d.add_heading("3. Objektif Program", level=2)
d.add_paragraph("Di akhir program, peserta dapat:")
objs = [
    "O1: Mengamalkan kepimpinan instruksional dan membangunkan profesionalisme guru melalui coaching dan mentoring.",
    "O2: Mengurus kurikulum, pengajaran dan pembelajaran serta menggunakan data untuk membuat keputusan.",
    "O3: Mengurus kewangan, aset dan perolehan sekolah mengikut peraturan.",
    "O4: Mengurus konflik dan krisis di sekolah secara berkesan.",
    "O5: Membina kerjasama strategik dengan PIBG dan komuniti setempat.",
]
for o in objs:
    d.add_paragraph(o, style="List Bullet")

d.add_heading("4. Kandungan dan Kaedah", level=2)
d.add_paragraph(
    "Kuliah dan bengkel oleh penceramah jemputan dan fasilitator IAB, satu lawatan sekolah, "
    "perbincangan kumpulan (8 kumpulan x 5 peserta) dan pembentangan pelan tindakan pada hari terakhir. "
    "Rujuk Jadual Program Draf untuk butiran sesi.")

d.add_heading("5. Penilaian", level=2)
for p in [
    "Penilaian reaksi: borang penilaian bagi setiap sesi utama dan item umum (skala 1–5).",
    "Penilaian pembelajaran: ujian pra dan pasca (50 markah).",
    "Penilaian pelaksanaan: kehadiran, perbelanjaan dan ketepatan masa laporan.",
]:
    d.add_paragraph(p, style="List Bullet")

d.add_heading("6. Sasaran KPI", level=2)
k = d.add_table(rows=1, cols=2)
k.style = "Table Grid"
k.rows[0].cells[0].text, k.rows[0].cells[1].text = "KPI", "Sasaran"
for a, b in [
    ("Kehadiran peserta", "≥ 95%"),
    ("Min kepuasan keseluruhan (1–5)", "≥ 4.20"),
    ("% peserta skor pasca ≥ 35/50", "≥ 80%"),
    ("Peningkatan min pra-pasca", "≥ 8 markah"),
    ("Perbelanjaan berbanding peruntukan", "≤ 100%"),
    ("Laporan program dihantar", "≤ 14 hari selepas program"),
]:
    c = k.add_row().cells
    c[0].text, c[1].text = a, b
d.save(OUT + "01_Profil_Program_PKPB3.docx")

# ---------------- A2 Senarai Penceramah ----------------
speakers = [
    ("PC01", "Dr. Ahmad Faizal bin Rahman", "Pensyarah Kanan", "Institut Aminuddin Baki", "Kepimpinan Instruksional", "7 Dis 2026, 8 Dis 2026", "faizal.rahman@contoh.gov.my"),
    ("PC02", "Pn. Noraini binti Ismail", "Pengetua Cemerlang", "SMK Taman Seri Kembangan (rekaan)", "Pengurusan Kurikulum dan PdP", "8 Dis 2026, 9 Dis 2026", "noraini.ismail@contoh.edu.my"),
    ("PC03", "En. Lim Wei Ming", "Ketua Penolong Pengarah", "Bahagian Kewangan (rekaan)", "Pengurusan Kewangan Sekolah", "9 Dis 2026", "lim.weiming@contoh.gov.my"),
    ("PC04", "Dr. Siti Hajar binti Mohd Noor", "Pensyarah", "Universiti Pendidikan (rekaan)", "Coaching dan Mentoring", "9 Dis 2026, 10 Dis 2026", "sitihajar@contoh.edu.my"),
    ("PC05", "Tn. Hj. Kamarul Ariffin bin Yusof", "Bekas Pengarah Pendidikan Negeri", "Perunding Bebas", "Hubungan Komuniti dan PIBG", "10 Dis 2026, 11 Dis 2026", "kamarul.ariffin@contoh.com"),
]
wb = Workbook()
ws = wb.active
ws.title = "Penceramah"
write_sheet(ws, ["ID", "Nama", "Jawatan", "Organisasi", "Tajuk Sesi", "Tarikh Tersedia", "E-mel"],
            speakers, [7, 34, 28, 34, 30, 24, 30])
wb.save(OUT + "02_Senarai_Penceramah.xlsx")

# ---------------- A4 Jadual Program Draf (4 kesilapan sengaja) ----------------
# Ralat 1: sesi 11 penceramah tiada dalam A2 (En. Rosli bin Hamid)
# Ralat 2: sesi 5 Dr. Siti Hajar pada 8 Dis (tersedia 9 & 10 Dis sahaja)
# Ralat 3: sesi 7 & 8 bertindih di Bilik Kuliah 2 pada 9 Dis
# Ralat 4: O4 tiada sesi berkaitan
sched = [
    (1, "Hari 1", "7 Dis 2026", "9.00 – 10.30 pagi", "Taklimat Program dan Ujian Pra", "Urus Setia IAB", "Dewan Seri Cempaka", "-"),
    (2, "Hari 1", "7 Dis 2026", "11.00 pagi – 1.00 tgh", "Kepimpinan Instruksional", "Dr. Ahmad Faizal bin Rahman", "Dewan Seri Cempaka", "O1"),
    (3, "Hari 1", "7 Dis 2026", "2.30 – 4.30 ptg", "Pembentukan Kumpulan dan Aktiviti Suai Kenal", "Fasilitator IAB", "Bilik Kuliah 1", "-"),
    (4, "Hari 2", "8 Dis 2026", "9.00 pagi – 12.30 tgh", "Pengurusan Kurikulum dan PdP", "Pn. Noraini binti Ismail", "Dewan Seri Cempaka", "O2"),
    (5, "Hari 2", "8 Dis 2026", "2.30 – 4.30 ptg", "Coaching dan Mentoring", "Dr. Siti Hajar binti Mohd Noor", "Bilik Kuliah 1", "O1"),
    (6, "Hari 2", "8 Dis 2026", "4.30 – 5.30 ptg", "Refleksi Harian", "Fasilitator IAB", "Bilik Kuliah 1", "-"),
    (7, "Hari 3", "9 Dis 2026", "9.00 pagi – 12.30 tgh", "Pengurusan Kewangan Sekolah", "En. Lim Wei Ming", "Bilik Kuliah 2", "O3"),
    (8, "Hari 3", "9 Dis 2026", "11.00 pagi – 12.30 tgh", "Data dan Pelaporan Sekolah", "Fasilitator IAB", "Bilik Kuliah 2", "O2"),
    (9, "Hari 3", "9 Dis 2026", "2.30 – 5.00 ptg", "Lawatan Sekolah", "Fasilitator IAB", "Sekolah sekitar Bentong", "O1, O2"),
    (10, "Hari 4", "10 Dis 2026", "9.00 – 11.00 pagi", "Hubungan Komuniti dan PIBG", "Tn. Hj. Kamarul Ariffin bin Yusof", "Dewan Seri Cempaka", "O5"),
    (11, "Hari 4", "10 Dis 2026", "11.30 pagi – 1.00 tgh", "Pengurusan Aset dan Perolehan", "En. Rosli bin Hamid", "Bilik Kuliah 1", "O3"),
    (12, "Hari 4", "10 Dis 2026", "2.30 – 4.30 ptg", "Perbincangan Kumpulan: Pelan Strategik Sekolah", "Fasilitator IAB", "Bilik Kuliah 1 & 2", "O1, O2, O5"),
    (13, "Hari 5", "11 Dis 2026", "8.30 – 10.30 pagi", "Pembentangan Pelan Tindakan Kumpulan", "Fasilitator IAB", "Dewan Seri Cempaka", "O1, O2, O3, O5"),
    (14, "Hari 5", "11 Dis 2026", "10.45 – 11.45 pagi", "Ujian Pasca dan Penilaian Program", "Urus Setia IAB", "Dewan Seri Cempaka", "-"),
    (15, "Hari 5", "11 Dis 2026", "11.45 pagi – 12.30 tgh", "Majlis Penutupan", "Pengarah IAB", "Dewan Seri Cempaka", "-"),
]
wb = Workbook()
ws = wb.active
ws.title = "Jadual Draf"
write_sheet(ws, ["Bil", "Hari", "Tarikh", "Masa", "Sesi", "Penceramah / Fasilitator", "Bilik", "Objektif Berkaitan"],
            sched, [5, 8, 13, 24, 42, 32, 24, 16])
wb.save(OUT + "04_Jadual_Program_Draf.xlsx")

# ---------------- A6 Senarai Peserta ----------------
negeri = ["Johor", "Kedah", "Kelantan", "Melaka", "Negeri Sembilan", "Pahang", "Perak", "Perlis",
          "Pulau Pinang", "Sabah", "Sarawak", "Selangor", "Terengganu", "WP Kuala Lumpur"]
male = ["Mohd Hafiz bin Abdullah", "Azman bin Salleh", "Rajesh a/l Subramaniam", "Tan Chee Keong",
        "Khairul Anuar bin Zakaria", "Shahrul Nizam bin Idris", "Jaafar bin Mansor", "Lee Kok Wai",
        "Mohd Rizal bin Hassan", "Zulkifli bin Omar", "Ganesan a/l Muniandy", "Awang Ismail bin Awang Bakar",
        "Faizul bin Ahmad", "Ong Boon Hock", "Roslan bin Daud", "Nazri bin Kamaruddin",
        "Johnny anak Jalong", "Hasbullah bin Mat Isa", "Amirul bin Rashid"]
female = ["Siti Aminah binti Yusof", "Nor Azlina binti Ahmad", "Kavitha a/p Ramasamy", "Wong Mei Ling",
          "Rohana binti Ibrahim", "Faridah binti Hashim", "Zuraidah binti Musa", "Chong Li Ping",
          "Nurul Huda binti Ismail", "Salmah binti Ali", "Mariam binti Osman", "Juliana anak Linggi",
          "Haslinda binti Mohd Zain", "Norhayati binti Razak", "Lim Siew Lan", "Aishah binti Jamal",
          "Rosnah binti Abu Bakar", "Sharifah Nadia binti Syed Ahmad", "Tengku Farah binti Tengku Ismail",
          "Halimah binti Sulaiman", "Saraswathy a/p Krishnan"]
names = [(n, "Lelaki") for n in male] + [(n, "Perempuan") for n in female]
random.shuffle(names)
participants = []
for i, (nm, jt) in enumerate(names, 1):
    pid = f"P{i:02d}"
    neg = negeri[(i * 5) % len(negeri)]
    jenis = "SMK" if i % 5 in (0, 2, 4) else "SK"
    peng = "< 1 tahun" if i % 3 == 0 else "1 – 3 tahun"
    participants.append([pid, nm, jt, neg, jenis, peng])
wb = Workbook()
ws = wb.active
ws.title = "Peserta"
write_sheet(ws, ["ID", "Nama (rekaan)", "Jantina", "Negeri", "Jenis Sekolah", "Pengalaman sebagai Pengetua/GB"],
            participants, [6, 36, 11, 18, 13, 18])
wb.save(OUT + "06_Senarai_Peserta.xlsx")

# ---------------- A3 Komen Peserta ----------------
comments = [
    "Sesi kepimpinan instruksional sangat praktikal. Contoh kes sekolah sebenar membantu.",
    "Masa sesi petang terlalu panjang. Peserta penat selepas 4.30 petang.",
    "Penceramah sesi pengurusan kewangan membaca slaid sahaja.",
    "Bilik penginapan selesa tetapi Wi-Fi lemah di blok asrama.",
    "Mohon lebih banyak masa untuk perbincangan kumpulan.",
    "Lawatan sekolah amat bermanfaat. Cadang dilanjutkan sehari.",
    "Bahan rujukan lewat diedarkan.",
    "Sesi coaching dan mentoring membuka minda saya.",
    "Makanan baik. Jadual makan kadang-kadang lewat.",
    "Perlu sesi khusus tentang pengurusan konflik dengan PIBG.",
    "Sesi kewangan terlalu teknikal dan banyak pekeliling, kurang contoh di sekolah.",
    "Fasilitator kumpulan sangat membantu dan mesra.",
    "Sesi data dan pelaporan bertindih dengan sesi kewangan, jadi masa terlalu singkat.",
    "Saya mahu templat pelan strategik yang boleh terus diguna di sekolah.",
    "Pengangkutan ke lawatan sekolah lewat 30 minit.",
    "Sesi PIBG sangat relevan, penceramah berpengalaman luas.",
    "Kandungan aset dan perolehan berguna tetapi penyampaian kurang menarik.",
    "Program tersusun dan urus setia responsif.",
    "Cadangkan sesi pengurusan krisis seperti kes buli dan tular media sosial.",
    "Ujian pasca terlalu panjang untuk masa yang diberi.",
]
with open(OUT + "03_Komen_Peserta.txt", "w", encoding="utf-8") as f:
    f.write("KOMEN TERBUKA PESERTA – PROGRAM KEPIMPINAN PENGETUA BAHARU SIRI 3/2026\n")
    f.write("Soalan: Nyatakan ulasan dan cadangan penambahbaikan anda.\n(Data latihan rekaan)\n\n")
    for i, c in enumerate(comments, 1):
        f.write(f"P{i:02d}: {c}\n")

# ---------------- A7 Penilaian mentah (data kotor) ----------------
eval_sessions = [
    ("S02 Kepimpinan Instruksional", 4.50), ("S04 Kurikulum dan PdP", 4.25),
    ("S05 Coaching dan Mentoring", 4.40), ("S07 Pengurusan Kewangan", 3.35),
    ("S08 Data dan Pelaporan", 3.40), ("S09 Lawatan Sekolah", 4.60),
    ("S10 Komuniti dan PIBG", 4.35), ("S11 Aset dan Perolehan", 3.50),
    ("S12 Pelan Strategik", 4.15), ("S13 Pembentangan Pelan Tindakan", 4.20),
]
general = [("U1 Objektif program jelas", 4.30), ("U2 Kandungan relevan dengan tugas", 4.20),
           ("U3 Logistik dan pengurusan masa", 3.70), ("U4 Penginapan dan makanan", 3.95),
           ("U5 Kepuasan keseluruhan", 3.90)]
items = eval_sessions + general


def gen_score(mu, adj):
    v = rng.normal(mu + adj, 0.65)
    return int(min(5, max(1, round(v))))


neg_variants = {"Johor": ["Johor", "JOHOR", "Jhr"], "Selangor": ["Selangor", "SELANGOR", "Sel."],
                "WP Kuala Lumpur": ["WP Kuala Lumpur", "Kuala Lumpur", "KL"],
                "Pulau Pinang": ["Pulau Pinang", "P. Pinang", "Penang"]}
date_formats = ["11/12/2026", "2026-12-11", "11 Dis 2026", "11.12.2026"]
eval_rows = []
for p in participants:
    pid, _, jt, neg, jenis, peng = p
    adj = 0.12 if jenis == "SK" else -0.12
    adj += 0.05 if peng.startswith("<") else 0
    scores = [gen_score(mu, adj) for _, mu in items]
    idx = int(pid[1:])
    neg_out = random.choice(neg_variants[neg]) if neg in neg_variants else neg
    komen = comments[idx - 1] if idx <= 20 else random.choice(
        ["", "Tiada.", "Teruskan program ini.", "Terima kasih urus setia.", "Masa rehat perlu ditambah.", "-"])
    eval_rows.append([pid, jt, neg_out, jenis, peng, random.choice(date_formats)] + scores + [komen])

# Out-of-range values (3)
eval_rows[6][6 + 3] = 0     # P07 S07
eval_rows[17][6 + 7] = 6    # P18 S11
eval_rows[28][6 + 14] = 55  # P29 U5
# Empty cells (5)
for r, c in [(3, 6 + 1), (11, 6 + 4), (19, 6 + 9), (24, 6 + 12), (33, 6 + 7)]:
    eval_rows[r][c] = None
# Duplicates (2)
dups = [list(eval_rows[9]), list(eval_rows[21])]
eval_rows.insert(15, dups[0])
eval_rows.append(dups[1])

wb = Workbook()
ws = wb.active
ws.title = "Respons"
headers = ["ID", "Jantina", "Negeri", "Jenis Sekolah", "Pengalaman", "Tarikh Hantar"] + [n for n, _ in items] + ["Komen"]
write_sheet(ws, headers, eval_rows, [6, 11, 16, 9, 13, 13] + [13] * len(items) + [55])
ws.row_dimensions[1].height = 45
ws2 = wb.create_sheet("Skala")
write_sheet(ws2, ["Skor", "Maksud"], [[1, "Sangat tidak setuju"], [2, "Tidak setuju"], [3, "Kurang setuju"],
                                      [4, "Setuju"], [5, "Sangat setuju"]], [8, 24])
wb.save(OUT + "07_Penilaian_PKPB3_mentah.xlsx")

# ---------------- A8 Ujian Pra Pasca ----------------
pp_rows = []
pre = np.clip(rng.normal(26, 5, 40).round(), 10, 42).astype(int)
gain = np.clip(rng.normal(10.5, 3.5, 40).round(), 1, None).astype(int)
for i in (7, 23, 36):  # 3 peserta menurun
    gain[i] = -int(rng.integers(1, 4))
post = np.clip(pre + gain, 0, 50).astype(int)
for i, p in enumerate(participants):
    pp_rows.append([p[0], int(pre[i]), int(post[i])])
wb = Workbook()
ws = wb.active
ws.title = "Skor"
write_sheet(ws, ["ID", "Skor Pra (/50)", "Skor Pasca (/50)"], pp_rows, [6, 15, 16])
wb.save(OUT + "08_Ujian_Pra_Pasca_PKPB3.xlsx")

# ---------------- A9 Data Operasi ----------------
wb = Workbook()
ws = wb.active
ws.title = "Kehadiran"
absent = {("P05", 2), ("P12", 4), ("P12", 5), ("P19", 3), ("P27", 1), ("P33", 4), ("P38", 5)}
att_rows = []
for p in participants:
    att_rows.append([p[0]] + [("Tidak hadir" if (p[0], d_) in absent else "Hadir") for d_ in range(1, 6)])
write_sheet(ws, ["ID", "Hari 1 (7 Dis)", "Hari 2 (8 Dis)", "Hari 3 (9 Dis)", "Hari 4 (10 Dis)", "Hari 5 (11 Dis)"],
            att_rows, [6, 15, 15, 15, 15, 15])

ws = wb.create_sheet("Kewangan")
fin = [
    ("Penginapan dan makan minum", 30000, 29640),
    ("Saguhati penceramah", 6000, 5400),
    ("Bahan kursus dan cetakan", 4000, 4385),
    ("Pengangkutan lawatan sekolah", 3500, 3200),
    ("Cenderamata penceramah", 1500, 1325),
    ("Pelbagai / kontingensi", 3000, 2900),
]
fin_rows = [[a, b, c] for a, b, c in fin]
write_sheet(ws, ["Perkara", "Peruntukan (RM)", "Perbelanjaan Sebenar (RM)"], fin_rows, [34, 18, 26])
n = len(fin_rows) + 1
ws.append(["JUMLAH", f"=SUM(B2:B{n})", f"=SUM(C2:C{n})"])
for c in range(1, 4):
    ws.cell(row=n + 1, column=c).font = Font(bold=True)
for row in ws.iter_rows(min_row=2, min_col=2, max_col=3):
    for cell in row:
        cell.number_format = '#,##0.00'

ws = wb.create_sheet("KPI Siri 2-2026")
write_sheet(ws, ["KPI", "Sasaran", "Pencapaian Siri 2/2026"], [
    ["Kehadiran peserta", "≥ 95%", "94.0%"],
    ["Min kepuasan keseluruhan (1–5)", "≥ 4.20", "4.05"],
    ["% peserta skor pasca ≥ 35/50", "≥ 80%", "72%"],
    ["Peningkatan min pra-pasca", "≥ 8 markah", "7.4"],
    ["Perbelanjaan berbanding peruntukan", "≤ 100%", "101.2%"],
    ["Laporan program dihantar", "≤ 14 hari", "18 hari"],
], [36, 14, 22])
ws = wb.create_sheet("Nota")
ws["A1"] = "Nota: Tarikh laporan Siri 3/2026 dihantar belum direkodkan (data tiada). Data latihan rekaan."
wb.save(OUT + "09_Data_Operasi_PKPB3.xlsx")

# ---------------- A5 Nota Mesyuarat ----------------
nota = """nota mesy JK pelaksana PKPB siri 3 - bil 2/2026
18/11 rabu 3ptg bilik mesy utama

hadir: ketua urus setia (pengerusi), pgw kandungan, pgw logistik, pgw penilaian, pembantu tadbir, pgw ICT
x hadir: TPengarah (urusan rasmi)

- minit lalu ok, sahkan by pgw kandungan, sokong pgw ICT
- berbangkit: penginapan blok B dah confirm 40 bilik?? logistik kata 38 je confirm, 2 lagi tunggu
- surat jemputan penceramah: 4 dah hantar, 1 belum (Tn Hj Kamarul?) - pembantu tadbir follow up minggu ni
- bas lawatan sekolah - sebut harga 2 syarikat, pilih yg murah tp kena check lesen. logistik. sblm 25/11
- ujian pra pasca: pgw penilaian siapkan soalan 30 item?? atau 50 markah - perlu putus. pengerusi kata ikut siri 2 (50 markah)
- wifi asrama lemah siri lepas - ICT cek router blok B, tambah AP kalau boleh. tarikh?
- bahan kursus: cetak 45 set (40 + 5 simpanan). pgw kandungan hantar softcopy ke percetakan 1/12
- sesi pengurusan konflik - ada cadangan tambah, pengerusi minta pgw kandungan kaji, bawa mesy akan datang
- borang penilaian online / kertas? setuju online (google form?) + backup kertas. penilaian
- hal lain: cenderamata penceramah - beli sblm 3/12. pembantu tadbir
- mesy akan datang 2/12 (selasa) 10pg

tamat 4.40ptg
"""
with open(OUT + "05_Nota_Mesyuarat_JK.txt", "w", encoding="utf-8") as f:
    f.write(nota)

print("Siap")
