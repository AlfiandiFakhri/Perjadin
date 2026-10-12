import io
import os
import copy
import datetime
import zipfile
import json
import re
import streamlit as st
import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor, Cm
from docxtpl import DocxTemplate

# ==============================================================================
# 1. KONFIGURASI HALAMAN & STYLING (GLOBAL)
# ==============================================================================
st.set_page_config(
    page_title="Sistem Dokumen Perjadin",
    page_icon="📋",
    layout="wide",
)

st.markdown("""
    <style>
        .main-header {
            background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
            padding: 20px;
            border-radius: 10px;
            color: white;
            text-align: center;
            margin-bottom: 25px;
        }
        .main-header h1 {
            font-size: 24px;
            margin: 0;
            font-weight: 600;
        }
        .main-header p {
            font-size: 14px;
            margin: 5px 0 0 0;
            opacity: 0.85;
        }
        div[data-testid="stTabs"] button {
            font-size: 16px;
            font-weight: 600;
            padding-top: 10px;
            padding-bottom: 10px;
        }
    </style>
    <div class="main-header">
        <h1>SISTEM PENGELOLAAN DOKUMEN PERJALANAN DINAS</h1>
        <p>Kementerian Usaha Mikro, Kecil, dan Menengah Republik Indonesia</p>
    </div>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. DATABASE & KONSTANTA (GABUNGAN)
# ==============================================================================
BULAN_INDO = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus", 
    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

# --- Data Modul 1: Nota Dinas & Surat Tugas ---
DATA_PEGAWAI_JABATAN = {
    "Hengki Andrianto": "Asisten Deputi Pembiayaan dan Investasi Usaha Menengah",
    "Ira Dunggio": "Kepala Bidang Renbang Pembiayaan dan Investasi Usaha Menengah",
    "Nadia Wulandari": "Perencana Ahli Muda",
    "Nurmina": "Pengembang Kewirausahaan Ahli Muda",
    "Feris Banda Mukti": "Analis Kebijakan Ahli Pertama",
    "Fakhri Alfiandi": "Pengembang Kewirausahaan Ahli Pertama",
    "Iqbal Mumtaz Fahmy": "Pengembang Kewirausahaan Ahli Pertama",
    "Tatag Suryo Pambudi": "Pengembang Kewirausahaan Ahli Pertama",
    "Maya Annisa": "Pengembang Kewirausahaan Ahli Pertama",
    "Beny Johani": "Pelaksana",
    "Lainnya (Ketik Manual...)": "Lainnya"
}

DATA_WILAYAH = {
    "Nanggroe Aceh Darussalam (NAD)": ["Banda Aceh", "Sabang", "Lhokseumawe", "Langsa", "Subulussalam", "Kab. Aceh Besar", "Kab. Pidie", "Kab. Bireuen", "Kab. Aceh Utara", "Kab. Aceh Timur", "Kab. Aceh Tengah", "Kab. Aceh Barat", "Kab. Aceh Selatan", "Kab. Aceh Singkil", "Kab. Aceh Tamiang", "Kab. Gayo Lues"],
    "Sumatera Utara": ["Medan", "Binjai", "Tebing Tinggi", "Pematangsiantar", "Tanjungbalai", "Sibolga", "Padangsidimpuan", "Gunungsitoli", "Kab. Deli Serdang", "Kab. Langkat", "Kab. Karo", "Kab. Simalungun", "Kab. Asahan", "Kab. Tapanuli Utara", "Kab. Tapanuli Selatan", "Kab. Tapanuli Tengah", "Kab. Nias"],
    "Sumatera Barat": ["Padang", "Bukittinggi", "Payakumbuh", "Solok", "Sawahlunto", "Padang Panjang", "Pariaman", "Kab. Agam", "Kab. Pesisir Selatan", "Kab. Lima Puluh Kota", "Kab. Tanah Datar", "Kab. Padang Pariaman", "Kab. Solok Selatan", "Kab. Dharmasraya"],
    "Riau": ["Pekanbaru", "Dumai", "Kab. Kampar", "Kab. Rokan Hulu", "Kab. Rokan Hilir", "Kab. Pelalawan", "Kab. Siak", "Kab. Bengkalis", "Kab. Indragiri Hilir", "Kab. Indragiri Hulu", "Kab. Kuantan Singingi"],
    "Kepulauan Riau": ["Batam", "Tanjungpinang", "Kab. Bintan", "Kab. Karimun", "Kab. Natuna", "Kab. Kepulauan Anambas", "Kab. Lingga"],
    "Jambi": ["Jambi", "Sungai Penuh", "Kab. Batanghari", "Kab. Bungo", "Kab. Kerinci", "Kab. Merangin", "Kab. Muaro Jambi", "Kab. Sarolangun", "Kab. Tanjung Jabung Barat", "Kab. Tanjung Jabung Timur", "Kab. Tebo"],
    "Bengkulu": ["Bengkulu", "Kab. Bengkulu Selatan", "Kab. Bengkulu Tengah", "Kab. Bengkulu Utara", "Kab. Kaur", "Kab. Kepahiang", "Kab. Lebong", "Kab. Mukomuko", "Kab. Rejang Lebong", "Kab. Seluma"],
    "Sumatera Selatan": ["Palembang", "Prabumulih", "Pagar Alam", "Lubuklinggau", "Kab. Banyuasin", "Kab. Empat Lawang", "Kab. Lahat", "Kab. Muara Enim", "Kab. Musi Banyuasin", "Kab. Musi Rawas", "Kab. Ogan Ilir", "Kab. Ogan Komering Ilir", "Kab. Ogan Komering Ulu"],
    "Kepulauan Bangka Belitung": ["Pangkalpinang", "Kab. Bangka", "Kab. Bangka Barat", "Kab. Bangka Selatan", "Kab. Bangka Tengah", "Kab. Belitung", "Kab. Belitung Timur"],
    "Lampung": ["Bandar Lampung", "Metro", "Kab. Lampung Barat", "Kab. Lampung Selatan", "Kab. Lampung Tengah", "Kab. Lampung Timur", "Kab. Lampung Utara", "Kab. Mesuji", "Kab. Pesawaran", "Kab. Pringsewu", "Kab. Tanggamus", "Kab. Tulang Bawang", "Kab. Way Kanan"],
    "Banten": ["Serang", "Cilegon", "Tangerang", "Tangerang Selatan", "Kab. Lebak", "Kab. Pandeglang", "Kab. Serang", "Kab. Tangerang"],
    "DKI Jakarta": ["Jakarta Pusat", "Jakarta Utara", "Jakarta Barat", "Jakarta Selatan", "Jakarta Timur", "Kepulauan Seribu"],
    "Jawa Barat": ["Bandung", "Bekasi", "Bogor", "Cimahi", "Cirebon", "Depok", "Sukabumi", "Tasikmalaya", "Banjar", "Kab. Bandung", "Kab. Bandung Barat", "Kab. Bekasi", "Kab. Bogor", "Kab. Ciamis", "Kab. Cianjur", "Kab. Cirebon", "Kab. Garut", "Kab. Indramayu", "Kab. Karawang", "Kab. Kuningan", "Kab. Majalengka", "Kab. Pangandaran", "Kab. Purwakarta", "Kab. Subang", "Kab. Sukabumi", "Kab. Sumedang", "Kab. Tasikmalaya"],
    "Jawa Tengah": ["Semarang", "Surakarta (Solo)", "Magelang", "Pekalongan", "Salatiga", "Tegal", "Kab. Banyumas", "Kab. Batang", "Kab. Blora", "Kab. Boyolali", "Kab. Brebes", "Kab. Cilacap", "Kab. Demak", "Kab. Grobogan", "Kab. Jepara", "Kab. Karanganyar", "Kab. Kebumen", "Kab. Kendal", "Kab. Klaten", "Kab. Kudus", "Kab. Magelang", "Kab. Pati", "Kab. Pekalongan", "Kab. Pemalang", "Kab. Purbalingga", "Kab. Purworejo", "Kab. Rembang", "Kab. Semarang", "Kab. Sragen", "Kab. Sukoharjo", "Kab. Tegal", "Kab. Temanggung", "Kab. Wonogiri", "Kab. Wonosobo"],
    "DI Yogyakarta": ["Yogyakarta", "Kab. Bantul", "Kab. Gunungkidul", "Kab. Kulon Progo", "Kab. Sleman"],
    "Jawa Timur": ["Surabaya", "Malang", "Batu", "Kediri", "Blitar", "Madiun", "Mojokerto", "Pasuruan", "Probolinggo", "Kab. Bangkalan", "Kab. Banyuwangi", "Kab. Blitar", "Kab. Bojonegoro", "Kab. Bondowoso", "Kab. Gresik", "Kab. Jember", "Kab. Jombang", "Kab. Kediri", "Kab. Lamongan", "Kab. Lumajang", "Kab. Madiun", "Kab. Magetan", "Kab. Malang", "Kab. Mojokerto", "Kab. Nganjuk", "Kab. Ngawi", "Kab. Pacitan", "Kab. Pamekasan", "Kab. Pasuruan", "Kab. Ponorogo", "Kab. Probolinggo", "Kab. Sampang", "Kab. Sidoarjo", "Kab. Situbondo", "Kab. Sumenep", "Kab. Trenggalek", "Kab. Tuban", "Kab. Tulungagung"],
    "Bali": ["Denpasar", "Kab. Badung", "Kab. Bangli", "Kab. Buleleng", "Kab. Gianyar", "Kab. Jembrana", "Kab. Karangasem", "Kab. Klungkung", "Kab. Tabanan"],
    "Nusa Tenggara Barat (NTB)": ["Mataram", "Bima", "Kab. Bima", "Kab. Dompu", "Kab. Lombok Barat", "Kab. Lombok Tengah", "Kab. Lombok Timur", "Kab. Lombok Utara", "Kab. Sumbawa", "Kab. Sumbawa Barat"],
    "Nusa Tenggara Timur (NTT)": ["Kupang", "Kab. Alor", "Kab. Belu", "Kab. Ende", "Kab. Flores Timur", "Kab. Kupang", "Kab. Lembata", "Kab. Malaka", "Kab. Manggarai", "Kab. Manggarai Barat", "Kab. Manggarai Timur", "Kab. Nagekeo", "Kab. Ngada", "Kab. Rote Ndao", "Kab. Sabu Raijua", "Kab. Sikka", "Kab. Sumba Barat", "Kab. Sumba Barat Daya", "Kab. Sumba Tengah", "Kab. Sumba Timur", "Kab. Timor Tengah Selatan", "Kab. Timor Tengah Utara"],
    "Kalimantan Barat": ["Pontianak", "Singkawang", "Kab. Bengkayang", "Kab. Kapuas Hulu", "Kab. Kayong Utara", "Kab. Ketapang", "Kab. Kubu Raya", "Kab. Landak", "Kab. Melawi", "Kab. Mempawah", "Kab. Sambas", "Kab. Sanggau", "Kab. Sekadau", "Kab. Sintang"],
    "Kalimantan Tengah": ["Palangka Raya", "Kab. Barito Selatan", "Kab. Barito Timur", "Kab. Barito Utara", "Kab. Gunung Mas", "Kab. Kapuas", "Kab. Katingan", "Kab. Kotawaringin Barat", "Kab. Kotawaringin Timur", "Kab. Lamandau", "Kab. Murung Raya", "Kab. Pulang Pisau", "Kab. Sukamara", "Kab. Seruyan"],
    "Kalimantan Selatan": ["Banjarmasin", "Banjarbaru", "Kab. Balangan", "Kab. Banjar", "Kab. Barito Kuala", "Kab. Hulu Sungai Selatan", "Kab. Hulu Sungai Tengah", "Kab. Hulu Sungai Utara", "Kab. Kotabaru", "Kab. Tabalong", "Kab. Tanah Bumbu", "Kab. Tanah Laut", "Kab. Tapin"],
    "Kalimantan Timur": ["Samarinda", "Balikpapan", "Bontang", "Nusantara (IKN)", "Kab. Berau", "Kab. Kutai Barat", "Kab. Kutai Kartanegara", "Kab. Kutai Timur", "Kab. Mahakam Ulu", "Kab. Paser", "Kab. Penajam Paser Utara"],
    "Kalimantan Utara": ["Tarakan", "Kab. Bulungan", "Kab. Malinau", "Kab. Nunukan", "Kab. Tana Tidung"],
    "Sulawesi Utara": ["Manado", "Bitung", "Tomohon", "Kotamobagu", "Kab. Bolaang Mongondow", "Kab. Minahasa", "Kab. Minahasa Selatan", "Kab. Minahasa Utara", "Kab. Kepulauan Sangihe", "Kab. Kepulauan Talaud"],
    "Gorontalo": ["Gorontalo", "Kab. Boalemo", "Kab. Bone Bolango", "Kab. Gorontalo", "Kab. Gorontalo Utara", "Kab. Pohuwato"],
    "Sulawesi Tengah": ["Palu", "Kab. Banggai", "Kab. Donggala", "Kab. Morowali", "Kab. Poso", "Kab. Sigi", "Kab. Tojo Una-Una", "Kab. Tolitoli", "Kab. Parigi Moutong"],
    "Sulawesi Barat": ["Kab. Majene", "Kab. Mamasa", "Kab. Mamuju", "Kab. Mamuju Tengah", "Kab. Pasangkayu", "Kab. Polewali Mandar"],
    "Sulawesi Selatan": ["Makassar", "Parepare", "Palopo", "Kab. Bantaeng", "Kab. Barru", "Kab. Bone", "Kab. Bulukumba", "Kab. Enrekang", "Kab. Gowa", "Kab. Jeneponto", "Kab. Luwu", "Kab. Luwu Timur", "Kab. Luwu Utara", "Kab. Maros", "Kab. Pangkajene dan Kepulauan", "Kab. Pinrang", "Kab. Selayar", "Kab. Sinjai", "Kab. Sidenreng Rappang", "Kab. Soppeng", "Kab. Takalar", "Kab. Tana Toraja", "Kab. Toraja Utara", "Kab. Wajo"],
    "Sulawesi Tenggara": ["Kendari", "Baubau", "Kab. Bombana", "Kab. Buton", "Kab. Kolaka", "Kab. Konawe", "Kab. Muna", "Kab. Wakatobi"],
    "Maluku": ["Ambon", "Tual", "Kab. Buru", "Kab. Kepulauan Aru", "Kab. Maluku Tengah", "Kab. Maluku Tenggara", "Kab. Seram Bagian Barat", "Kab. Seram Bagian Timur"],
    "Maluku Utara": ["Ternate", "Tidore Kepulauan", "Kab. Halmahera Barat", "Kab. Halmahera Selatan", "Kab. Halmahera Tengah", "Kab. Halmahera Timur", "Kab. Halmahera Utara", "Kab. Kepulauan Sula", "Kab. Pulau Morotai"],
    "Papua Barat Daya": ["Sorong", "Kab. Sorong", "Kab. Sorong Selatan", "Kab. Raja Ampat", "Kab. Tambrauw", "Kab. Maybrat"],
    "Papua Barat": ["Kab. Manokwari", "Kab. Fakfak", "Kab. Kaimana", "Kab. Teluk Bintuni", "Kab. Teluk Wondama", "Kab. Pegunungan Arfak", "Kab. Manokwari Selatan"],
    "Papua Tengah": ["Kab. Nabire", "Kab. Mimika", "Kab. Paniai", "Kab. Dogiyai", "Kab. Deiyai", "Kab. Intan Jaya", "Kab. Puncak", "Kab. Puncak Jaya"],
    "Papua": ["Jayapura", "Kab. Jayapura", "Kab. Keerom", "Kab. Sarmi", "Kab. Mamberamo Raya", "Kab. Biak Numfor", "Kab. Supiori", "Kab. Kepulauan Yapen", "Kab. Waropen"],
    "Papua Pegunungan": ["Kab. Jayawijaya", "Kab. Lanny Jaya", "Kab. Mamberamo Tengah", "Kab. Nduga", "Kab. Tolikara", "Kab. Yahukimo", "Kab. Yalimo", "Kab. Pegunungan Bintang"],
    "Papua Selatan": ["Kab. Merauke", "Kab. Boven Digoel", "Kab. Mappi", "Kab. Asmat"]
}

# --- Data Modul 2: SPPD ---
DAFTAR_KOTA_KABUPATEN = [
    "Jakarta", "Jakarta Pusat", "Jakarta Selatan", "Jakarta Timur", "Jakarta Barat", "Jakarta Utara", "Kepulauan Seribu",
    "Bogor", "Kota Bogor", "Depok", "Tangerang", "Kota Tangerang", "Tangerang Selatan", "Bekasi", "Kota Bekasi",
    "Bandung", "Kota Bandung", "Bandung Barat", "Banjar", "Ciamis", "Cianjur", "Cirebon", "Kota Cirebon", 
    "Garut", "Indramayu", "Karawang", "Kuningan", "Majalengka", "Sumedang", "Subang", "Purwakarta",
    "Semarang", "Kota Semarang", "Yogyakarta", "Surabaya", "Medan", "Makassar", "Palembang", "Denpasar", "Bandar Lampung"
]
OPI_KOTA_BLANK = ["(Kosong / Belum diisi)"] + sorted(list(set(DAFTAR_KOTA_KABUPATEN)))

data_pegawai_sppd = {
    "Hengki Andrianto": {"pangkat": "IV", "jabatan": "Asisten Deputi Pembiayaan dan Investasi Usaha Menengah"},
    "Ira Dunggio": {"pangkat": "IV", "jabatan": "Kepala Bidang Renbang Pembiayaan dan Investasi Usaha Menengah"},
    "Nadia Wulandari": {"pangkat": "III", "jabatan": "Perencana Ahli Muda"},
    "Nurmina": {"pangkat": "III", "jabatan": "Pengembang Kewirausahaan Ahli Muda"},
    "Feris Banda Mukti": {"pangkat": "III", "jabatan": "Analis Kebijakan Ahli Pertama"},
    "Fakhri Alfiandi": {"pangkat": "III", "jabatan": "Pengembang Kewirausahaan Ahli Pertama"},
    "Iqbal Mumtaz Fahmy": {"pangkat": "III", "jabatan": "Pengembang Kewirausahaan Ahli Pertama"},
    "Tatag Suryo Pambudi": {"pangkat": "III", "jabatan": "Pengembang Kewirausahaan Ahli Pertama"},
    "Maya Annisa": {"pangkat": "III", "jabatan": "Pengembang Kewirausahaan Ahli Pertama"},
    "Taufik Adytia Nugraha": {"pangkat": "-", "jabatan": "Tenaga Ahli"},
    "Bagus Rachman": {"pangkat": "IV", "jabatan": "Deputi Bidang Usaha Menengah"},
    "Beny Johani": {"pangkat": "-", "jabatan": "Driver"}
}

daftar_angkutan = ["Pesawat", "Bus", "Kereta", "Travel", "Mobil"]

data_kegiatan = {
    "Koordinasi, Identifikasi, Persiapan Pelaksanaan Program Perluasan Pembiayaan KUR Klaster": "7475.QDG.001.051.A.524111",
    "Koordinasi, Indentifikasi, Persiapan Pelaksanaan Program Akselerasi Pembiayaan Inovatif": "7475.QDG.001.052.A.524111"
}


# ==============================================================================
# 3. FUNGSI PEMROSESAN (GABUNGAN)
# ==============================================================================

# --- Fungsi Nota Dinas & Surat Tugas ---
def set_document_font(doc, font_name="Arial", size_pt=12):
    style = doc.styles["Normal"]
    font = style.font
    font.name = font_name
    font.size = Pt(size_pt)
    font.color.rgb = RGBColor(0, 0, 0)

def setup_document_layout(section, kop_path, watermark_path, footer_path):
    section.header_distance = Inches(0.0)
    section.footer_distance = Inches(0.1)

    header = section.header
    hp_kop = header.paragraphs[0]
    hp_kop.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(kop_path):
        try:
            hp_kop.add_run().add_picture(kop_path, width=Inches(3.2))
        except Exception as e:
            st.warning(f"Gagal memuat gambar kop surat: {e}")
    if os.path.exists(watermark_path):
        try:
            hp_wm = header.add_paragraph()
            run_wm = hp_wm.add_run()
            run_wm.add_picture(watermark_path, width=Inches(6.5)) 
            drawing = run_wm._r.xpath('.//w:drawing')[0]
            inline = drawing.xpath('.//wp:inline')[0]
            extent = inline.xpath('.//wp:extent')[0]
            cx = extent.get('cx')
            cy = extent.get('cy')
            docPr = inline.xpath('.//wp:docPr')[0]
            docPr_id = docPr.get('id')
            docPr_name = docPr.get('name')
            graphic = copy.deepcopy(inline.xpath('.//a:graphic')[0])
            anchor_xml = f'''
            <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0"
                       behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
                <wp:simplePos x="0" y="0"/>
                <wp:positionH relativeFrom="page">
                    <wp:align>center</wp:align>
                </wp:positionH>
                <wp:positionV relativeFrom="page">
                    <wp:align>center</wp:align>
                </wp:positionV>
                <wp:extent cx="{cx}" cy="{cy}"/>
                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                <wp:wrapNone/>
                <wp:docPr id="{docPr_id}" name="{docPr_name}"/>
                <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks noChangeAspect="1"/>
                </wp:cNvGraphicFramePr>
            </wp:anchor>
            '''
            anchor = parse_xml(anchor_xml)
            anchor.append(graphic)
            drawing.replace(inline, anchor)
        except Exception as e:
            st.warning(f"Gagal memuat watermark floating: {e}")
    footer = section.footer
    hp_footer = footer.paragraphs[0]
    hp_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(footer_path):
        try:
            hp_footer.add_run().add_picture(footer_path, width=Inches(5.8))
        except Exception as e:
            st.warning(f"Gagal memuat gambar footer: {e}")

def generate_nota_dinas(data):
    doc = Document()
    set_document_font(doc, "Arial", 12)

    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        setup_document_layout(section, "kop_surat.png", "watermark.png", "footer_surat.png")

    p_judul = doc.add_paragraph()
    p_judul.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_judul.paragraph_format.space_after = Pt(2)
    r_judul = p_judul.add_run("NOTA DINAS")
    r_judul.bold = True
    r_judul.font.size = Pt(14)
    r_judul.font.name = "Arial"

    p_nomor = doc.add_paragraph()
    p_nomor.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_nomor.paragraph_format.space_after = Pt(6)
    r_nomor = p_nomor.add_run(f"Nomor: {data['nomor_memo']}")
    r_nomor.font.size = Pt(12)
    r_nomor.font.name = "Arial"

    header_table = doc.add_table(rows=3, cols=3)
    for cell in header_table.columns[0].cells:
        cell.width = Inches(1.1) 
    for cell in header_table.columns[1].cells:
        cell.width = Inches(0.2) 
    for cell in header_table.columns[2].cells:
        cell.width = Inches(4.9) 

    headers_data = [
        ("Kepada Yth.", ":", data['tujuan_memo']),
        ("Dari", ":", data['pengirim_memo']),
        ("Perihal", ":", "Permohonan Penerbitan ST"),
    ]

    for i, (label, separator, val) in enumerate(headers_data):
        cell_label = header_table.cell(i, 0)
        cell_separator = header_table.cell(i, 1)
        cell_val = header_table.cell(i, 2)
        
        p_label = cell_label.paragraphs[0]
        p_label.paragraph_format.space_after = Pt(0)
        p_label.paragraph_format.line_spacing = 1.0 
        p_label.add_run(label).font.name = "Arial"
        
        p_separator = cell_separator.paragraphs[0]
        p_separator.paragraph_format.space_after = Pt(0)
        p_separator.paragraph_format.line_spacing = 1.0
        p_separator.add_run(separator).font.name = "Arial"
        
        p_val = cell_val.paragraphs[0]
        p_val.paragraph_format.space_after = Pt(0)
        p_val.paragraph_format.line_spacing = 1.0
        p_val.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_val.add_run(val).font.name = "Arial"

    p_spasi = doc.add_paragraph()
    p_spasi.paragraph_format.space_after = Pt(2) 

    p_pembuka = doc.add_paragraph()
    p_pembuka.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_pembuka.paragraph_format.line_spacing = 1.20
    p_pembuka.paragraph_format.space_after = Pt(4)
    p_pembuka.paragraph_format.first_line_indent = Inches(0.5) 
    
    r_pembuka = p_pembuka.add_run(
        f"Dalam rangka {data['maksud_kegiatan']} dengan (MAK {data['mak_anggaran']}). Bersama ini "
        "disampaikan rencana perjalanan dinas pada Asisten Deputi Pembiayaan dan "
        "Investasi Usaha Menengah Tahun Anggaran 2026, dan dimohonkan dapat "
        "menerbitkan Surat Perintah Tugas atas nama sebagai berikut:"
    )
    r_pembuka.font.name = "Arial"

    table = doc.add_table(rows=1, cols=7)
    table.style = "Table Grid"
    table.autofit = False 
    table.allow_autofit = False 
    table.alignment = WD_TABLE_ALIGNMENT.CENTER  
    
    col_widths = [
        Inches(0.40), # No
        Inches(1.82), # Nama
        Inches(0.44), # Gol
        Inches(0.71), # Waktu
        Inches(1.48), # Tujuan
        Inches(1.69), # Tanggal Dinas
        Inches(1.02)  # Kendaraan
    ]
    
    hdr_cells = table.rows[0].cells
    headers_title = ["No", "Nama", "Gol", "Waktu", "Tujuan", "Tanggal Dinas", "Kendaraan"]
    
    for i, title_text in enumerate(headers_title):
        hdr_cells[i].text = title_text
        hdr_cells[i].width = col_widths[i]
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER 
        for paragraph in hdr_cells[i].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(1) 
            paragraph.paragraph_format.space_after = Pt(1)  
            for run in paragraph.runs:
                run.bold = True
                run.font.name = "Arial"
                run.font.size = Pt(12)

    for idx, peg in enumerate(data['pegawai_list'], start=1):
        row_cells = table.add_row().cells
        row_values = [str(idx), peg['nama'], peg['gol'], f"{peg['waktu_hari']} Hari", peg['kab_final'], peg['tgl_str'], peg['kend_final']]
        for col_idx in range(7):
            row_cells[col_idx].text = row_values[col_idx]
            row_cells[col_idx].width = col_widths[col_idx]
            row_cells[col_idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            for paragraph in row_cells[col_idx].paragraphs:
                paragraph.paragraph_format.space_before = Pt(1)
                paragraph.paragraph_format.space_after = Pt(1)
                paragraph.paragraph_format.line_spacing = 1.0
                
                if col_idx in [0, 2, 3, 4]: 
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                for run in paragraph.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(12)

    p_penutup = doc.add_paragraph()
    p_penutup.paragraph_format.space_before = Pt(4)
    p_penutup.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_penutup.paragraph_format.line_spacing = 1.0
    p_penutup.paragraph_format.space_after = Pt(2)
    p_penutup.paragraph_format.first_line_indent = Inches(0.5) 
    r_penutup = p_penutup.add_run("\nDemikian disampaikan, atas kerjasamanya diucapkan terima kasih.")
    r_penutup.font.name = "Arial"

    ttd_table = doc.add_table(rows=1, cols=2)
    ttd_table.autofit = False 
    ttd_table.allow_autofit = False 
    for cell in ttd_table.columns[0].cells:
        cell.width = Inches(3.27) 
    for cell in ttd_table.columns[1].cells:
        cell.width = Inches(3.00) 
        
    cell_kiri = ttd_table.cell(0, 0)
    cell_kanan = ttd_table.cell(0, 1)
    
    p_kiri = cell_kiri.paragraphs[0]
    p_kiri.paragraph_format.line_spacing = 1.0
    r_kiri = p_kiri.add_run("\n\n\n\n\n\nKabid 3.1.1.........................")
    r_kiri.font.name = "Arial"
    r_kiri.font.size = Pt(12)
    
    p_kanan = cell_kanan.paragraphs[0]
    p_kanan.paragraph_format.line_spacing = 1.0
    p_kanan.alignment = WD_ALIGN_PARAGRAPH.LEFT 
    
    r_tgl = p_kanan.add_run(f"\nJakarta, {data['tgl_hari']} {data['tgl_bulan']} {data['tgl_tahun']}\n")
    r_tgl.font.name = "Arial"
    r_tgl.font.size = Pt(12)
    
    r_jab = p_kanan.add_run("Asisten Deputi Pembiayaan dan\nInvestasi Usaha Menengah\n\n")
    r_jab.font.name = "Arial"
    r_jab.font.size = Pt(12)
    
    r_nama = p_kanan.add_run("\n\n\n\nHengki Andrianto\n")
    r_nama.bold = True
    r_nama.font.name = "Arial"
    r_nama.font.size = Pt(12)
    
    r_nip = p_kanan.add_run("NIP. 19820512 200912 1 001")
    r_nip.font.name = "Arial"
    r_nip.font.size = Pt(12)

    p_tembusan = doc.add_paragraph()
    p_tembusan.paragraph_format.space_before = Pt(4) 
    p_tembusan.paragraph_format.line_spacing = 1.0   
    r_tembusan = p_tembusan.add_run(
        "\n\nTembusan Yth:\n"
        "1. Deputi Bidang Usaha Menengah;\n"
        "2. PPK Satuan Kerja Deputi Bidang Usaha Menengah;\n"
        "3. PPSPM Satuan Kerja Deputi Bidang Usaha Menengah;\n"
        "4. BP Satuan Kerja Deputi Bidang Usaha Menengah."
    )
    r_tembusan.font.name = "Arial"
    r_tembusan.font.size = Pt(12) 

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def generate_surat_tugas(data):
    doc = Document()
    set_document_font(doc, "Bookman Old Style", 12)

    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(13.0)
        section.top_margin = Inches(0.5)     
        section.bottom_margin = Inches(0.4)  
        section.left_margin = Inches(0.6)    
        section.right_margin = Inches(0.3)   
        setup_document_layout(section, "kop_surat.png", "watermark.png", "footer_surat.png")

    p_judul = doc.add_paragraph()
    p_judul.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_judul.paragraph_format.space_after = Pt(1)
    p_judul.paragraph_format.line_spacing = 1.5
    r_judul = p_judul.add_run("SURAT TUGAS")
    r_judul.bold = True
    r_judul.font.size = Pt(14)
    r_judul.font.name = "Bookman Old Style"

    p_nomor = doc.add_paragraph()
    p_nomor.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_nomor.paragraph_format.space_after = Pt(6)
    p_nomor.paragraph_format.line_spacing = 1.5
    r_nomor = p_nomor.add_run(f"NOMOR: {data['nomor_st']}")
    r_nomor.font.size = Pt(12)
    r_nomor.font.name = "Bookman Old Style"

    TAB_POS = Inches(2.50)  
    
    p_diperintah = doc.add_paragraph()
    p_diperintah.paragraph_format.space_after = Pt(2)
    p_diperintah.paragraph_format.line_spacing = 1.5
    r_diperintah = p_diperintah.add_run("I. DIPERINTAHKAN KEPADA")
    r_diperintah.bold = False
    r_diperintah.font.name = "Bookman Old Style"
    r_diperintah.font.size = Pt(12)

    pegawai_list_padded = list(data['pegawai_list'])
    while len(pegawai_list_padded) < 4:
        pegawai_list_padded.append({"nama": "", "jabatan": ""})

    INDENT_ANGKA = Inches(0.3)
    for idx, peg in enumerate(pegawai_list_padded[:4], start=1):
        p_nama = doc.add_paragraph()
        p_nama.paragraph_format.left_indent = TAB_POS
        p_nama.paragraph_format.first_line_indent = -TAB_POS + INDENT_ANGKA
        p_nama.paragraph_format.tab_stops.add_tab_stop(TAB_POS)
        p_nama.paragraph_format.space_after = Pt(1)
        p_nama.paragraph_format.line_spacing = 1.5
        
        r_nama = p_nama.add_run(f"{idx}.  Nama\t: {peg['nama']}")
        r_nama.font.name = "Bookman Old Style"
        r_nama.font.size = Pt(12)

        p_jab = doc.add_paragraph()
        p_jab.paragraph_format.left_indent = TAB_POS
        p_jab.paragraph_format.first_line_indent = -TAB_POS + INDENT_ANGKA
        p_jab.paragraph_format.tab_stops.add_tab_stop(TAB_POS)
        p_jab.paragraph_format.space_after = Pt(3)
        p_jab.paragraph_format.line_spacing = 1.5
        
        r_jab = p_jab.add_run(f"     Jabatan\t: {peg['jabatan']}")
        r_jab.font.name = "Bookman Old Style"
        r_jab.font.size = Pt(12)

    TAB_COLON = Inches(2.50)  
    TAB_TEXT = Inches(2.60)  

    p_ii = doc.add_paragraph()
    p_ii.paragraph_format.left_indent = TAB_TEXT
    p_ii.paragraph_format.first_line_indent = -TAB_TEXT
    p_ii.paragraph_format.tab_stops.add_tab_stop(TAB_COLON)
    p_ii.paragraph_format.tab_stops.add_tab_stop(TAB_TEXT)
    p_ii.paragraph_format.space_after = Pt(2)
    p_ii.paragraph_format.line_spacing = 1.5
    
    r_ii = p_ii.add_run(f"II. MAKSUD PERJALANAN\t:\t{data['maksud_kegiatan']}")
    r_ii.font.name = "Bookman Old Style"
    r_ii.font.size = Pt(12)
    r_ii.bold = False

    other_details = [
        ("III. TUJUAN", f": {data['wilayah_tujuan_st']}"),
        ("IV. JANGKA WAKTU", f": {data['jangka_waktu_st']} Hari"),
        ("     Tanggal Berangkat", f": {data['tgl_berangkat_st']}"),
        ("     Tanggal Kembali", f": {data['tgl_kembali_st']}"),
        ("V. KETERANGAN", ": 1. Setelah selesai melakukan perjalanan dinas ini segera\n      membuat laporan kepada kami.\n  2. SPT ini berlaku sejak tanggal dikeluarkan.")
    ]

    for label, val in other_details:
        p_det = doc.add_paragraph()
        p_det.paragraph_format.left_indent = TAB_POS
        p_det.paragraph_format.first_line_indent = -TAB_POS
        p_det.paragraph_format.tab_stops.add_tab_stop(TAB_POS)
        p_det.paragraph_format.space_after = Pt(2)
        p_det.paragraph_format.line_spacing = 1.5
        
        r_det = p_det.add_run(f"{label}\t{val}")
        r_det.font.name = "Bookman Old Style"
        r_det.font.size = Pt(12)
        r_det.bold = False

    if "Deputi Bidang Usaha Menengah" in data['tujuan_memo'] and "Sekretaris" not in data['tujuan_memo']:
        jabatan_penandatangan = "Deputi Bidang Usaha Menengah,"
        nama_penandatangan = "Bagus Rachman"
        nip_penandatangan = "NIP. 19710501 199903 1 001"
    else:
        jabatan_penandatangan = "Sekretaris Deputi Bidang Usaha Menengah,"
        nama_penandatangan = "Christina Agustin"
        nip_penandatangan = "NIP. 19720801 199803 2 001"

    sig_indent = Inches(4.2)  
    tab_sig = Inches(5.5)     

    p_tgl1 = doc.add_paragraph()
    p_tgl1.paragraph_format.left_indent = sig_indent
    p_tgl1.paragraph_format.tab_stops.add_tab_stop(tab_sig)
    p_tgl1.paragraph_format.space_before = Pt(10)
    p_tgl1.paragraph_format.space_after = Pt(1)
    p_tgl1.paragraph_format.line_spacing = 1.0
    r = p_tgl1.add_run("Dikeluarkan di\t: Jakarta")
    r.font.name = "Bookman Old Style"
    r.font.size = Pt(12)

    p_tgl2 = doc.add_paragraph()
    p_tgl2.paragraph_format.left_indent = sig_indent
    p_tgl2.paragraph_format.tab_stops.add_tab_stop(tab_sig)
    p_tgl2.paragraph_format.space_after = Pt(14) 
    p_tgl2.paragraph_format.line_spacing = 1.0
    r = p_tgl2.add_run(f"Pada Tanggal\t:            {data['st_bulan']} {data['st_tahun']}")
    r.font.name = "Bookman Old Style"
    r.font.size = Pt(12)

    p_jab_ttd = doc.add_paragraph()
    p_jab_ttd.paragraph_format.left_indent = sig_indent
    p_jab_ttd.paragraph_format.space_after = Pt(36) 
    p_jab_ttd.paragraph_format.line_spacing = 1.0
    r = p_jab_ttd.add_run(jabatan_penandatangan)
    r.font.name = "Bookman Old Style"
    r.font.size = Pt(12)

    p_nama = doc.add_paragraph()
    p_nama.paragraph_format.left_indent = sig_indent
    p_nama.paragraph_format.space_after = Pt(0)
    p_nama.paragraph_format.line_spacing = 1.0
    r = p_nama.add_run(f"\n\n{nama_penandatangan}")
    r.bold = True
    r.font.name = "Bookman Old Style"
    r.font.size = Pt(12)

    p_nip = doc.add_paragraph()
    p_nip.paragraph_format.left_indent = sig_indent
    p_nip.paragraph_format.space_after = Pt(0)
    p_nip.paragraph_format.line_spacing = 1.0
    r = p_nip.add_run(nip_penandatangan)
    r.font.name = "Bookman Old Style"
    r.font.size = Pt(12)

    p_temb_title = doc.add_paragraph()
    p_temb_title.paragraph_format.space_before = Pt(8)
    p_temb_title.paragraph_format.space_after = Pt(1)
    p_temb_title.paragraph_format.line_spacing = 1.0
    r = p_temb_title.add_run("\nTembusan:")
    r.font.name = "Bookman Old Style"
    r.font.size = Pt(12)

    tembusan_list = [
        "1. Pejabat Pembuat Komitmen;",
        "2. Pejabat Penguji dan Penandatangan SPM;",
        "3. Bendahara Pengeluaran; dan",
        "4. Arsip."
    ]
    for item in tembusan_list:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.space_after = Pt(0)
        p_item.paragraph_format.line_spacing = 1.0
        r = p_item.add_run(item)
        r.font.name = "Bookman Old Style"
        r.font.size = Pt(12)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- Fungsi SPPD ---
def format_tgl_indo(d):
    if d is None:
        return ""
    return f"{d.day:02d} {BULAN_INDO[d.month]} {d.year}"

def format_tanggal_indo_depan(date_obj):
    if date_obj is None:  
        return ""
    return f"{str(date_obj.day).zfill(2)} {BULAN_INDO[date_obj.month]} {date_obj.year}"

def update_mak_callback():
    selected = st.session_state.get("mk_d")
    if selected and selected in data_kegiatan:
        st.session_state["ak_m"] = f"150.01.EN.{data_kegiatan[selected]}"
    else:
        st.session_state["ak_m"] = ""

def reset_form():
    for key in list(st.session_state.keys()):
        # Pertahankan mode cetak dan navigasi utama
        if key not in ["mc_ui_b", "app_mode_nav"]: 
            del st.session_state[key]
    st.rerun()

def isi_sel_terformat(cell, text_content, line_spacing, mode_cetak):
    cell.text = ""
    lines = text_content.strip("\n").split("\n")
    for idx, line in enumerate(lines):
        p = cell.add_paragraph() if idx > 0 else cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = line_spacing
        
        is_bold, is_underline = False, False
        if "<b>" in line:
            is_bold = True
            line = line.replace("<b>", "")
        if "<u>" in line:
            is_underline = True
            line = line.replace("<u>", "")
        
        if "NIP." in line:
            p.paragraph_format.tab_stops.add_tab_stop(Cm(0.7), WD_TAB_ALIGNMENT.LEFT)
            p.paragraph_format.tab_stops.add_tab_stop(Cm(4.75), WD_TAB_ALIGNMENT.CENTER)
        else:
            p.paragraph_format.tab_stops.add_tab_stop(Cm(0.7), WD_TAB_ALIGNMENT.LEFT)
            p.paragraph_format.tab_stops.add_tab_stop(Cm(4.3), WD_TAB_ALIGNMENT.LEFT)
        
        if line.startswith("<c>") and line.endswith("</c>"):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            line = line[3:-4]
        elif line.startswith("<j>") and line.endswith("</j>"):
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            line = line[3:-4]
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
        parts = re.split(r'(<d>.*?</d>)', line)
        for part in parts:
            if not part: continue
            FONT_SIZE = Pt(11)
            run = p.add_run("")
            if part.startswith('<d>') and part.endswith('</d>'):
                val = part[3:-4]
                run.font.name = "Calibri"
                run.font.bold = is_bold            
                run.font.underline = is_underline  
                run.font.size = Pt(9.9) if len(val) > 35 else (Pt(10.5) if len(val) > 25 else FONT_SIZE)
                run.text = "" if mode_cetak == "Blangko Kosong" else val
                run.font.color.rgb = RGBColor(0, 0, 0)
            else:
                run.font.name = "Calibri"
                run.font.size = FONT_SIZE
                run.font.bold = is_bold            
                run.font.underline = is_underline  
                run.font.color.rgb = RGBColor(254, 254, 254) if mode_cetak == "OVERPRINT" else RGBColor(0, 0, 0)
                run.text = part

def generate_blanko_docx(data, mode_cetak):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Cm(0.6)
        section.bottom_margin = Cm(0.6)
        section.left_margin = Cm(1.0)
        section.right_margin = Cm(1.0)
        section.page_width = Cm(21.5)  
        section.page_height = Cm(33.0) 

    table = doc.add_table(rows=8, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    border_color = "000000" if mode_cetak != "OVERPRINT" else "FFFFFF"
    
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="6" w:color="{border_color}"/>\n'
        f'  <w:left w:val="single" w:sz="6" w:color="{border_color}"/>\n'
        f'  <w:bottom w:val="single" w:sz="6" w:color="{border_color}"/>\n'
        f'  <w:right w:val="single" w:sz="6" w:color="{border_color}"/>\n'
        f'  <w:insideH w:val="single" w:sz="4" w:color="{border_color}"/>\n'
        f'  <w:insideV w:val="single" w:sz="4" w:color="{border_color}"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)

    rows_data_templates = [
        ("", (
            "I.\tBerangkat dari\t: <d>{berangkat_1}</d>\n"
            "\t(tempat kedudukan)\n"
            "\tKe\t: <d>{ke_1}</d>\n"
            "\tPada Tanggal\t: <d>{tanggal_1}</d>\n"
            "<b><c><d>{jabatan_1}</d></c>\n"
            "<c></c>\n<c></c>\n<c></c>\n"
            "<b><u><c><d>{nama_1}</d></c>\n"
            "<b>\tNIP.\t<d>{nip_1}</d>"
        )),
        ((
            "II.\tTiba di\t: <d>{tiba_ii}</d>\n"
            "\tPada Tanggal\t: <d>{tgl_tiba_ii}</d>\n"
            "\tKepala\t:\n"
            "<c><d>{pejabat_ii}</d></c>\n"
            "<c></c>\n<c></c>\n<c></c>\n"
            "<b><c><d>{nama_ii}</d></c>\n"
            "<c>(.......................................................)</c>\n"
            "<b>\tNIP.\t<d>{nip_ii}</d>"
        ), (
            "\tBerangkat dari\t: <d>{berangkat_ii}</d>\n"
            "\tKe\t: <d>{tujuan_ii}</d>\n"
            "\tPada Tanggal\t: <d>{tgl_berangkat_ii}</d>\n"
            "<c><d>{pejabat_ii}</d></c>\n"
            "<c></c>\n<c></c>\n<c></c>\n"
            "<b><c><d>{nama_ii}</d></c>\n"
            "<c>(.......................................................)</c>\n"
            "<b>\tNIP.\t<d>{nip_ii}</d>"
        )),
        ((
            "III.\tTiba di\t: <d>{tiba_iii}</d>\n"
            "\tPada Tanggal\t: <d>{tgl_tiba_iii}</d>\n"
            "\tKepala\t:\n"
            "<c><d>{pejabat_iii}</d></c>\n"
            "<c></c>\n<c></c>\n<c></c>\n"
            "<b><c><d>{nama_iii}</d></c>\n"
            "<c>(.......................................................)</c>\n"
            "<b>\tNIP.\t<d>{nip_iii}</d>"
        ), (
            "\tBerangkat dari\t: <d>{berangkat_iii}</d>\n"
            "\tKe\t: <d>{tujuan_iii}</d>\n"
            "\tPada Tanggal\t: <d>{tgl_berangkat_iii}</d>\n"
            "<c><d>{pejabat_iii}</d></c>\n"
            "<c></c>\n<c></c>\n<c></c>\n"
            "<b><c><d>{nama_iii}</d></c>\n"
            "<c>(.......................................................)</c>\n"
            "<b>\tNIP.\t<d>{nip_iii}</d>"
        )),
        ((
            "IV.\tTiba di\t: <d>{tiba_iv}</d>\n"
            "\tPada Tanggal\t: <d>{tgl_tiba_iv}</d>\n"
            "\tKepala\t:\n"
            "<c><d>{pejabat_iv}</d></c>\n"
            "<c></c>\n<c></c>\n<c></c>\n"
            "<b><c><d>{nama_iv}</d></c>\n"
            "<c>(.......................................................)</c>\n"
            "<b>\tNIP.\t<d>{nip_iv}</d>"
        ), (
            "\tBerangkat dari\t: <d>{berangkat_iv}</d>\n"
            "\tKe\t: <d>{tujuan_iv}</d>\n"
            "\tPada Tanggal\t: <d>{tgl_berangkat_iv}</d>\n"
            "<c><d>{pejabat_iv}</d></c>\n"
            "<c></c>\n<c></c>\n<c></c>\n"
            "<b><c><d>{nama_iv}</d></c>\n"
            "<c>(.......................................................)</c>\n"
            "<b>\tNIP.\t<d>{nip_iv}</d>"
        )),
        ((
            "V.\tTiba di\t:\n"
            "\tPada Tanggal\t:\n"
            "\tKepala\t:\n"
            "<c></c>\n<c></c>\n<c></c>\n<c></c>\n"
            "<c>(.......................................................)</c>\n"
            "<b>\tNIP.\t"
        ), (
            "\tBerangkat dari\t:\n"
            "\tKe\t:\n"
            "\tPada Tanggal\t:\n"
            "<c></c>\n<c></c>\n<c></c>\n<c></c>\n"
            "<c>(.......................................................)</c>\n"
            "<b>\tNIP.\t"
        )),
        ((
            "VI.\tTiba di\t: <d>{tiba_vi}</d>\n"
            "\t(tempat kedudukan)\n"
            "\tPada Tanggal\t: <d>{tgl_vi}</d>\n\n"
            "<b><c><d>{jabatan_vi}</d></c>\n"
            "<c></c>\n<c></c>\n<c></c>\n"
            "<b><u><c><d>{nama_vi}</d></c>\n"
            "<b>\tNIP.\t<d>{nip_vi}</d>"
        ), (
            "<j>Telah diperiksa dengan keterangan bahwa perjalanan tersebut atas perintahnya "
            "dan semata-mata untuk kepentingan jabatan dalam waktu yang sesingkat-singkatnya.</j>\n"
            "<b><c><d>{jabatan_vi}</d></c>\n"
            "<c></c>\n<c></c>\n<c></c>\n"
            "<b><u><c><d>{nama_vi}</d></c>\n"
            "<b>\tNIP.\t<d>{nip_vi}</d>"
        )),
        ("VII.\tCatatan lain-lain\t:", None),
        (("<j>VIII. PERHATIAN :</j>\n"
          "<j>KPA/PPK yang menerbitkan SPD, pegawai yang melakukan perjalanan dinas, "
          "para pejabat yang mengesahkan tanggal berangkat/tiba, serta bendahara pengeluaran bertanggung jawab "
          "berdasarkan peraturan-peraturan Keuangan Negara, apabila negara menderita rugi akibat kesalahan, kelalaian dan kealpaannya.</j>"), None),
    ]

    for idx, row_template in enumerate(rows_data_templates):
        row = table.rows[idx]
        row.height = Cm(0.7) if idx == 6 else (Cm(1.8) if idx == 7 else Cm(3.7))
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        
        left_str = row_template[0].format(**data) if isinstance(row_template[0], str) else ""
        right_str = row_template[1].format(**data) if row_template[1] is not None else None

        if right_str is None:
            cell_0 = row.cells[0]
            cell_0.merge(row.cells[1])
            isi_sel_terformat(cell_0, left_str, 1.0, mode_cetak)
        else:
            isi_sel_terformat(row.cells[0], left_str, 1.0, mode_cetak)
            isi_sel_terformat(row.cells[1], right_str, 1.0, mode_cetak)

    for row in table.rows:
        row.cells[0].width = row.cells[1].width = Cm(10.11)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ==============================================================================
# 4. SISTEM NAVIGASI & SIDEBAR UTAMA
# ==============================================================================
with st.sidebar:
    st.markdown("### 🧭 Navigasi Modul")
    app_mode = st.radio(
        "Pilih Dokumen yang Ingin Dibuat:",
        ["Nota Dinas & Surat Tugas", "SPPD (Depan & Belakang)"],
        key="app_mode_nav"
    )
    
    # Render pengaturan khusus SPPD jika modul SPPD dipilih
    if app_mode == "SPPD (Depan & Belakang)":
        st.markdown("---")
        st.markdown("### ⚙️ Pengaturan Global SPPD")
        mode_cetak_ui = st.radio(
            "Pilih Mode Cetak Dokumen:",
            options=["Blangko Kosong", "OVERPRINT"],
            index=1,
            key="mc_ui_b",
            help="Gunakan 'OVERPRINT' jika Anda sudah menge-print Blangko Kosong."
        )
        
        st.markdown("---")
        st.markdown("### 💾 Manajemen Konfigurasi (Preset)")
        
        current_state = {k: v for k, v in st.session_state.items() if not k.startswith("FormSubmitter")}
        state_json = json.dumps(current_state, default=str, indent=4)
        st.download_button(
            label="📥 Simpan Preset (.json)",
            data=state_json,
            file_name="sppd_preset.json",
            mime="application/json",
            use_container_width=True
        )
        
        uploaded_preset = st.file_uploader("📂 Muat Preset (.json)", type=["json"])
        if uploaded_preset is not None:
            try:
                loaded_data = json.load(uploaded_preset)
                for k, v in loaded_data.items():
                    st.session_state[k] = v
                st.success("Preset berhasil dimuat!")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal memuat preset: {e}")

        st.markdown("---")
        st.button("🔄 Reset Seluruh Form", type="secondary", on_click=reset_form, use_container_width=True)

# ==============================================================================
# 5. TAMPILAN ANTARMUKA: MODUL 1 (NOTA DINAS & SURAT TUGAS)
# ==============================================================================
if app_mode == "Nota Dinas & Surat Tugas":
    st.header("📝 Pembuatan Nota Dinas & Surat Tugas Perjadin")

    col_setup1, col_setup2, col_setup3 = st.columns(3)
    with col_setup1:
        jumlah_pegawai = st.number_input(
            "👥 Jumlah pegawai Pelaksana Dinas", 
            min_value=1, 
            max_value=100, 
            value=1, 
            step=1
        )

    with col_setup2:
        selected_mak_full = st.selectbox(
            "💰 Pilih MAK Anggaran",
            [
                "7475.QDG.001.051.A.524111 (Koordinasi, Identifikasi, Persiapan Pelaksanaan Program Perluasan Pembiayaan KUR Klaster)",
                "7475.QDG.001.052.A.524111 (Koordinasi, Identifikasi, Persiapan Pelaksanaan Program Akselerasi Pembiayaan Inovatif bagi Usaha Menengah)"
            ]
        )

    with col_setup3:
        jenis_output = st.multiselect(
            "📄 Pilih Dokumen yang Ingin Dibuat",
            ["Nota Dinas", "Surat Tugas (ST)"],
            default=["Nota Dinas", "Surat Tugas (ST)"]
        )

    mak_anggaran = selected_mak_full.split(" (")[0]
    maksud_kegiatan = selected_mak_full.split(" (")[1].replace(")", "")

    DAFTAR_KENDARAAN = ["Pesawat", "Mobil", "BUS", "Kereta Api", "Travel", "Lainya (Bisa ketik manual)"]

    st.subheader("📌 1. Informasi Nomor Surat")
    col1, col2, col3 = st.columns(3)
    with col1:
        nomor_memo = st.text_input(
            "Nomor Nota Dinas",
            value="      /D.3.1.UMKM/TU.     /2026",
            placeholder="Contoh: 0000/D.3.1.UMKM/TU.00/2026",
        )
    with col2:
        nomor_st = st.text_input(
            "Nomor Surat Tugas (ST)",
            value="    /ST/  D.3.UMKM/TU.    /2026",
            placeholder="Contoh:     /ST/  D.3.UMKM/TU.    /2026",
        )
    with col3:
        tujuan_memo = st.selectbox(
            "Kepada Yth.",
            [
                "Deputi Bidang Usaha Menengah selaku Pejabat Penandatangan Surat Tugas (ST)",
                "Sekretaris Deputi Bidang Usaha Menengah selaku Pejabat Penandatangan Surat Tugas (ST)",
            ],
        )

    pengirim_memo = "Asisten Deputi Pembiayaan dan Investasi Usaha Menengah"

    st.markdown("---")
    st.subheader("👥 2. Data Pegawai Pelaksana Perjalanan Dinas")

    pegawai_data = []
    for i in range(jumlah_pegawai):
        st.markdown(f"**👤 Pegawai {i+1}**")
        
        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            pil_nama = st.selectbox(
                f"Pilih Nama Pegawai {i+1}", 
                list(DATA_PEGAWAI_JABATAN.keys()), 
                key=f"pil_nama_{i}"
            )
            
            if pil_nama == "Lainnya (Ketik Manual...)":
                nama_manual = st.text_input(f"Ketik Nama Pegawai {i+1}", placeholder="Nama Lengkap...", key=f"nama_man_{i}")
                jabatan_manual = st.text_input(f"Ketik Jabatan Pegawai {i+1}", placeholder="Jabatan Resmi...", key=f"jab_man_{i}")
            else:
                nama_manual = pil_nama
                jabatan_manual = DATA_PEGAWAI_JABATAN[pil_nama]

            gol = st.selectbox(f"Golongan {i+1}", ["III", "IV", "II", "-"], key=f"gol_{i}")
        
        with col_p2:
            prov = st.selectbox(f"Provinsi Tujuan {i+1}", list(DATA_WILAYAH.keys()), key=f"prov_{i}")
            
            opsi_kab = DATA_WILAYAH[prov] + ["Lainnya (Ketik Manual...)"]
            pil_kab = st.selectbox(f"Kabupaten/Kota Tujuan {i+1}", opsi_kab, key=f"pil_kab_{i}")
            
            if pil_kab == "Lainnya (Ketik Manual...)":
                kab_manual = st.text_input(f"Ketik Kabupaten/Kota {i+1}", placeholder="Nama Kab/Kota...", key=f"kab_man_{i}")
                kab_final = kab_manual
            else:
                kab_final = pil_kab
                
            today = datetime.date.today()
            tgl_input = st.date_input(
                f"Rentang Tanggal Dinas {i+1}",
                value=(today, today + datetime.timedelta(days=1)),
                key=f"tgl_{i}"
            )
            
            if isinstance(tgl_input, tuple):
                if len(tgl_input) == 2:
                    d1, d2 = tgl_input
                    jumlah_hari = (d2 - d1).days + 1
                    if d1.month == d2.month and d1.year == d2.year:
                        tgl_str = f"{d1.day} s.d. {d2.day} {BULAN_INDO[d1.month]} {d1.year}"
                    else:
                        tgl_str = f"{d1.day} {BULAN_INDO[d1.month]} s.d. {d2.day} {BULAN_INDO[d2.month]} {d2.year}"
                    tgl_berangkat_str = f"{d1.day} {BULAN_INDO[d1.month]} {d1.year}"
                    tgl_kembali_str = f"{d2.day} {BULAN_INDO[d2.month]} {d2.year}"
                elif len(tgl_input) == 1:
                    d1 = tgl_input[0]
                    jumlah_hari = 1
                    tgl_str = f"{d1.day} {BULAN_INDO[d1.month]} {d1.year}"
                    tgl_berangkat_str = tgl_str
                    tgl_kembali_str = tgl_str
                else:
                    jumlah_hari = 1
                    tgl_str = ""
                    tgl_berangkat_str = ""
                    tgl_kembali_str = ""
            else:
                d1 = tgl_input
                jumlah_hari = 1
                tgl_str = f"{d1.day} {BULAN_INDO[d1.month]} {d1.year}"
                tgl_berangkat_str = tgl_str
                tgl_kembali_str = tgl_str

        with col_p3:
            st.markdown(f"**Waktu (Hari) {i+1}**")
            st.info(f"⏱️ **{jumlah_hari} Hari**")
            
            pil_kendaraan = st.selectbox(f"Kendaraan {i+1}", DAFTAR_KENDARAAN, key=f"pil_kend_{i}")
            
            if pil_kendaraan == "Lainya (Bisa ketik manual)":
                kend_manual = st.text_input(f"Ketik Kendaraan {i+1}", placeholder="Nama Kendaraan...", key=f"kend_man_{i}")
                kend_final = kend_manual
            else:
                kend_final = pil_kendaraan
        
        pegawai_data.append({
            "nama": nama_manual,
            "jabatan": jabatan_manual,
            "gol": gol,
            "waktu_hari": jumlah_hari,
            "prov": prov,
            "kab_final": kab_final,
            "tgl_str": tgl_str,
            "tgl_berangkat": tgl_berangkat_str,
            "tgl_kembali": tgl_kembali_str,
            "kend_final": kend_final
        })
        
        if i < jumlah_pegawai - 1:
            st.divider()

    st.markdown("---")
    st.subheader("📅 3. Tanggal Surat")

    col_dt1, col_dt2, col_dt3 = st.columns(3)
    with col_dt1:
        tgl_surat_input = st.date_input("Tanggal Nota Dinas", value=datetime.date.today(), key="tgl_surat_dinas")
        tgl_hari = tgl_surat_input.day
        tgl_bulan = BULAN_INDO[tgl_surat_input.month]
        tgl_tahun = tgl_surat_input.year

    with col_dt2:
        st_bulan_pilih = st.selectbox(
            "Bulan Surat Tugas", 
            list(BULAN_INDO.values()), 
            index=datetime.date.today().month - 1, 
            key="st_bulan_pilih"
        )

    with col_dt3:
        st_tahun_pilih = st.number_input(
            "Tahun Surat Tugas", 
            min_value=2024, 
            max_value=2030, 
            value=datetime.date.today().year, 
            key="st_tahun_pilih"
        )

    submitted = st.button("🚀 PROSES DOKUMEN NOTA DINAS & ST")

    if submitted:
        if not jenis_output:
            st.error("⚠️ Mohon pilih minimal satu jenis dokumen yang ingin dibuat pada opsi di atas.")
        else:
            pegawai_list = []
            for idx, peg in enumerate(pegawai_data):
                if peg["nama"].strip(): 
                    pegawai_list.append(peg)

            wilayah_st = pegawai_data[0]["kab_final"] if len(pegawai_data) == 1 else pegawai_data[0]["prov"]

            data = {
                "nomor_memo": nomor_memo,
                "nomor_st": nomor_st,
                "tujuan_memo": tujuan_memo,
                "pengirim_memo": pengirim_memo,
                "maksud_kegiatan": maksud_kegiatan,
                "mak_anggaran": mak_anggaran,
                "pegawai_list": pegawai_list,
                "wilayah_tujuan_st": wilayah_st,
                "jangka_waktu_st": pegawai_data[0]["waktu_hari"],
                "tgl_berangkat_st": pegawai_data[0]["tgl_berangkat"],
                "tgl_kembali_st": pegawai_data[0]["tgl_kembali"],
                "tgl_hari": tgl_hari,
                "tgl_bulan": tgl_bulan,
                "tgl_tahun": tgl_tahun,
                "st_bulan": st_bulan_pilih,
                "st_tahun": st_tahun_pilih
            }

            st.success("🎉 Dokumen berhasil dibuat!")
            
            nama_kota_clean = pegawai_data[0]["kab_final"].replace(" ", "_").replace("/", "-")
            rentang_tgl_clean = f"{pegawai_data[0]['tgl_berangkat']}_s.d._{pegawai_data[0]['tgl_kembali']}".replace(" ", "_")

            files_to_download = {}
            if "Nota Dinas" in jenis_output:
                buffer_nd = generate_nota_dinas(data)
                files_to_download[f"Nota_Dinas_Perjadin_{nama_kota_clean}_{rentang_tgl_clean}.docx"] = buffer_nd.getvalue()
                    
            if "Surat Tugas (ST)" in jenis_output:
                buffer_st = generate_surat_tugas(data)
                files_to_download[f"Surat_Tugas_Perjadin_{nama_kota_clean}_{rentang_tgl_clean}.docx"] = buffer_st.getvalue()

            zip_filename = f"ND_dan_SPT_{nama_kota_clean}_{rentang_tgl_clean}.zip"

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for filename, file_bytes in files_to_download.items():
                    zip_file.writestr(filename, file_bytes)
            zip_buffer.seek(0)
            
            st.download_button(
                label="📥 Unduh Dokumen (.zip)",
                data=zip_buffer,
                file_name=zip_filename,
                mime="application/zip",
            )


# ==============================================================================
# 6. TAMPILAN ANTARMUKA: MODUL 2 (SPPD DEPAN & BELAKANG)
# ==============================================================================
elif app_mode == "SPPD (Depan & Belakang)":
    
    # Pengaturan mode_cetak diambil dari sidebar modul SPPD
    mode_cetak = "Blangko Kosong" if "Blangko Kosong" in st.session_state.get("mc_ui_b", "OVERPRINT") else "OVERPRINT"

    tab_depan, tab_belakang = st.tabs(["📄  SPPD Depan", "📑  SPPD Belakang"])

    # --- FORMULIR SPPD HALAMAN DEPAN ---
    with tab_depan:
        with st.container(border=True):
            st.markdown("### 📝 Informasi Surat & Pegawai")
            nomor_surat = st.text_input("Nomor SPPD", value="", placeholder="MM.DD.X/SPD/PPK/D.3.UMKM/20XX", key="ns_d")
            
            # --- MODIFIKASI: FITUR KETIK MANUAL ---
            opsi_pegawai_sppd = list(data_pegawai_sppd.keys()) + ["Lainnya (Ketik Manual...)"]
            nama_pegawai_pil = st.selectbox("Nama Pegawai", opsi_pegawai_sppd, index=None, placeholder="Nama Pegawai...", key="np_d_pil")
            
            daftar_pangkat = ["-", "I", "II", "III", "IV"]
            daftar_jabatan = sorted(list(set([data["jabatan"] for data in data_pegawai_sppd.values()])))
            
            if nama_pegawai_pil == "Lainnya (Ketik Manual...)":
                # Tampilkan Input Teks Manual
                nama_pegawai = st.text_input("Ketik Nama Pegawai Manual", placeholder="Masukkan Nama Lengkap...", key="np_man_d")
                
                col_pangkat, col_jabatan = st.columns([1, 2])
                with col_pangkat:
                    pangkat = st.selectbox("Pangkat/Golongan", daftar_pangkat, index=None, placeholder="Pilih Pangkat...", key="pg_man_d")
                with col_jabatan:
                    jabatan = st.text_input("Ketik Jabatan Manual", placeholder="Masukkan Jabatan Resmi...", key="jb_man_d")
            else:
                # Gunakan Variabel Normal
                nama_pegawai = nama_pegawai_pil
                
                col_pangkat, col_jabatan = st.columns([1, 2])
                with col_pangkat:
                    pangkat = st.selectbox("Pangkat/Golongan", daftar_pangkat, index=None, placeholder="Pilih Pangkat...", key="pg_d")
                with col_jabatan:
                    jabatan = st.selectbox("Jabatan", daftar_jabatan, index=None, placeholder="Pilih Jabatan...", key="jb_d")

        with st.container(border=True):
            st.markdown("### 🎯 Detail Perjalanan Dinas")
            maksud_perjalanan = st.selectbox(
                "Maksud Perjalanan Dinas (Dalam Rangka)", 
                list(data_kegiatan.keys()), 
                index=None, 
                placeholder="Pilih Maksud Perjalanan...", 
                key="mk_d",
                on_change=update_mak_callback
            )
            
            col_angkut, col_mak = st.columns(2)
            with col_angkut:
                alat_angkut = st.selectbox("Moda Transportasi yang Dipergunakan", daftar_angkutan, index=None, placeholder="Pilih Moda Transportasi...", key="ak_d")
            with col_mak:
                akun = st.text_input("Pembebanan Anggaran / Akun MAK", key="ak_m")

            col1, col2 = st.columns(2)
            with col1:
                tempat_berangkat = st.selectbox("Tempat Berangkat", OPI_KOTA_BLANK[1:], index=None, placeholder="Pilih Kota Asal...", key="as_d")
                tgl_berangkat_obj = st.date_input("Tanggal Berangkat", value=None, format="YYYY/MM/DD", key="tb_d")
            with col2:
                tempat_tujuan_list = st.multiselect("Tempat Tujuan (Bisa pilih lebih dari satu)", OPI_KOTA_BLANK[1:], default=[], placeholder="Pilih Kota Tujuan...", key="tj_d")
                tempat_tujuan = ", ".join(tempat_tujuan_list)
                tgl_kembali_obj = st.date_input("Tanggal Harus Kembali", value=None, format="YYYY/MM/DD", key="tk_d")

            if tgl_berangkat_obj and tgl_kembali_obj:
                selisih_hari = (tgl_kembali_obj - tgl_berangkat_obj).days + 1 
                if selisih_hari > 0:
                    lama_perjalanan = f"{selisih_hari} Hari"
                    st.info(f"⏱️ **Lama Perjalanan:** {lama_perjalanan}")
                else:
                    lama_perjalanan = ""
                    st.error("⚠️ Tanggal Kembali tidak boleh mendahului Tanggal Berangkat!")
            else:
                lama_perjalanan = ""
                
            tgl_dikeluarkan_obj = st.date_input("Tanggal Dikeluarkan", value=None, format="YYYY/MM/DD", key="to_d")

        st.markdown("")
        
        if not nama_pegawai or not maksud_perjalanan or not tgl_berangkat_obj or not tgl_kembali_obj or not tempat_tujuan_list:
            st.warning("⚠️ Harap lengkapi Nama Pegawai, Maksud Perjalanan, Tanggal, dan Tempat Tujuan untuk mengaktifkan unduhan SPPD Depan.")
        else:
            try:
                template_path = "SPPD DEPAN DRAFT.docx"
                doc = DocxTemplate(template_path)
                
                context = {
                    "nomor_surat": nomor_surat,
                    "nama_pegawai": nama_pegawai,
                    "pangkat": pangkat if pangkat else "", 
                    "jabatan": jabatan if jabatan else "",
                    "maksud_perjalanan": maksud_perjalanan,
                    "alat_angkut": alat_angkut if alat_angkut else "",
                    "akun": akun,
                    "tempat_berangkat": tempat_berangkat if tempat_berangkat else "",
                    "tempat_tujuan": tempat_tujuan,
                    "tanggal_berangkat": format_tanggal_indo_depan(tgl_berangkat_obj),
                    "tanggal_kembali": format_tanggal_indo_depan(tgl_kembali_obj),
                    "tanggal_dikeluarkan": format_tanggal_indo_depan(tgl_dikeluarkan_obj),
                    "lama_perjalanan": lama_perjalanan
                }
                
                doc.render(context)
                bio = io.BytesIO()
                doc.save(bio)
                
                st.download_button(
                    label="🚀 Unduh Dokumen SPPD Depan (.docx)",
                    data=bio.getvalue(),
                    file_name=f"SPPD_Depan_{nama_pegawai.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_d"
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan saat menyiapkan dokumen (pastikan file 'SPPD DEPAN DRAFT.docx' ada di folder yang sama): {e}")

    # --- FORMULIR SPPD HALAMAN BELAKANG ---
    with tab_belakang:
        st.markdown("")
        sub_tab1, sub_tab2, sub_tab3, sub_tab4, sub_tab5 = st.tabs(["Bagian I", "Bagian II", "Bagian III", "Bagian IV", "Bagian VI"])
        data_belakang = {}

        with sub_tab1:
            with st.container(border=True):
                st.subheader("I. Berangkat (Tempat Kedudukan)")
                filled_count = sum(1 for k in ['ke_1', 'tanggal_1'] if st.session_state.get(f"sel_{k}" if 'ke' in k else "dt_tgl_1"))
                st.progress(filled_count / 2, text=f"Progres Pengisian Bagian I: {filled_count}/2")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.text_input("Berangkat dari", "Jakarta", disabled=True, key="dis_berangkat_1")
                    data_belakang['berangkat_1'] = "Jakarta"
                    
                    selected_ke_1 = st.selectbox("Ke (Tujuan)", OPI_KOTA_BLANK, index=0, placeholder="Pilih Kota Tujuan...", key="sel_ke_1")
                    data_belakang['ke_1'] = "" if selected_ke_1 == "(Kosong / Belum diisi)" else selected_ke_1
                    
                    use_tgl_1 = st.checkbox("Isi Tanggal Berangkat I", value=False, key="chk_tgl_1")
                    raw_tanggal_1 = st.date_input("Pada Tanggal", value=None, format="YYYY/MM/DD", key="dt_tgl_1")
                    data_belakang['tanggal_1'] = format_tgl_indo(raw_tanggal_1) if (use_tgl_1 and raw_tanggal_1) else ""

                with col2:
                    data_belakang['jabatan_1'] = st.text_area("Jabatan Penandatangan", "ASISTEN DEPUTI PEMBIAYAAN DAN INVESTASI USAHA MENENGAH", key="ta_jab_1").replace("\n", "</d></c>\n<b><c><d>")
                    data_belakang['nama_1'] = st.text_input("Nama Pejabat", "HENGKI ANDRIANTO", key="ti_nama_1")
                    data_belakang['nip_1'] = st.text_input("NIP Pejabat", "19820512 200912 1 001", key="ti_nip_1")

        with sub_tab2:
            with st.container(border=True):
                st.subheader("II. Perjalanan Tujuan Pertama")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Sisi Kiri (Tiba)**")
                    selected_tiba_ii = st.selectbox("Tiba di (Kiri)", OPI_KOTA_BLANK, index=0, placeholder="Pilih Kota Tiba...", key="sel_tiba_ii")
                    data_belakang['tiba_ii'] = "" if selected_tiba_ii == "(Kosong / Belum diisi)" else selected_tiba_ii
                    
                    use_tgl_tiba_ii = st.checkbox("Isi Tanggal Tiba (Kiri)", value=False, key='chk_tgl_tiba_ii')
                    raw_tgl_tiba_ii = st.date_input("Tanggal Tiba", value=None, format="YYYY/MM/DD", key='dt_tgl_tiba_ii')
                    data_belakang['tgl_tiba_ii'] = format_tgl_indo(raw_tgl_tiba_ii) if (use_tgl_tiba_ii and raw_tgl_tiba_ii) else ""

                with col2:
                    st.markdown("**Sisi Kanan (Berangkat)**")
                    selected_berangkat_ii = st.selectbox("Berangkat dari (Kanan)", OPI_KOTA_BLANK, index=0, placeholder="Pilih Kota Berangkat...", key='sel_brgkt_ii')
                    data_belakang['berangkat_ii'] = "" if selected_berangkat_ii == "(Kosong / Belum diisi)" else selected_berangkat_ii
                    
                    selected_tujuan_ii = st.selectbox("Ke (Kanan)", OPI_KOTA_BLANK, index=0, placeholder="Pilih Kota Tujuan...", key='tujuan_ii_select')
                    data_belakang['tujuan_ii'] = "" if selected_tujuan_ii == "(Kosong / Belum diisi)" else selected_tujuan_ii
                    
                    use_tgl_berangkat_ii = st.checkbox("Isi Tanggal Berangkat (Kanan)", value=False, key='chk_tgl_brgkt_ii')
                    raw_tgl_berangkat_ii = st.date_input("Tanggal Berangkat", value=None, format="YYYY/MM/DD", key='dt_tgl_brgkt_ii')
                    data_belakang['tgl_berangkat_ii'] = format_tgl_indo(raw_tgl_berangkat_ii) if (use_tgl_berangkat_ii and raw_tgl_berangkat_ii) else ""

                st.markdown("---")
                st.markdown("**Pejabat Pengesah Bagian II**")
                col3, col4, col5 = st.columns(3)
                data_belakang['pejabat_ii'] = col3.text_input("Jabatan Pengesah", "", key="ti_pj_ii", placeholder="Masukkan Jabatan Pengesah...").replace("\n", "</d></c>\n<c><d>")
                data_belakang['nama_ii'] = col4.text_input("Nama Pengesah", "", key="ti_nm_ii", placeholder="Masukkan Nama Pengesah...")
                data_belakang['nip_ii'] = col5.text_input("NIP Pengesah", "", key="ti_np_ii", placeholder="Masukkan NIP Pengesah...")

        with sub_tab3:
            with st.container(border=True):
                st.subheader("III. Perjalanan Tujuan Kedua")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Sisi Kiri (Tiba)**")
                    selected_tiba_iii = st.selectbox("Tiba di (Kiri)", OPI_KOTA_BLANK, index=0, placeholder="Pilih Kota Tiba...", key="tiba_iii_select")
                    data_belakang['tiba_iii'] = "" if selected_tiba_iii == "(Kosong / Belum diisi)" else selected_tiba_iii
                    
                    use_tgl_tiba_iii = st.checkbox("Isi Tanggal Tiba (Kiri)", value=False, key="use_tgl_tiba_iii")
                    raw_tgl_tiba_iii = st.date_input("Tanggal Tiba", value=None, format="YYYY/MM/DD", key='dt_tgl_tiba_iii')
                    data_belakang['tgl_tiba_iii'] = format_tgl_indo(raw_tgl_tiba_iii) if (use_tgl_tiba_iii and raw_tgl_tiba_iii) else ""

                with col2:
                    st.markdown("**Sisi Kanan (Berangkat)**")
                    selected_berangkat_iii = st.selectbox("Berangkat dari (Kanan)", OPI_KOTA_BLANK, index=0, placeholder="Pilih Kota Berangkat...", key="berangkat_iii_select")
                    data_belakang['berangkat_iii'] = "" if selected_berangkat_iii == "(Kosong / Belum diisi)" else selected_berangkat_iii
                    
                    selected_tujuan_iii = st.selectbox("Ke (Kanan)", OPI_KOTA_BLANK, index=0, placeholder="Pilih Kota Tujuan...", key="tujuan_iii_select")
                    data_belakang['tujuan_iii'] = "" if selected_tujuan_iii == "(Kosong / Belum diisi)" else selected_tujuan_iii
                    
                    use_tgl_berangkat_iii = st.checkbox("Isi Tanggal Berangkat (Kanan)", value=False, key="use_tgl_brgkt_iii")
                    raw_tgl_berangkat_iii = st.date_input("Tanggal Berangkat", value=None, format="YYYY/MM/DD", key='dt_tgl_brgkt_iii')
                    data_belakang['tgl_berangkat_iii'] = format_tgl_indo(raw_tgl_berangkat_iii) if (use_tgl_berangkat_iii and raw_tgl_berangkat_iii) else ""

                st.markdown("---")
                st.markdown("**Pejabat Pengesah Bagian III**")
                col3, col4, col5 = st.columns(3)
                data_belakang['pejabat_iii'] = col3.text_input("Jabatan Pengesah", "", key="pejabat_iii", placeholder="Masukkan Jabatan Pengesah...").replace("\n", "</d></c>\n<c><d>")
                data_belakang['nama_iii'] = col4.text_input("Nama Pengesah", "", key="nama_iii", placeholder="Masukkan Nama Pengesah...")
                data_belakang['nip_iii'] = col5.text_input("NIP Pengesah", "", key="nip_iii", placeholder="Masukkan NIP Pengesah...")

        with sub_tab4:
            with st.container(border=True):
                st.subheader("IV. Perjalanan Tujuan Ketiga")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**Sisi Kiri (Tiba)**")
                    selected_tiba_iv = st.selectbox("Tiba di (Kiri)", OPI_KOTA_BLANK, index=0, placeholder="Pilih Kota Tiba...", key="tiba_iv_select")
                    data_belakang['tiba_iv'] = "" if selected_tiba_iv == "(Kosong / Belum diisi)" else selected_tiba_iv
                    
                    use_tgl_tiba_iv = st.checkbox("Isi Tanggal Tiba (Kiri)", value=False, key="use_tgl_tiba_iv")
                    raw_tgl_tiba_iv = st.date_input("Tanggal Tiba", value=None, format="YYYY/MM/DD", key='dt_tgl_tiba_iv')
                    data_belakang['tgl_tiba_iv'] = format_tgl_indo(raw_tgl_tiba_iv) if (use_tgl_tiba_iv and raw_tgl_tiba_iv) else ""

                with col2:
                    st.markdown("**Sisi Kanan (Berangkat)**")
                    selected_berangkat_iv = st.selectbox("Berangkat dari (Kanan)", OPI_KOTA_BLANK, index=0, placeholder="Pilih Kota Berangkat...", key="berangkat_iv_select")
                    data_belakang['berangkat_iv'] = "" if selected_berangkat_iv == "(Kosong / Belum diisi)" else selected_berangkat_iv
                    
                    selected_tujuan_iv = st.selectbox("Ke (Kanan)", OPI_KOTA_BLANK, index=0, placeholder="Pilih Kota Tujuan...", key="tujuan_iv_select")
                    data_belakang['tujuan_iv'] = "" if selected_tujuan_iv == "(Kosong / Belum diisi)" else selected_tujuan_iv
                    
                    use_tgl_berangkat_iv = st.checkbox("Isi Tanggal Berangkat (Kanan)", value=False, key="use_tgl_brgkt_iv")
                    raw_tgl_berangkat_iv = st.date_input("Tanggal Berangkat", value=None, format="YYYY/MM/DD", key='dt_tgl_brgkt_iv')
                    data_belakang['tgl_berangkat_iv'] = format_tgl_indo(raw_tgl_berangkat_iv) if (use_tgl_berangkat_iv and raw_tgl_berangkat_iv) else ""

                st.markdown("---")
                st.markdown("**Pejabat Pengesah Bagian IV**")
                col3, col4, col5 = st.columns(3)
                data_belakang['pejabat_iv'] = col3.text_input("Jabatan Pengesah", "", key="pejabat_iv", placeholder="Masukkan Jabatan Pengesah...").replace("\n", "</d></c>\n<c><d>")
                data_belakang['nama_iv'] = col4.text_input("Nama Pengesah", "", key="nama_iv", placeholder="Masukkan Nama Pengesah...")
                data_belakang['nip_iv'] = col5.text_input("NIP Pengesah", "", key="nip_iv", placeholder="Masukkan NIP Pengesah...")

        with sub_tab5:
            with st.container(border=True):
                st.subheader("VI. Tiba Kembali")
                col1, col2 = st.columns(2)
                with col1:
                    data_belakang['tiba_vi'] = "Jakarta"
                    st.text_input("Tiba di", "Jakarta", disabled=True, key="dis_tiba_vi")
                    use_tgl_vi = st.checkbox("Isi Tanggal Kembali (VI)", value=False, key="chk_tgl_vi")
                    raw_tgl_vi = st.date_input("Pada Tanggal (Kembali)", value=None, format="YYYY/MM/DD", key='dt_tgl_kembali_vi')
                    data_belakang['tgl_vi'] = format_tgl_indo(raw_tgl_vi) if (use_tgl_vi and raw_tgl_vi) else ""

                with col2:
                    data_belakang['jabatan_vi'] = st.text_area("Jabatan Penandatangan VI", "DEPUTI BIDANG USAHA MENENGAH\nPejabat Pembuat Komitmen", key="ta_jab_vi").replace("\n", "</d></c>\n<b><c><d>")
                    data_belakang['nama_vi'] = st.text_input("Nama Penandatangan VI", "Christina Agustin", key="ti_nama_vi")
                    data_belakang['nip_vi'] = st.text_input("NIP Penandatangan VI", "19890520 201403 1 001", key="ti_nip_vi")

        st.markdown("")
        try:
            docx_file_belakang = generate_blanko_docx(data_belakang, mode_cetak)
            file_name_suffix = {
                "Blangko Kosong": "Blangko_Kosong",
                "OVERPRINT": "OVERPRINT"
            }[mode_cetak]
            
            st.download_button(
                label=f"🚀 Unduh Dokumen SPPD Belakang ({file_name_suffix}) (.docx)",
                data=docx_file_belakang,
                file_name=f"SPPD_Belakang_{file_name_suffix}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
                key="dl_btn_belakang"
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan saat membuat dokumen belakang: {e}")