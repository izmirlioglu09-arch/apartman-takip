import streamlit as st
import pandas as pd
import datetime
import os
import io

# PDF Altyapısı (ReportLab)
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Veritabanı Dosya İsimleri
AIDAT_FILE = "aidat_veritabanı.xlsx"
MASRAF_FILE = "masraf_veritabanı.xlsx"
SAKIN_AYAR_FILE = "sakin_ayarları.xlsx"

# Bulut Sunucu Standart Fontları
FONT_N = 'Helvetica'
FONT_B = 'Helvetica-Bold'

def sakin_ayarlarini_yukle():
    if os.path.exists(SAKIN_AYAR_FILE): 
        return pd.read_excel(SAKIN_AYAR_FILE)
    df = pd.DataFrame([{"Daire No": f"Daire {i}", "Sakin Adı": "Boş / Tanımsız", "Aylık Aidat (TL)": 500.0} for i in range(1, 11)])
    df.to_excel(SAKIN_AYAR_FILE, index=False)
    return df

def aidat_veritabani_yukle(df_sakinler):
    aylar = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
    if os.path.exists(AIDAT_FILE):
        try:
            df_aidat = pd.read_excel(AIDAT_FILE)
            guncel = []
            for _, row in df_sakinler.iterrows():
                g_adi = f"{row['Daire No']} - {row['Sakin Adı']}"
                for ay in aylar:
                    eski = df_aidat[(df_aidat["Daire No"].str.startswith(f"{row['Daire No']} -")) & (df_aidat["Dönem Ay"] == ay)]
                    v_odenen = 0.0
                    v_tarih = "-"
                    v_durum = "Ödenmedi"
                    if not eski.empty:
                        v_odenen = float(eski["Ödenen Tutar (TL)"].sum())
                        v_tarih = str(eski["Ödeme Tarihi"].iloc)
                        v_durum = str(eski["Durum"].iloc)
                    guncel.append({
                        "Daire No": g_adi, "Dönem Ay": ay, "Aidat Tutarı (TL)": float(row["Aylık Aidat (TL)"]),
                        "Ödenen Tutar (TL)": v_odenen, "Ödeme Tarihi": v_tarih, "Durum": v_durum
                    })
            df = pd.DataFrame(guncel)
            df.to_excel(AIDAT_FILE, index=False)
            return df
        except: 
            pass
    
    veriler = []
    for _, r in df_sakinler.iterrows():
        for ay in aylar:
            veriler.append({
                "Daire No": f"{r['Daire No']} - {r['Sakin Adı']}", "Dönem Ay": ay, "Aidat Tutarı (TL)": float(r["Aylık Aidat (TL)"]), "Ödenen Tutar (TL)": 0.0, "Ödeme Tarihi": "-", "Durum": "Ödenmedi"
            })
    df = pd.DataFrame(veriler)
    df.to_excel(AIDAT_FILE, index=False)
    return df

def masraf_veritabani_yukle():
    if os.path.exists(MASRAF_FILE): 
        return pd.read_excel(MASRAF_FILE)
    df = pd.DataFrame(columns=["Tarih", "Cari Adı / Kurum", "Kategori", "Açıklama", "Tutar (TL)"])
    df.to_excel(MASRAF_FILE, index=False)
    return df

# PDF Üretme Modülü
def pdf_ekstre_uret(cari_isim, b_tar, bit_tar, df_veri, t_giriis, t_cikis):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    
    s_baslik = ParagraphStyle('B', fontName=FONT_B, fontSize=16, spaceAfter=15, alignment=1)
    s_normal = ParagraphStyle('N', fontName=FONT_N, fontSize=9)
    s_t_baslik = ParagraphStyle('TB', fontName=FONT_B, fontSize=9, textColor=colors.whitesmoke)
    
    story.append(Paragraph("APARTMAN YONETIMI MALI CARI EKSTRE RAPORU", s_baslik))
    story.append(Paragraph(f"Cari Hesap: {cari_isim} | Donem: {b_tar.strftime('%d.%m.%Y')} - {bit_tar.strftime('%d.%m.%Y')}", s_normal))
    story.append(Spacer(1, 15))
    
    t_data = [[Paragraph("Tarih", s_t_baslik), Paragraph("Islem", s_t_baslik), Paragraph("Cari Hesap / Kisi / Kurum", s_t_baslik), Paragraph("Gelir (TL)", s_t_baslik), Paragraph("Gider (TL)", s_t_baslik)]]
    
    for _, r in df_veri.iterrows():
        t_data.append([
            Paragraph(str(r["Tarih"]), s_normal), Paragraph(str(r["Islem"]), s_normal),
            Paragraph(str(r["Cari Hesap / Kisi / Kurum"]), s_normal), Paragraph(f"{r['Gelir']:,.2f} TL", s_normal), Paragraph(f"{r['Gider']:,.2f} TL", s_normal)
        ])
        
    t_data.append([Paragraph("TOPLAM", s_normal), Paragraph("", s_normal), Paragraph("", s_normal), Paragraph(f"{t_giriis:,.2f} TL", s_normal), Paragraph(f"{t_cikis:,.2f} TL", s_normal)])
    t_data.append([Paragraph("NET DURUM (BAKIYE)", s_normal), Paragraph("", s_normal), Paragraph("", s_normal), Paragraph(f"{(t_giriis - t_cikis):,.2f} TL", s_normal), Paragraph("", s_normal)])
    
    tablo = Table(t_data, colWidths=[65, 95, 160, 110, 110])
    tablo.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2c3e50')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('PADDING', (0,0), (-1,-1), 5),
        ('BACKGROUND', (0,-2), (-1,-1), colors.HexColor('#f8f9fa'))
    ]))
    story.append(tablo)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

st.set_page_config(page_title="Cari Apartman Takip", layout="wide", page_icon="🏢")
st.title("🏢 Kurumsal Cari Tabanlı Apartman Yönetim Sistemi")

df_sakin_ayarlari = sakin_ayarlarini_yukle()
df_aidat = aidat_veritabani_yukle(df_sakin_ayarlari)
df_masraf = masraf_veritabani_yukle()

# Sekme tanımlamaları arasındaki çakışma yaratan eski kod satırı tamamen silindi
m_s1, m_s2, m_s3, m_s4, m_s5 = st.tabs([
    "⚙️ 1. Sakin & Aidat Tanımlama", 
    "💰 2. Cari Aidat Tahsilatı", 
    "📉 3. Cari Masraf & Kalem Girişi", 
    "🔍 4. Tarih & Cari Raporlama Paneli",
    "📊 5. Aylık Mali Özet"
])

# --- MODÜL 1: SAKİN VE SABİT AİDAT TANIMLAMA ---
with m_s1:
    st.header("⚙️ Bina Sakinlerini ve Aidat Miktarlarını Tanımla")
    c1, c2, c3 = st.columns(3)
    ay_d_no = c1.text_input("Daire Numarası (Örn: Daire 1)", key="t_d")
    ay_s_ad = c2.text_input("Sakin Adı Soyadı (Cari Ünvan)", key="t_s")
    ay_a_tut = c3.number_input("Aylık Sabit Aidat Borcu (TL)", min_value=0.0, value=500.0, step=50.0, key="t_a")
    
    btn_sakin_tetik = st.button("🔄 Sakin Bilgisini ve Aidatını Sisteme İşle", key="t_btn", type="primary")
    if btn_sakin_tetik and ay_d_no and ay_s_ad:
        if ay_d_no in df_sakin_ayarlari["Daire No"].values:
            idx = df_sakin_ayarlari[df_sakin_ayarlari["Daire No"] == ay_d_no].index
            df_sakin_ayarlari.loc[idx, "Sakin Adı"] = ay_s_ad
            df_sakin_ayarlari.loc[idx, "Aylık Aidat (TL)"] = ay_a_tut
        else:
            df_sakin_ayarlari = pd.concat([df_sakin_ayarlari, pd.DataFrame([{"Daire No": ay_d_no, "Sakin Adı": ay_s_ad, "Aylık Aidat (TL)": ay_a_tut}])], ignore_index=True)
        df_sakin_ayarlari.to_excel(SAKIN_AYAR_FILE, index=False)
        st.success("✅ Sakin ve aidat şablonu başarıyla güncellendi!")
        st.rerun()
    st.dataframe(df_sakin_ayarlari, use_container_width=True)

# --- MODÜL 2: CARİ BAZLI AİDAT TAHSİLAT PANELİ ---
with m_s2:
    st.header("💰 Cari Hesap Aidat Tahsilat Girişi")
    col1, col2, col3, col4, col5 = st.columns(5)
    sec_daire = col1.selectbox("Tahsilat Yapılan Sakin / Cari", df_aidat["Daire No"].unique(), key="c_d")
    sec_ay = col2.selectbox("Aidat Dönem Ayı", df_aidat["Dönem Ay"].unique(), key="c_a")
    
    filt = df_aidat[(df_aidat["Daire No"] == sec_daire) & (df_aidat["Dönem Ay"] == sec_ay)]
    v_tutar = 500.0
    if not filt.empty:
        v_tutar = float(filt["Aidat Tutarı (TL)"].sum())
        
    od_tutar = col3.number_input("Tahsil Edilen Tutar (TL)", min_value=0.0, value=v_tutar, key="c_t")
    od_durum = col4.selectbox("Ödeme Durumu", ["Ödendi", "Ödenmedi"], key="c_dur")
    sec_tarih = col5.date_input("Kasa İşlem Tarihi", value=datetime.date.today(), format="DD.MM.YYYY", key="c_tar")
    
    b1, b2 = st.columns(2)
    btn_kaydet_tetik = b1.button("💾 Tahsilatı Cari Kasa Defterine İşle", type="primary", key="c_kaydet")
    btn_sil_tetik = b2.button("❌ Seçili Ayın Ödemesini İptal Et (Sil)", key="c_sil")
    
    if btn_kaydet_tetik:
        idx = df_aidat[(df_aidat["Daire No"] == sec_daire) & (df_aidat["Dönem Ay"] == sec_ay)].index
        df_aidat.loc[idx, "Ödenen Tutar (TL)"] = od_tutar
        df_aidat.loc[idx, "Durum"] = od_durum
        df_aidat.loc[idx, "Ödeme Tarihi"] = sec_tarih.strftime("%d.%m.%Y") if od_durum == "Ödendi" else "-"
        df_aidat.to_excel(AIDAT_FILE, index=False)
        st.success("✅ Aidat tahsilatı ilgili carinin hesabına işlendi!")
        st.rerun()
        
    if btn_sil_tetik:
        idx = df_aidat[(df_aidat["Daire No"] == sec_daire) & (df_aidat["Dönem Ay"] == sec_ay)].index
        df_aidat.loc[idx, "Ödenen Tutar (TL)"] = 0.0
        df_aidat.loc[idx, "Durum"] = "Ödenmedi"
        df_aidat.loc[idx, "Ödeme Tarihi"] = "-"
        df_aidat.to_excel(AIDAT_FILE, index=False)
        st.warning("🗑️ Ödeme iptal edildi, cari borçlandırıldı.")
        st.rerun()

# --- MODÜL 3: CARİ BAZLI MASRAF VE KALEM GİRİŞİ ---
with m_s3:
    st.header("📉 Cari Firma / Kurum Gider Girişi")
    mg1, mg2, mg3, mg4, mg5 = st.columns(5)
    mas_cari = mg1.text_input("Ödeme Yapılan Firma / Kurum (Cari Adı)", placeholder="Örn: Aydem Elektrik", key="g_c")
    mas_kat = mg2.selectbox("Gider Kalemi / Kategorisi", ["Elektrik Faturası", "Su Faturası", "Asansör Bakımı", "Temizlik Görevlisi", "Tamirat / Tadilat", "Yönetim Masrafı", "Diğer"], key="g_k")
    mas_acik = mg3.text_input("Fatura / Harcama Açıklaması", placeholder="Örn: Ortak alan sayacı", key="g_a")
    mas_tutar = mg4.number_input("Fatura Tutarı (TL)", min_value=0.0, step=50.0, key="g_t")
    mas_tarih = mg5.date_input("Fatura Ödeme Tarihi", value=datetime.date.today(), format="DD.MM.YYYY", key="g_tar")
    
    btn_masraf_tetik = st.button("🔴 Gider Kalemini Kasaya İşle", type="primary", key="g_btn")
    if btn_masraf_tetik and mas_cari and mas_tutar > 0:
        yeni_m = pd.DataFrame([{"Tarih": mas_tarih.strftime("%d.%m.%Y"), "Cari Adı / Kurum": mas_cari, "Kategori": mas_kat, "Açıklama": mas_acik, "Tutar (TL)": mas_tutar}])
