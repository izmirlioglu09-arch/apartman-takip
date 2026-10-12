import streamlit as st
import pandas as pd
import datetime
import os

# Veritabanı Dosya İsimleri
AIDAT_FILE = "aidat_veritabanı.xlsx"
MASRAF_FILE = "masraf_veritabanı.xlsx"
SAKIN_AYAR_FILE = "sakin_ayarları.xlsx"

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
            for idx_s, row in df_sakinler.iterrows():
                g_adi = f"{row['Daire No']} - {row['Sakin Adı']}"
                for ay in aylar:
                    eski = df_aidat[(df_aidat["Daire No"].str.startswith(f"{row['Daire No']} -")) & (df_aidat["Dönem Ay"] == ay)]
                    v_odenen = 0.0
                    v_tarih = "-"
                    v_durum = "Ödenmedi"
                    if not list(eski.index) == []:
                        v_odenen = float(eski["Ödenen Tutar (TL)"].iloc)
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
    for idx_s, r in df_sakinler.iterrows():
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
    df = pd.DataFrame(columns=["Tarih", "Kategori", "Açıklama", "Tutar (TL)"])
    df.to_excel(MASRAF_FILE, index=False)
    return df

# Sayfa Yapılandırması
st.set_page_config(page_title="Apartman Yönetimi", layout="wide", page_icon="🏢")
st.title("🏢 Profesyonel Apartman Finans Takip Sistemi")

df_sakin_ayarlari = sakin_ayarlarini_yukle()
df_aidat = aidat_veritabani_yukle(df_sakin_ayarlari)
df_masraf = masraf_veritabani_yukle()

# =================================================================
# 📋 BÖLÜM 1: AİDAT TAHSİLATI ALMA VE İPTAL ETME (SİLME) ALANI
# =================================================================
st.markdown("### 💰 1. Yeni Aidat Ödemesi Al / Düzenle / İptal Et")
col1, col2, col3, col4, col5 = st.columns(5)
with col1: secilen_daire = st.selectbox("Daire / Sakin Seçin", df_aidat["Daire No"].unique(), key="s1_d")
with col2: secilen_ay = st.selectbox("Dönem Ayı Seçin", df_aidat["Dönem Ay"].unique(), key="s1_a")
with col3:
    filt = df_aidat[(df_aidat["Daire No"] == secilen_daire) & (df_aidat["Dönem Ay"] == secilen_ay)]
    v_tutar = 500.0
    if not list(filt.index) == []:
        v_tutar = float(filt["Aidat Tutarı (TL)"].iloc)
    od_tutar = st.number_input("Ödenen Tutar (TL)", min_value=0.0, value=v_tutar, key="s1_t")
with col4: od_durum = st.selectbox("Durum", ["Ödendi", "Ödenmedi"], key="s1_dur")
with col5: sec_tarih = st.date_input("Tahsilat Tarihi", value=datetime.date.today(), format="DD.MM.YYYY", key="s1_tar")
    
c_btn1, c_btn2 = st.columns(2)
with c_btn1:
    if st.button("💾 Kaydet", type="primary", key="s1_btn"):
        idx = df_aidat[(df_aidat["Daire No"] == secilen_daire) & (df_aidat["Dönem Ay"] == secilen_ay)].index
        df_aidat.loc[idx, "Ödenen Tutar (TL)"] = od_tutar
        df_aidat.loc[idx, "Durum"] = od_durum
        df_aidat.loc[idx, "Ödeme Tarihi"] = sec_tarih.strftime("%d.%m.%Y") if od_durum == "Ödendi" else "-"
        df_aidat.to_excel(AIDAT_FILE, index=False)
        st.success("✅ Aidat kaydı güncellendi!")
        st.rerun()
with c_btn2:
    if st.button("❌ Bu Ayın Aidat Ödemesini Sil (Sıfırla)", type="secondary", key="s1_sil_btn"):
        idx = df_aidat[(df_aidat["Daire No"] == secilen_daire) & (df_aidat["Dönem Ay"] == secilen_ay)].index
        df_aidat.loc[idx, "Ödenen Tutar (TL)"] = 0.0
        df_aidat.loc[idx, "Durum"] = "Ödenmedi"
        df_aidat.loc[idx, "Ödeme Tarihi"] = "-"
        df_aidat.to_excel(AIDAT_FILE, index=False)
        st.warning("🗑️ Aidat ödeme kaydı silindi ve borç ödenmedi durumuna getirildi!")
        st.rerun()

st.markdown("---")

# =================================================================
# 📋 BÖLÜM 2: BINA MASRAFI EKLEME VE SEÇİLENİ SİLME ALANI
# =================================================================
st.markdown("### 📉 2. Yeni Bina Masrafı Ekle / Kayıtlı Masrafı Sil")
m1, m2, m3, m4 = st.columns(4)
with m1: m_kat = st.selectbox("Masraf Kategorisi", ["Elektrik Faturası", "Su Faturası", "Asansör Bakımı", "Temizlik Görevlisi", "Tamirat / Tadilat", "Diğer"], key="s2_k")
with m2: m_acik = st.text_input("Masraf Açıklaması", key="s2_a")
with m3: m_tutar = st.number_input("Masraf Tutarı (TL)", min_value=0.0, step=50.0, key="s2_t")
with m4: m_tar = st.date_input("Masraf Tarihi", value=datetime.date.today(), format="DD.MM.YYYY", key="s2_tar")
    
if st.button("🔴 Masrafı Kasaya İşle", key="s2_btn"):
    yeni = pd.DataFrame([{"Tarih": m_tar.strftime("%d.%m.%Y"), "Kategori": m_kat, "Açıklama": m_acik, "Tutar (TL)": m_tutar}])
    df_masraf = pd.concat([df_masraf, yeni], ignore_index=True)
    df_masraf.to_excel(MASRAF_FILE, index=False)
    st.error("📉 Gider kasaya işlendi.")
    st.rerun()

if not df_masraf.empty:
    st.markdown("##### 🗑️ Giriş Yapılmış Masrafı Listeden Sil")
    df_masraf_gosterim = df_masraf.copy()
    df_masraf_gosterim["Listele"] = "Tarih: " + df_masraf_gosterim["Tarih"].astype(str) + " | " + df_masraf_gosterim["Kategori"].astype(str) + " | " + df_masraf_gosterim["Açıklama"].astype(str) + " (" + df_masraf_gosterim["Tutar (TL)"].astype(str) + " TL)"
    secilen_silinecek_masraf = st.selectbox("Silmek İstediğiniz Masraf Kaydını Seçin", df_masraf_gosterim["Listele"].unique(), key="m_sil_sec")
    if st.button("🗑️ Seçilen Masrafı Tamamen Sil", type="secondary", key="m_sil_btn"):
        secili_indeks = df_masraf_gosterim[df_masraf_gosterim["Listele"] == secilen_silinecek_masraf].index
        df_masraf = df_masraf.drop(secili_indeks).reset_index(drop=True)
        df_masraf.to_excel(MASRAF_FILE, index=False)
        st.success("✅ Seçilen masraf kaydı başarıyla silindi!")
        st.rerun()

st.markdown("---")

# =================================================================
# 📋 BÖLÜM 3: AYLIK MALİ DURUM RAPORU
# =================================================================
st.markdown("### 📊 3. Aylık Finansal Özet Raporları")
rep_ay = st.selectbox("Raporunu Görmek İstediğiniz Ayı Seçin", ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"], key="s3_ay")
f_gelir = df_aidat[(df_aidat["Dönem Ay"] == rep_ay) & (df_aidat["Durum"].str.contains("Ödendi", na=False))]
t_gelir = f_gelir["Ödenen Tutar (TL)"].sum()
t_gider = df_masraf["Tutar (TL)"].sum() if not df_masraf.empty else 0.0
t_net = t_gelir - t_gider
c1, c2, c3 = st.columns(3)
c1.metric("🟢 Toplam Gelir", f"{t_gelir:,.2f} TL")
c2.metric("🔴 Toplam Gider", f"{t_gider:,.2f} TL")
c3.metric("⚖️ Net Kasa", f"{t_net:,.2f} TL")

st.markdown("---")

# =================================================================
# 📋 BÖLÜM 4: FINANSAL CARI EKSTRE LISTESI (TÜM GİRİNTİLER TEMİZLENDİ)
# =================================================================
st.markdown("### 🔍 4. Finansal Cari Ekstre Listesi")
ex_col1, ex_col2 = st.columns(2)
with ex_col1:
    cari_liste = ["Tümü (Tüm Bina / Gelir+Gider)"] + list(df_aidat["Daire No"].unique())
    secilen_cari = st.selectbox("Cari Hesap Seçin", cari_liste, key="ex_cari")
with ex_col2:
    ekstre_turu = st.selectbox("Filtreleme Türü", ["Tümü", "Sadece Gelirler (Aidat)", "Sadece Giderler (Masraf)"], key="ex_tur")

ekstre_listesi = []

# Gelir Verilerini Ekle
df_gelirler_filtreli = df_aidat[df_aidat["Durum"].str.contains("Ödendi", na=False)].copy()
if not df_gelirler_filtreli.empty:
    if secilen_cari != "Tümü (Tüm Bina / Gelir+Gider)":
        df_gelirler_filtreli = df_gelirler_filtreli[df_gelirler_filtreli["Daire No"] == secilen_cari]
    for idx_g, row in df_gelirler_filtreli.iterrows():
        ekstre_listesi.append({
            "Tarih": str(row["Ödeme Tarihi"]), "İşlem Türü": "Gelir (Aidat)",
            "Açıklama / Sakin": str(row["Daire No"]), "Kategori / Dönem": str(row["Dönem Ay"]),
            "Giren Bakiye (Gelir)": float(row["Ödenen Tutar (TL)"]), "Çıkan Bakiye (Gider)": 0.0
        })

# Gider Verilerini Ekle
if not df_masraf.empty:
    if secilen_cari == "Tümü (Tüm Bina / Gelir+Gider)":
        for idx_m, row in df_masraf.iterrows():
            ekstre_listesi.append({
                "Tarih": str(row["Tarih"]), "İşlem Türü": "Gider (Masraf)",
                "Açıklama / Sakin": str(row["Açıklama"]), "Kategori / Dönem": str(row["Kategori"]),
                "Giren Bakiye (Gelir)": 0.0, "Çıkan Bakiye (Gider)": float(row["Tutar (TL)"])
            })

# KESİN ÇÖZÜM: Girinti hatası veren if bloğu tamamen kaldırıldı, doğrudan düz dataframe basılıyor
df_ekstre = pd.DataFrame(ekstre_listesi)

if not df_ekstre.empty:
    if ekstre_turu == "Sadece Gelirler (Aidat)":
        df_ekstre = df_ekstre[df_ekstre["İşlem Türü"] == "Gelir (Aidat)"]
