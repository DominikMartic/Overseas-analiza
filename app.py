import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Kontrola Logističkih Računa", page_icon="📦", layout="wide"
)

st.title("📦 Sustav za Kontrolu i Analizu Logističkih Računa")
st.write(
    "Učitaj mjesečnu tablicu pošiljaka. Cjenik je ažuriran i uvećan za 5%."
)

# Definiranje Zona 3 prema tablici (otoci i posebni režim dostave)
zona_3_pbr = [
    20210,
    20213,
    20231,
    20216,
    20000,
    20215,
    20207,
    20236,
    20234,
    20218,
    20217,
    20232,
    20205,
    20233,
    20235,
]


def odredis_zonu(pbr):
  try:
    pbr = int(pbr)
  except:
    return "Zona 2"

  if pbr in zona_3_pbr:
    return "Zona 3"
  elif 10000 <= pbr <= 10450:
    return "Zona 1"
  else:
    return "Zona 2"


# Ugovoreni cjenik po zonama i masama (uvećan za 5%, bez PDV-a)
cjenik = {
    "Zona 1": {
        1.0: 2.84 * 1.05,
        2.0: 3.15 * 1.05,
        5.0: 3.44 * 1.05,
        10.0: 4.79 * 1.05,
        15.0: 5.39 * 1.05,
        20.0: 6.07 * 1.05,
        25.0: 6.82 * 1.05,
        30.0: 7.34 * 1.05,
        35.0: 8.09 * 1.05,
        40.0: 8.38 * 1.05,
        45.0: 8.46 * 1.05,
        50.0: 9.21 * 1.05,
    },
    "Zona 2": {
        1.0: 3.31 * 1.05,
        2.0: 3.80 * 1.05,
        5.0: 4.19 * 1.05,
        10.0: 5.54 * 1.05,
        15.0: 6.29 * 1.05,
        20.0: 7.18 * 1.05,
        25.0: 8.09 * 1.05,
        30.0: 8.68 * 1.05,
        35.0: 9.59 * 1.05,
        40.0: 10.04 * 1.05,
        45.0: 10.49 * 1.05,
        50.0: 11.24 * 1.05,
    },
    "Zona 3": {
        1.0: 3.31 * 1.05,
        2.0: 3.80 * 1.05,
        5.0: 4.19 * 1.05,
        10.0: 5.54 * 1.05,
        15.0: 6.29 * 1.05,
        20.0: 7.18 * 1.05,
        25.0: 8.09 * 1.05,
        30.0: 8.68 * 1.05,
        35.0: 9.59 * 1.05,
        40.0: 10.04 * 1.05,
        45.0: 10.49 * 1.05,
        50.0: 11.24 * 1.05,
    },
}

cijena_preko_50_z1 = 0.19 * 1.05
cijena_preko_50_z2 = 0.22 * 1.05


def izracunaj_osnovnu_cijenu(masa, zona):
  granice = [1.0, 2.0, 5.0, 10.0, 15.0, 20.0, 25.0, 30.0, 35.0, 40.0, 45.0, 50.0]
  z_tablica = cjenik.get(zona, cjenik["Zona 2"])

  if masa <= 50.0:
    for g in granice:
      if masa <= g:
        osnova = z_tablica[g]
        break
    else:
      osnova = z_tablica[50.0]
  else:
    baza = z_tablica[50.0]
    višak = masa - 50.0
    dodatak_po_kg = (
        cijena_preko_50_z1 if zona == "Zona 1" else cijena_preko_50_z2
    )
    osnova = baza + višak * dodatak_po_kg

  # Uvjet za Zonu 3: na početnu cijenu dodaje se 25%
  if zona == "Zona 3":
    osnova = osnova * 1.25

  return osnova


# Izračun postotka goriva prema razredima
fn_gorivo = lambda c: (
    0.0
    if c <= 1.35
    else (
        1.0
        if c <= 1.42
        else (
            2.0
            if c <= 1.49
            else (
                3.0
                if c <= 1.56
                else 3.0 + (int((c - 1.56) // 0.07) + 1) * 1.0
            )
        )
    )
)


def izracunaj_radne_dane(datum_slanja, datum_dostave):
  try:
    d1 = pd.to_datetime(datum_slanja, format="%d.%m.%Y", errors="coerce")
    d2 = pd.to_datetime(datum_dostave, format="%d.%m.%Y", errors="coerce")
    if pd.isna(d1) or pd.isna(d2):
      return "Nema informacije"
    radni_dani = pd.bdate_range(start=d1, end=d2).shape[0] - 1
    return max(0, radni_dani)
  except:
    return "Nema informacije"


def to_excel(df):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Izvjestaj")
  return output.getvalue()


# Bočna traka
st.sidebar.header("Parametri obračuna")
trenutna_cijena_goriva = st.sidebar.number_input(
    "Prosječna cijena goriva (€ bez PDV-a):",
    min_value=1.00,
    max_value=3.00,
    value=1.35,
    step=0.01,
)
posto_goriva = fn_gorivo(trenutna_cijena_goriva)
st.sidebar.info(
    f"Izračunati dodatak za gorivo prema razredima: **{posto_goriva:.1f}%**"
)

# Učitavanje datoteke
uploaded_file = st.file_uploader(
    "Učitaj Excel ili CSV tablicu s pošiljkama", type=["xlsx", "csv"]
)

if uploaded_file is not None:
  if uploaded_file.name.endswith(".xlsx"):
    df = pd.read_excel(uploaded_file)
  else:
    df = pd.read_csv(uploaded_file)

  st.success("Tablica uspješno učitana!")

  if st.button("Generiraj izvještaje"):
    rezultati = []
    ukupno_pošiljaka = len(df)

    ukupno_kartona = (
        int(df["Number of Parcels"].sum())
        if "Number of Parcels" in df.columns
        else 0
    )

    if ukupno_kartona >= 5000:
      popust_posto = 5.0
    elif ukupno_kartona >= 4001:
      popust_posto = 4.0
    elif ukupno_kartona >= 3001:
      popust_posto = 3.0
    elif ukupno_kartona >= 2000:
      popust_posto = 2.0
    else:
      popust_posto = 0.0

    st.info(
        f"📊 Obrađeno pošiljaka: **{ukupno_pošiljaka}** | Ukupno kartona:"
        f" **{ukupno_kartona}** | Ostvareni količinski popust na fakturu:"
        f" **{popust_posto}%**"
    )

    usluge_lista = [
        "CODC",
        "CODH",
        "OVSC",
        "OVWC",
        "OVWT",
        "OVSZ",
        "Returned Parcel",
        "RTSC",
        "SMS Notification",
    ]

    for idx, row in df.iterrows():
      pbr = row.get("Consignee ZIP Code", 10000)
      masa = float(row.get("Weight", 0.0))
      naplaceni_transport = float(row.get("Transport Price", 0.0))
      naplaceno_gorivo = float(row.get("Fuel Surcharge", 0.0))

      d_slanja = row.get("Shipping Date", None)
      d_dostave = row.get("Delivery Time", None)
      tranzit_dani = izracunaj_radne_dane(d_slanja, d_dostave)

      zona = odredis_zonu(pbr)

      # Očekivane ugovorene vrijednosti po novom cjeniku (+5%)
      ugovorena_osnova = izracunaj_osnovnu_cijenu(masa, zona)
      ugovoreno_gorivo = ugovorena_osnova * (posto_goriva / 100.0)

      # Zbrajanje svih naplaćenih dodatnih usluga za ovu pošiljku
      zbroj_naplacenih_dodatnih = 0.0
      postoji_dodatna_naplata = False

      red_podataka = {
          "RedniBroj": idx + 1,
          "Shipment ID": row.get("Shipment ID", ""),
          "Consignee Name": row.get("Consignee Name", ""),
          "Consignee Town": row.get("Consignee Town", ""),
          "Number of Parcels": row.get("Number of Parcels", 1),
          "Reference 1": row.get("Reference 1", ""),
          "ZIP": pbr,
          "Zona": zona,
          "Masa (kg)": masa,
          "Slanje": d_slanja,
          "Dostava": d_dostave,
          "Tranzit (radni dani)": tranzit_dani,
          "Naplaćeni Transport (€)": round(naplaceni_transport, 2),
          "Ugovorena Osnova (€)": round(ugovorena_osnova, 2),
          "Naplaćeno Gorivo (€)": round(naplaceno_gorivo, 2),
          "Ugovoreno Gorivo (€)": round(ugovoreno_gorivo, 2),
      }

      for usluga in usluge_lista:
        p_col = next(
            (
                c
                for c in df.columns
                if c.lower().replace(" ", "")
                == f"price{usluga.lower().replace(' ', '')}"
            ),
            None,
        )
        q_col = next(
            (
                c
                for c in df.columns
                if c.lower().replace(" ", "")
                == f"quantity{usluga.lower().replace(' ', '')}"
            ),
            None,
        )

        p_iznos = float(row[p_col]) if p_col and pd.notna(row[p_col]) else 0.0
        q_iznos = float(row[q_col]) if q_col and pd.notna(row[q_col]) else 0.0

        if p_iznos > 0 or q_iznos > 0:
          postoji_dodatna_naplata = True
          zbroj_naplacenih_dodatnih += p_iznos

        red_podataka[f"{usluga} - Naplaćeno (€)"] = round(p_iznos, 2)

      red_podataka["Naplaćene Dodatne Usluge Ukupno (€)"] = round(
          zbroj_naplacenih_dodatnih, 2
      )
      red_podataka["Sveukupno Naplaćeno (€)"] = round(
          naplaceni_transport + naplaceno_gorivo + zbroj_naplacenih_dodatnih, 2
      )

      očekivano_sveukupno = (
          ugovorena_osnova + ugovoreno_gorivo + zbroj_naplacenih_dodatnih
      )
      red_podataka["Sveukupno Očekivano (€)"] = round(očekivano_sveukupno, 2)

      red_podataka["Ima Dodatnih Usluga"] = postoji_dodatna_naplata
      rezultati.append(red_podataka)

    res_df = pd.DataFrame(rezultati)

    # Kreiranje 5 taba (izvještaja)
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 1. Izvještaj: Tranzit po zonama",
        "⚖️ 2. Izvještaj: Usporedba svih cijena",
        "🚨 3. Izvještaj: Samo razlike i preplate",
        "📈 4. Izvještaj: Zbirne sume fakture",
        "🛠️️ 5. Izvještaj: Dodatne usluge",
    ])

    # --- TAB 1: TRANZIT PO ZONAMA ---
    with tab1:
      st.subheader("Analiza tranzita pošiljaka po zonama (u radnim danima)")
      tranzit_view = res_df[
          [
              "RedniBroj",
              "Shipment ID",
              "Consignee Name",
              "Consignee Town",
              "Number of Parcels",
              "Reference 1",
              "ZIP",
              "Zona",
              "Slanje",
              "Dostava",
              "Tranzit (radni dani)",
          ]
      ]
      st.dataframe(tranzit_view, use_container_width=True)
      st.download_button(
          "📥 Preuzmi Izvještaj 1 (Excel)",
          to_excel(tranzit_view),
          "tranzit_po_zonama.xlsx",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )

    # --- TAB 2: USPOREDBA SVIH CIJENA ---
    with tab2:
      st.subheader("Detaljna usporedba za sve pošiljke")
      st.dataframe(res_df, use_container_width=True)
      st.download_button(
          "📥 Preuzmi Izvještaj 2 (Excel)",
          to_excel(res_df),
          "sve_usporedbe.xlsx",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )

    # --- TAB 3: SAMO RAZLIKE I PREPLATE ---
    with tab3:
      st.subheader(
          "Izdvojene preplate na transportu, gorivu i dodatnim uslugama"
      )
      res_df["Razlika Fakture (€)"] = (
          res_df["Sveukupno Naplaćeno (€)"]
          - res_df["Sveukupno Očekivano (€)"]
      )
      sumnjive = res_df[res_df["Razlika Fakture (€)"] > 0.05]
      st.dataframe(sumnjive, use_container_width=True)
      st.download_button(
          "📥 Preuzmi Izvještaj 3 (Excel)",
          to_excel(sumnjive),
          "preplate.xlsx",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )

    # --- TAB 4: ZBIRNE SUME FAKTURE (BEZ PDV-a SA POPUSTOM NA KRAJU) ---
    with tab4:
      st.subheader(
          "📈 Zbirni financijski pregled cijele fakture (Sve cijene bez PDV-a)"
      )

      # Sumiranje po stavkama s računa (što su naplatili)
      uk_naplaceni_transport = res_df["Naplaćeni Transport (€)"].sum()
      uk_naplaceno_gorivo = res_df["Naplaćeno Gorivo (€)"].sum()
      uk_naplacene_dodatne = res_df["Naplaćene Dodatne Usluge Ukupno (€)"].sum()
      sveukupno_naplaceno_racun = (
          uk_naplaceni_transport + uk_naplaceno_gorivo + uk_naplacene_dodatne
      )

      # Sumiranje po ugovoru (što je trebalo biti)
      uk_ugovorena_osnova = res_df["Ugovorena Osnova (€)"].sum()
      uk_ugovoreno_gorivo = res_df["Ugovoreno Gorivo (€)"].sum()

      # Primjena popusta na UKUPNU zbrojenu osnovicu transporta na razini fakture
      iznos_popusta_faktura = uk_ugovorena_osnova * (popust_posto / 100.0)
      ugovorena_osnova_nakon_popusta = (
          uk_ugovorena_osnova - iznos_popusta_faktura
      )

      # Preračun goriva nakon popusta na osnovu
      ugovoreno_gorivo_nakon_popusta = ugovorena_osnova_nakon_popusta * (
          posto_goriva / 100.0
      )

      # Sveukupno očekivano po ugovoru nakon popusta na cijelu fakturu
      sveukupno_ocekivano_ugovor = (
          ugovorena_osnova_nakon_popusta
          + ugovoreno_gorivo_nakon_popusta
          + uk_naplacene_dodatne
      )

      # Konačna preplata / razlika
      konačna_preplata = (
          sveukupno_naplaceno_racun - sveukupno_ocekivano_ugovor
      )

      # Prikaz preko metrika
      col1, col2, col3 = st.columns(3)
      col1.metric(
          label="Sveukupno su naplatili (Bez PDV-a)",
          value=f"{sveukupno_naplaceno_racun:,.2f} €",
      )
      col2.metric(
          label="Sveukupno trebalo po ugovoru",
          value=f"{sveukupno_ocekivano_ugovor:,.2f} €",
      )
      col3.metric(
          label="Ukupna preplata / Višak za povrat",
          value=f"{max(0, konačna_preplata):,.2f} €",
      )

      st.markdown("---")
      st.markdown("### Detaljna struktura zbroja fakture:")

      zbirni_detalji = pd.DataFrame([
          {
              "Kategorija troška": "Transport (Osnovna cijena - uvećano 5%)",
              "Što su naplatili (€)": round(uk_naplaceni_transport, 2),
              "Što je trebalo biti (€)": round(uk_ugovorena_osnova, 2),
          },
          {
              "Kategorija troška": (
                  f"Količinski popust na fakturu ({popust_posto}%)"
              ),
              "Što su naplatili (€)": 0.00,
              "Što je trebalo biti (€)": round(-iznos_popusta_faktura, 2),
          },
          {
              "Kategorija troška": (
                  f"Dodatak za gorivo ({posto_goriva:.1f}%)"
              ),
              "Što su naplatili (€)": round(uk_naplaceno_gorivo, 2),
              "Što je trebalo biti (€)": round(
                  ugovoreno_gorivo_nakon_popusta, 2
              ),
          },
          {
              "Kategorija troška": "Sve dodatne usluge (CODC, OVWT, SMS...)",
              "Što su naplatili (€)": round(uk_naplacene_dodatne, 2),
              "Što je trebalo biti (€)": round(uk_naplacene_dodatne, 2),
          },
          {
              "Kategorija troška": "SVEUKUPNO ZA CIJELU FAKTURU",
              "Što su naplatili (€)": round(sveukupno_naplaceno_racun, 2),
              "Što je trebalo biti (€)": round(sveukupno_ocekivano_ugovor, 2),
          },
      ])

      st.dataframe(zbirni_detalji, use_container_width=True)

      st.download_button(
          label="📥 Preuzmi Zbirni Financijski Izvještaj (Excel)",
          data=to_excel(zbirni_detalji),
          file_name="zbirni_financijski_izvjestaj_faktura.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

    # --- TAB 5: DODATNE USLUGE ---
    with tab5:
      st.subheader("🛠️ Izvještaj pošiljaka s naplaćenim dodatnim uslugama")
      dodatne_df = res_df[res_df["Ima Dodatnih Usluga"] == True]

      if dodatne_df.empty:
        st.success("Nema pošiljaka s naplaćenim dodatnim uslugama u ovoj tablici!")
      else:
        st.write(
            f"Pronađeno pošiljaka s dodatnim uslugama: {len(dodatne_df)}"
        )
        st.dataframe(dodatne_df, use_container_width=True)
        st.download_button(
            label="📥 Preuzmi Izvještaj Dodatnih Usluga (Excel)",
            data=to_excel(dodatne_df),
            file_name="izvjestaj_dodatne_usluge.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
