import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Kontrola Logističkih Računa", page_icon="📦", layout="wide"
)

st.title("📦 Sustav za Kontrolu i Analizu Logističkih Računa")
st.write(
    "Učitaj mjesečnu tablicu pošiljaka. Cjenik je uvećan za 5%, a tranzit se"
    " automatski kontrolira prema ugovorenim rokovima."
)


# Učitavanje dopuštenih dana isporuke po poštanskim brojevima iz tablice
@st.cache_data
def ucitaj_dopucene_rokove():
  # Ovdje pretpostavljamo da se tablica "mjesto, dani isporuke.xlsx" nalazi u istom direktoriju
  try:
    df_rokovi = pd.read_excel("mjesto, dani isporuke.xlsx")
    pbr_col = next((c for c in df_rokovi.columns if "poštanski" in c.lower()), None)
    dostava_col = next((c for c in df_rokovi.columns if "dostava" in c.lower()), None)

    if pbr_col and dostava_col:
      r_dict = {}
      for _, row in df_rokovi.iterrows():
        try:
          pbr = int(row[pbr_col])
          val_str = str(row[dostava_col])
          match = re.search(r"(\d+)\s*//", val_str)
          if match:
            dani = int(match.group(1))
          else:
            match_any = re.search(r"(\d+)", val_str)
            dani = int(match_any.group(1)) if match_any else 1
          
          # Uzimamo maksimalni dopušteni rok ako postoji više naselja za isti PBR
          if pbr in r_dict:
            r_dict[pbr] = max(r_dict[pbr], dani)
          else:
            r_dict[pbr] = dani
        except:
          continue
      return r_dict
  except Exception as e:
    st.warning(
        f"Napomena: Datoteka 'mjesto, dani isporuke.xlsx' nije pronađena ili je"
        f" greška: {e}. Koristit će se standardni zadani rokovi."
    )
  return {}


dopusteni_rokovi_dict = ucitaj_dopucene_rokove()

# Službene liste poštanskih brojeva za Zonu 2 i Zonu 3 prema dostavljenim tablicama
zona_3_pbr = [
    20230,
    20240,
    20242,
    20243,
    20244,
    20245,
    20246,
    20247,
    20248,
    20250,
    20260,
    20263,
    20264,
    20267,
    20269,
    20270,
    20271,
    20272,
    20273,
    20274,
    20275,
    21400,
    21403,
    21404,
    21405,
    21410,
    21412,
    21413,
    21420,
    21423,
    21424,
    21425,
    21426,
    21450,
    21454,
    21460,
    21462,
    21463,
    21465,
    21466,
    21467,
    21468,
    21469,
    21480,
    21483,
    21485,
    22240,
    22242,
    22243,
    22244,
    23212,
    23234,
    23249,
    23250,
    23251,
    23262,
    23263,
    23264,
    23271,
    23272,
    23273,
    23274,
    23275,
    51280,
    51281,
    51500,
    51511,
    51512,
    51513,
    51514,
    51515,
    51516,
    51517,
    51521,
    51522,
    51523,
    51550,
    51551,
    51554,
    51555,
    51556,
    51557,
    51559,
    51564,
    53291,
    53294,
    53296,
]

zona_2_pbr = [
    23440,
    23445,
    23446,
    31300,
    31321,
    31322,
    31323,
    31324,
    31542,
    31543,
    31555,
    43270,
    43271,
    43273,
    44202,
    44203,
    44210,
    44213,
    44214,
    44221,
    44222,
    44251,
    44271,
    44272,
    44273,
    44400,
    44401,
    44402,
    44412,
    44425,
    44430,
    44450,
    47220,
    47221,
    47222,
    47245,
    47246,
    47304,
    47305,
    47306,
    47307,
    47313,
    47314,
    48260,
    48265,
    48267,
    48323,
    51212,
    51213,
    51251,
    51300,
    51311,
    51312,
    51313,
    51314,
    51315,
    51316,
    51321,
    51322,
    51323,
    51324,
    51325,
    51328,
    51329,
    51414,
    51418,
    52000,
    52402,
    52420,
    52421,
    52422,
    52425,
    52426,
    52427,
    52428,
    52434,
    53000,
    53202,
    53203,
    53205,
    53206,
    53211,
    53212,
    53213,
    53221,
    53222,
    53223,
    53224,
    53230,
    53231,
    53234,
    53235,
    53244,
    53250,
    53252,
    53260,
    53261,
    53262,
    53284,
    53285,
    53286,
    53287,
    53288,
]


def odredis_zonu(pbr):
  try:
    pbr = int(pbr)
  except:
    return "Zona 2"

  if pbr in zona_3_pbr:
    return "Zona 3"
  elif pbr in zona_2_pbr:
    return "Zona 2"
  elif 10000 <= pbr <= 10450:
    return "Zona 1"
  else:
    return "Zona 2"  # Ostatak kopnene Hrvatske


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

  if zona == "Zona 3":
    osnova = osnova * 1.25

  return osnova


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
      return None
    radni_dani = pd.bdate_range(start=d1, end=d2).shape[0] - 1
    return max(0, radni_dani)
  except:
    return None


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
      try:
        pbr_int = int(pbr)
      except:
        pbr_int = 10000

      masa = float(row.get("Weight", 0.0))
      naplaceni_transport = float(row.get("Transport Price", 0.0))
      naplaceno_gorivo = float(row.get("Fuel Surcharge", 0.0))

      d_slanja = row.get("Shipping Date", None)
      d_dostave = row.get("Delivery Time", None)
      tranzit_dani = izracunaj_radne_dane(d_slanja, d_dostave)

      # Dohvat dopuštenog roka za PBR (ako nema u tablici, zadano je 1 dan za Z1, 2 za Z2/3)
      dopušteni_rok = dopusteni_rokovi_dict.get(
          pbr_int, (1 if 10000 <= pbr_int <= 10450 else 2)
      )

      if tranzit_dani is not None:
        kasni = tranzit_dani > dopušteni_rok
        tranzit_status = (
            f"{tranzit_dani} rad. dana (U roku)"
            if not kasni
            else f"🔴 {tranzit_dani} rad. dana (Kasni, rok je {dopušteni_rok})"
        )
      else:
        kasni = False
        tranzit_status = "Nema informacije"

      zona = odredis_zonu(pbr_int)
      ugovorena_osnova = izracunaj_osnovnu_cijenu(masa, zona)
      ugovoreno_gorivo = ugovorena_osnova * (posto_goriva / 100.0)

      zbroj_naplacenih_dodatnih = 0.0
      postoji_dodatna_naplata = False

      red_podataka = {
          "RedniBroj": idx + 1,
          "Shipment ID": row.get("Shipment ID", ""),
          "Consignee Name": row.get("Consignee Name", ""),
          "Consignee Town": row.get("Consignee Town", ""),
          "Number of Parcels": row.get("Number of Parcels", 1),
          "Reference 1": row.get("Reference 1", ""),
          "ZIP": pbr_int,
          "Zona": zona,
          "Masa (kg)": masa,
          "Slanje": d_slanja,
          "Dostava": d_dostave,
          "Stvarni Tranzit (dani)": (
              tranzit_dani if tranzit_dani is not None else -1
          ),
          "Dopušteni Rok (dani)": dopušteni_rok,
          "Status Dostave": tranzit_status,
          "Kasni": kasni,
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

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 1. Izvještaj: Tranzit i rokovi isporuke",
        "⚖️ 2. Izvještaj: Usporedba svih cijena",
        "🚨 3. Izvještaj: Samo razlike i preplate",
        "📈 4. Izvještaj: Zbirne sume fakture",
        "🛠 5. Izvještaj: Dodatne usluge",
    ])

    # --- TAB 1: TRANZIT I ROKOVI ISPORUKE ---
    with tab1:
      st.subheader(
          "Analiza tranzita pošiljaka i provjera ugovorenih rokova isporuke"
      )

      # Izračun postotka kašnjenja
      valid_tranzit = res_df[res_df["Stvarni Tranzit (dani)"] >= 0]
      uk_validnih = len(valid_tranzit)
      if uk_validnih > 0:
        uk_kasni = valid_tranzit["Kasni"].sum()
        postotak_kasnjenja = (uk_kasni / uk_validnih) * 100.0
        postotak_urednih = 100.0 - postotak_kasnjenja
      else:
        uk_kasni = 0
        postotak_kasnjenja = 0.0
        postotak_urednih = 100.0

      col_a, col_b, col_c = st.columns(3)
      col_a.metric(
          label="Uredno isporučeno u roku",
          value=f"{postotak_urednih:.1f}%",
          delta=f"{uk_validnih - uk_kasni} pošiljaka",
      )
      col_b.metric(
          label="Izvan ugovorenog roka (Kašnjenje)",
          value=f"{postotak_kasnjenja:.1f}%",
          delta=f"-{uk_kasni} pošiljaka",
          delta_color="inverse",
      )
      col_c.metric(
          label="Ukupno analizirano pošiljaka s datumima",
          value=f"{uk_validnih}",
      )

      st.markdown("---")

      tranzit_view = res_df[
          [
              "RedniBroj",
              "Shipment ID",
              "Consignee Name",
              "Consignee Town",
              "ZIP",
              "Zona",
              "Slanje",
              "Dostava",
              "Stvarni Tranzit (dani)",
              "Dopušteni Rok (dani)",
              "Status Dostave",
          ]
      ]
      st.dataframe(tranzit_view, use_container_width=True)
      st.download_button(
          "📥 Preuzmi Izvještaj 1 (Excel)",
          to_excel(tranzit_view),
          "analiza_tranzita_i_rokova.xlsx",
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

    # --- TAB 4: ZBIRNE SUME FAKTURE ---
    with tab4:
      st.subheader(
          "📈 Zbirni financijski pregled cijele fakture (Sve cijene bez PDV-a)"
      )
      uk_naplaceni_transport = res_df["Naplaćeni Transport (€)"].sum()
      uk_naplaceno_gorivo = res_df["Naplaćeno Gorivo (€)"].sum()
      uk_naplacene_dodatne = res_df["Naplaćene Dodatne Usluge Ukupno (€)"].sum()
      sveukupno_naplaceno_racun = (
          uk_naplaceni_transport + uk_naplaceno_gorivo + uk_naplacene_dodatne
      )

      uk_ugovorena_osnova = res_df["Ugovorena Osnova (€)"].sum()
      iznos_popusta_faktura = uk_ugovorena_osnova * (popust_posto / 100.0)
      ugovorena_osnova_nakon_popusta = (
          uk_ugovorena_osnova - iznos_popusta_faktura
      )
      ugovoreno_gorivo_nakon_popusta = ugovorena_osnova_nakon_popusta * (
          posto_goriva / 100.0
      )
      sveukupno_ocekivano_ugovor = (
          ugovorena_osnova_nakon_popusta
          + ugovoreno_gorivo_nakon_popusta
          + uk_naplacene_dodatne
      )
      konačna_preplata = (
          sveukupno_naplaceno_racun - sveukupno_ocekivano_ugovor
      )

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
