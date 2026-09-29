import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Kontrola Logističkih Računa", page_icon="📦", layout="wide"
)

st.title("📦 Sustav za Kontrolu i Analizu Logističkih Računa")
st.write(
    "Učitaj mjesečnu tablicu pošiljaka. Ugrađeni su točni ugovoreni uvjeti i"
    " cjenik dodatnih usluga."
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


# Ugovoreni cjenik po zonama i masama
cjenik = {
    "Zona 1": {
        1.0: 2.84,
        2.0: 3.15,
        5.0: 3.44,
        10.0: 4.79,
        15.0: 5.39,
        20.0: 6.07,
        25.0: 6.82,
        30.0: 7.34,
        35.0: 8.09,
        40.0: 8.38,
        45.0: 8.46,
        50.0: 9.21,
    },
    "Zona 2": {
        1.0: 3.31,
        2.0: 3.80,
        5.0: 4.19,
        10.0: 5.54,
        15.0: 6.29,
        20.0: 7.18,
        25.0: 8.09,
        30.0: 8.68,
        35.0: 9.59,
        40.0: 10.04,
        45.0: 10.49,
        50.0: 11.24,
    },
    "Zona 3": {
        1.0: 3.31,
        2.0: 3.80,
        5.0: 4.19,
        10.0: 5.54,
        15.0: 6.29,
        20.0: 7.18,
        25.0: 8.09,
        30.0: 8.68,
        35.0: 9.59,
        40.0: 10.04,
        45.0: 10.49,
        50.0: 11.24,
    },
}

cijena_preko_50_z1 = 0.19
cijena_preko_50_z2 = 0.22


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
    "Prosječna cijena goriva (€):",
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
        f" **{ukupno_kartona}** | Ostvareni popust na kartone:"
        f" **{popust_posto}%**"
    )

    usluge_definicije = {
        "CODC": 0.53,
        "CODH": 0.66,
        "OVSC": 0.0,
        "OVWC": 0.0,
        "OVWT": 6.40,
        "OVSZ": 6.40,
        "Returned Parcel": 0.5,  # 50% od osnovne cijene pošiljke
        "RTSC": 0.0,
        "SMS Notification": 0.13,
    }

    for idx, row in df.iterrows():
      pbr = row.get("Consignee ZIP Code", 10000)
      masa = float(row.get("Weight", 0.0))
      naplaceni_transport = float(row.get("Transport Price", 0.0))
      naplaceno_gorivo = float(row.get("Fuel Surcharge", 0.0))

      d_slanja = row.get("Shipping Date", None)
      d_dostave = row.get("Delivery Time", None)
      tranzit_dani = izracunaj_radne_dane(d_slanja, d_dostave)

      zona = odredis_zonu(pbr)

      osnovna_cjenik = izracunaj_osnovnu_cijenu(masa, zona)
      popust_iznos = osnovna_cjenik * (popust_posto / 100.0)
      trebalo_s_popustom = osnovna_cjenik - popust_iznos

      naplaceno = naplaceni_transport
      razlika_transport = naplaceno - trebalo_s_popustom

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
          "Naplaćeno (€)": round(naplaceno, 2),
          "Trebalo po cjeniku (€)": round(osnovna_cjenik, 2),
          "Popust iznos (€)": round(popust_iznos, 2),
          "Trebalo s popustom (€)": round(trebalo_s_popustom, 2),
          "Razlika (Preplata) (€)": round(razlika_transport, 2),
      }

      postoji_dodatna_naplata = False
      sve_dodatne_razlike = 0.0

      for usluga, ugovorena_cijena_usluge in usluge_definicije.items():
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

        # Izračun ugovorene cijene za ovu specifičnu uslugu
        if usluga == "Returned Parcel":
          ocekivana_usluga = trebalo_s_popustom * 0.5 if q_iznos > 0 else 0.0
        elif ugovorena_cijena_usluge > 0:
          ocekivana_usluga = ugovorena_cijena_usluge * (
              q_iznos if q_iznos > 0 else 1
          )
        else:
          ocekivana_usluga = (
              p_iznos  # Ako nemamo fiksnu cijenu, pratimo naplaćeno
          )

        razlika_usluge = p_iznos - ocekivana_usluga
        sve_dodatne_razlike += max(0, razlika_usluge)

        red_podataka[f"{usluga} - Naplaćeno (€)"] = round(p_iznos, 2)
        red_podataka[f"{usluga} - Očekivano (€)"] = round(ocekivana_usluga, 2)

      red_podataka["Ima Dodatnih Usluga"] = postoji_dodatna_naplata
      red_podataka["Ukupna preplata dodatnih usluga (€)"] = round(
          sve_dodatne_razlike, 2
      )
      rezultati.append(red_podataka)

    res_df = pd.DataFrame(rezultati)

    # Kreiranje 5 taba (izvještaja)
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 1. Izvještaj: Tranzit po zonama",
        "⚖️ 2. Izvještaj: Usporedba svih cijena",
        "🚨 3. Izvještaj: Samo razlike i preplate",
        "📈 4. Izvještaj: Ukupne sume i financije",
        "🛠️ 5. Izvještaj: Dodatne usluge",
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
      sumnjive = res_df[
          (res_df["Razlika (Preplata) (€)"] > 0.05)
          | (res_df["Ukupna preplata dodatnih usluga (€)"] > 0.05)
      ]
      st.dataframe(sumnjive, use_container_width=True)
      st.download_button(
          "📥 Preuzmi Izvještaj 3 (Excel)",
          to_excel(sumnjive),
          "preplate.xlsx",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )

    # --- TAB 4: UKUPNE SUME I FINANCIJE ---
    with tab4:
      st.subheader(
          "💰 Zbirni financijski pregled (Ukupne sume za cijelu tablicu)"
      )
      sum_naplaceno = res_df["Naplaćeno (€)"].sum()
      sum_cjenik = res_df["Trebalo po cjeniku (€)"].sum()
      sum_popust = res_df["Popust iznos (€)"].sum()
      sum_s_popustom = res_df["Trebalo s popustom (€)"].sum()
      sum_razlika = res_df["Razlika (Preplata) (€)"].clip(lower=0).sum()
      sum_dodatne_preplate = res_df[
          "Ukupna preplata dodatnih usluga (€)"
      ].sum()

      col1, col2, col3 = st.columns(3)
      col1.metric(
          label="Ukupno su naplatili", value=f"{sum_naplaceno:,.2f} €"
      )
      col2.metric(
          label="Ukupno trebalo po cjeniku", value=f"{sum_cjenik:,.2f} €"
      )
      col3.metric(
          label="Ukupni iznos popusta", value=f"-{sum_popust:,.2f} €"
      )

      col4, col5, col6 = st.columns(3)
      col4.metric(
          label="Ukupno trebalo biti s popustom", value=f"{sum_s_popustom:,.2f} €"
      )
      col5.metric(
          label="Preplata na transportu", value=f"{sum_razlika:,.2f} €"
      )
      col6.metric(
          label="Preplata na dodatnim uslugama",
          value=f"{sum_dodatne_preplate:,.2f} €",
      )

      zbirni_df = pd.DataFrame([{
          "Ukupno pošiljaka": ukupno_pošiljaka,
          "Ukupno kartona": ukupno_kartona,
          "Ostvareni popust (%)": f"{popust_posto}%",
          "Ukupno naplaćeno (€)": round(sum_naplaceno, 2),
          "Ukupno po cjeniku (€)": round(sum_cjenik, 2),
          "Ukupni iznos popusta (€)": round(sum_popust, 2),
          "Ukupno s popustom (€)": round(sum_s_popustom, 2),
          "Preplata transport (€)": round(sum_razlika, 2),
          "Preplata dodatne usluge (€)": round(sum_dodatne_preplate, 2),
      }])

      st.markdown("---")
      st.write("Pregled zbirnih podataka za preuzimanje:")
      st.dataframe(zbirni_df, use_container_width=True)
      st.download_button(
          "📥 Preuzmi Zbirni Izvještaj (Excel)",
          to_excel(zbirni_df),
          "zbirni_financijski_izvjestaj.xlsx",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )

    # --- TAB 5: DODATNE USLUGE ---
    with tab5:
      st.subheader(
          "🛠️ Izvještaj naplaćenih dodatnih usluga uspoređenih s ugovorom"
      )
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
