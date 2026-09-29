import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Kontrola Logističkih Računa", page_icon="📦", layout="wide"
)

st.title("📦 Napredna Kontrola Logističkih Računa i Tranzita")
st.write(
    "Učitaj tablicu s pošiljkama, upiši cijenu goriva sa strane i pokreni provjeru. Zona 3, radni dani tranzita i sve dodatne usluge sada su razdvojeni u zasebne stupce."
)

# Definiranje Zona 3 prema tvojoj tablici (otoci i posebni režim dostave)
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
    "Zona 3": {  # Zona 3 koristi osnovu Zone 2 ili specifično
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
        return z_tablica[g]
    return z_tablica[50.0]
  else:
    baza = z_tablica[50.0]
    višak = masa - 50.0
    dodatak_po_kg = (
        cijena_preko_50_z1 if zona == "Zona 1" else cijena_preko_50_z2
    )
    return baza + višak * dodatak_po_kg


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


# Bočna traka za unos cijene goriva
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

  if st.button("Pokreni detaljnu provjeru i analizu"):
    rezultati = []
    ukupno_pošiljaka = len(df)

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
      ugovorena_osnova = izracunaj_osnovnu_cijenu(masa, zona)
      ugovoreno_gorivo = ugovorena_osnova * (posto_goriva / 100.0)

      razlika_transport = naplaceni_transport - ugovorena_osnova
      razlika_gorivo = naplaceno_gorivo - ugovoreno_gorivo

      red_podataka = {
          "RedniBroj": idx + 1,
          "Shipment ID": row.get("Shipment ID", ""),
          "ZIP": pbr,
          "Zona": zona,
          "Masa (kg)": masa,
          "Slanje": d_slanja,
          "Dostava": d_dostave,
          "Tranzit (radni dani)": tranzit_dani,
          "Ugovorena Osnova (€)": round(ugovorena_osnova, 2),
          "Naplaćeni Transport (€)": round(naplaceni_transport, 2),
          "Razlika Transport (€)": round(razlika_transport, 2),
          "Ugovoreno Gorivo (€)": round(ugovoreno_gorivo, 2),
          "Naplaćeno Gorivo (€)": round(naplaceno_gorivo, 2),
          "Razlika Gorivo (€)": round(razlika_gorivo, 2),
      }

      postoji_dodatna_naplata = False
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

        red_podataka[f"{usluga} - Cijena (€)"] = round(p_iznos, 2)
        red_podataka[f"{usluga} - Količina"] = q_iznos

      red_podataka["Ima Dodatnih Usluga"] = postoji_dodatna_naplata
      rezultati.append(red_podataka)

    res_df = pd.DataFrame(rezultati)

    # Količinski popust na kraju na temelju broja pošiljaka
    if ukupno_pošiljaka >= 5000:
      popust_posto = 5.0
    elif ukupno_pošiljaka >= 4001:
      popust_posto = 4.0
    elif ukupno_pošiljaka >= 3001:
      popust_posto = 3.0
    elif ukupno_pošiljaka >= 2000:
      popust_posto = 2.0
    else:
      popust_posto = 0.0

    st.success(
        f"Analiza završena! Obrađeno pošiljaka: {ukupno_pošiljaka} | Mjesečni"
        f" količinski popust: **{popust_posto}%**"
    )

    # Poseban prikaz pošiljaka za Zonu 3
    zona_3_df = res_df[res_df["Zona"] == "Zona 3"]
    if not zona_3_df.empty:
      st.subheader(
          f"🏝️ Poseban pregled pošiljaka za Zonu 3 (Ukupno: {len(zona_3_df)})"
      )
      st.dataframe(zona_3_df)

    # Filtriramo stavke s preplatom ili dodatnim uslugama
    sumnjive = res_df[
        (res_df["Razlika Transport (€)"] > 0.05)
        | (res_df["Razlika Gorivo (€)"] > 0.05)
        | (res_df["Ima Dodatnih Usluga"] == True)
    ]

    col1, col2 = st.columns(2)
    col1.metric(
        label="Ukupna preplata na osnovnoj cijeni (kilaža)",
        value=f"{res_df['Razlika Transport (€)'].clip(lower=0).sum():.2f} €",
    )
    col2.metric(
        label="Ukupna preplata na dodatku za gorivo",
        value=f"{res_df['Razlika Gorivo (€)'].clip(lower=0).sum():.2f} €",
    )

    st.subheader(
        "Popis pošiljaka s pogrešnim naplatama / dodatnim uslugama i radnim"
        " danima tranzita:"
    )
    st.dataframe(sumnjive)

    # Preuzimanje izvještaja
    csv = sumnjive.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Preuzmi detaljni izvještaj za reklamaciju (CSV)",
        data=csv,
        file_name="detaljne_reklamacije_logistika.csv",
        mime="text/csv",
    )
