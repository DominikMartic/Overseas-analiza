import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Kontrola Logističkih Računa", page_icon="📦", layout="wide"
)

st.title("📦 Sustav za Kontrolu i Analizu Logističkih Računa")
st.write(
    "Učitaj mjesečnu tablicu pošiljaka i pregledaj podatke kroz 3 različita"
    " izvještaja u nastavku."
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

  if st.button("Generiraj 3 izvještaja"):
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

    # Kreiranje 3 taba (izvještaja)
    tab1, tab2, tab3 = st.tabs([
        "📊 1. Izvještaj: Tranzit po zonama",
        "⚖️ 2. Izvještaj: Usporedba svih cijena",
        "🚨 3. Izvještaj: Samo razlike i preplate",
    ])

    # --- TAB 1: TRANZIT PO ZONAMA ---
    with tab1:
      st.subheader("Analiza tranzita pošiljaka po zonama (u radnim danima)")
      tranzit_view = res_df[
          [
              "RedniBroj",
              "Shipment ID",
              "ZIP",
              "Zona",
              "Slanje",
              "Dostava",
              "Tranzit (radni dani)",
          ]
      ]
      st.dataframe(tranzit_view, use_container_width=True)

      excel_t1 = to_excel(tranzit_view)
      st.download_button(
          label="📥 Preuzmi Izvještaj 1 (Excel)",
          data=excel_t1,
          file_name="izvjestaj_tranzit_po_zonama.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

    # --- TAB 2: USPOREDBA SVIH CIJENA ---
    with tab2:
      st.subheader(
          "Detaljna usporedba naplaćenog vs. ugovorenog za sve pošiljke"
      )
      st.dataframe(res_df, use_container_width=True)

      excel_t2 = to_excel(res_df)
      st.download_button(
          label="📥 Preuzmi Izvještaj 2 (Excel)",
          data=excel_t2,
          file_name="izvjestaj_sve_usporedbe.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

    # --- TAB 3: SAMO RAZLIKE I PREPLATE ---
    with tab3:
      st.subheader("Izdvojene preplate i nepravilnosti za reklamaciju")
      sumnjive = res_df[
          (res_df["Razlika Transport (€)"] > 0.05)
          | (res_df["Razlika Gorivo (€)"] > 0.05)
          | (res_df["Ima Dodatnih Usluga"] == True)
      ]

      if sumnjive.empty:
        st.success("Nema pronađenih preplata ni nepravilnosti!")
      else:
        st.metric(
            label="Ukupan iznos preplate za povrat",
            value=f"{sumnjive['Razlika Transport (€)'].clip(lower=0).sum() + sumnjive['Razlika Gorivo (€)'].clip(lower=0).sum():.2f} €",
        )
        st.dataframe(sumnjive, use_container_width=True)

        excel_t3 = to_excel(sumnjive)
        st.download_button(
            label="📥 Preuzmi Izvještaj 3 za Reklamaciju (Excel)",
            data=excel_t3,
            file_name="izvjestaj_samo_preplate_reklamacije.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
