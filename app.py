import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

# Læs template-filen lokalt
@st.cache_data
def load_template():
    with open("apendix2-template.xlsx", "rb") as f:
        return f.read()

# Læs masterdata-filen lokalt
@st.cache_data
def load_masterdata():
    return pd.read_excel("Muuto_Master_Data_CON_January_2025_DKK.xlsx", engine="openpyxl")

st.title("Streamlit App til Data Mapping")
st.write("Indtast varenumre (et varenummer per linje):")
user_input = st.text_area("Varenumre", height=200)

if st.button("Start behandling"):
    progress_bar = st.progress(0)
    status_text = st.empty()

    # Hent de nødvendige filer
    template_content = load_template()
    masterdata_df = load_masterdata()

    # Split inputtet i en liste af varenumre og fjern tomme linjer
    varenumre = [line.strip() for line in user_input.splitlines() if line.strip()]
    results = []
    unmatched = []

    total = len(varenumre)
    for i, varenr in enumerate(varenumre):
        status_text.text(f"Behandler varenummer {i+1} af {total}")
        # Antag, at masterdata indeholder varenummer i kolonnen "PRODUCT"
        match = masterdata_df[masterdata_df["PRODUCT"] == varenr]
        if not match.empty:
            row = match.iloc[0]
            # Hent værdier fra masterdata
            product_name = row["PRODUCT"]
            product_description = row["PRODUCT DESCRIPTION"]
            manufacturing_country = row["COUNTRY OF ORIGIN"]
            lead_time = row["LEAD TIME"]
            if lead_time == "-":
                lead_time = "Ready to ship"
            warranty = row["WARRANTY"]
            contract_price = row["CONTRACT PRICE"]
            list_price = f"{contract_price} DKK"
            
            # For "Product Series name" er der ikke specificeret en kilde – sættes som tom
            product_series_name = ""
            
            results.append({
                "Product Series name": product_series_name,
                "Product Name": product_name,
                "Product item number": varenr,
                "Product Description": product_description,
                "Manufacturing country": manufacturing_country,
                "Lead time [weeks]": lead_time,
                "Product Guarantee period [years]": warranty,
                "List Price [your currency]": list_price
            })
        else:
            unmatched.append(varenr)
        progress_bar.progress((i+1) / total)
    
    # Åbn template-filen med openpyxl
    wb = load_workbook(filename=BytesIO(template_content))
    ws = wb.active

    # Indsæt resultaterne i template-filen fra række 7 (kolonnerne B til I)
    start_row = 7
    for idx, res in enumerate(results, start=start_row):
        ws[f"B{idx}"] = res["Product Series name"]
        ws[f"C{idx}"] = res["Product Name"]
        ws[f"D{idx}"] = res["Product item number"]
        ws[f"E{idx}"] = res["Product Description"]
        ws[f"F{idx}"] = res["Manufacturing country"]
        ws[f"G{idx}"] = res["Lead time [weeks]"]
        ws[f"H{idx}"] = res["Product Guarantee period [years]"]
        ws[f"I{idx}"] = res["List Price [your currency]"]

    # Gem den udfyldte fil til en BytesIO-strøm for download
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("Behandlingen er færdig!")
    st.download_button("Download udfyldt fil",
                       data=output,
                       file_name="udfyldt_template.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    if unmatched:
        st.warning("Følgende varenumre blev ikke fundet i masterdata:")
        st.text("\n".join(unmatched))
