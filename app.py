import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

@st.cache_data
def load_template():
    """Læs template-filen lokalt."""
    with open("apendix2-template.xlsx", "rb") as f:
        return f.read()

@st.cache_data
def load_masterdata():
    """Læs masterdata-filen lokalt (med engine='openpyxl')."""
    return pd.read_excel("Muuto_Master_Data_CON_January_2025_DKK.xlsx", engine="openpyxl")

st.title("Streamlit App til Data Mapping")
st.write("Indtast varenumre (et varenummer per linje):")
user_input = st.text_area("Varenumre", height=200)

if st.button("Start behandling"):
    progress_bar = st.progress(0)
    status_text = st.empty()

    # Hent filer
    template_content = load_template()
    masterdata_df = load_masterdata()

    # Debug: vis masterdata kolonnenavne
    st.write("Masterdata kolonner:", masterdata_df.columns.tolist())

    # Opdel brugerens input i en liste og fjern tomme linjer
    varenumre = [line.strip() for line in user_input.splitlines() if line.strip()]
    results = []
    unmatched = []

    total = len(varenumre)
    for i, varenr in enumerate(varenumre):
        status_text.text(f"Behandler varenummer {i+1} af {total}")

        # 1) Forsøg præcist match (konverter masterdata-værdier til string)
        match = masterdata_df[masterdata_df["PRODUCT"].astype(str) == varenr]

        # 2) Hvis ingen præcis match, så kig efter "alt før ' - '" (f.eks. "25936 - All colors")
        if match.empty:
            partial_matches = masterdata_df[
                masterdata_df["PRODUCT"]
                .astype(str)
                .apply(lambda x: x.split(" - ")[0] == varenr)
            ]
            if not partial_matches.empty:
                # Hvis flere partial matches, tag den første
                match = partial_matches.iloc[[0]]

        # Hvis vi fandt noget, tag første række
        if not match.empty:
            row = match.iloc[0]
            # Map felterne i henhold til specifikationen
            product_series_name = ""  # Intet angivet
            product_name = row["PRODUCT"]  # Kolonne C
            product_item_number = varenr   # Brugerens input
            product_description = row["PRODUCT DESCRIPTION"]  # Kolonne K
            manufacturing_country = row["COUNTRY OF ORIGIN"]  # Kolonne M
            lead_time = row["LEAD TIME"]  # Kolonne AU
            if str(lead_time).strip() == "-":
                lead_time = "Ready to ship"
            warranty = row["WARRANTY"]  # Kolonne AQ
            contract_price = row["CONTRACT PRICE"]  # Kolonne AV
            list_price = f"{contract_price} DKK"
            
            results.append({
                "Product Series name": product_series_name,
                "Product Name": product_name,
                "Product item number": product_item_number,
                "Product Description": product_description,
                "Manufacturing country": manufacturing_country,
                "Lead time [weeks]": lead_time,
                "Product Guarantee period [years]": warranty,
                "List Price [your currency]": list_price
            })
        else:
            unmatched.append(varenr)

        progress_bar.progress((i+1) / total)

    # Hvis der ikke blev fundet nogen matches i det hele taget
    if not results:
        st.error("Ingen matches fundet. Tjek at de indtastede varenumre stemmer overens med data i masterdata.")
    else:
        # Åbn template-filen og indsæt data
        wb = load_workbook(filename=BytesIO(template_content))
        ws = wb.active

        # Overskrifter i række 7 => indsæt data fra række 8 og frem
        start_row = 8
        for idx, res in enumerate(results, start=start_row):
            ws[f"B{idx}"] = res["Product Series name"]
            ws[f"C{idx}"] = res["Product Name"]
            ws[f"D{idx}"] = res["Product item number"]
            ws[f"E{idx}"] = res["Product Description"]
            ws[f"F{idx}"] = res["Manufacturing country"]
            ws[f"G{idx}"] = res["Lead time [weeks]"]
            ws[f"H{idx}"] = res["Product Guarantee period [years]"]
            ws[f"I{idx}"] = res["List Price [your currency]"]

        # Gem den udfyldte fil i en BytesIO-strøm
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("Behandlingen er færdig!")
        st.download_button(
            "Download udfyldt fil",
            data=output,
            file_name="udfyldt_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    if unmatched:
        st.warning("Følgende varenumre blev ikke fundet i masterdata (hverken som eksakt match eller partial match):")
        st.text("\n".join(unmatched))
