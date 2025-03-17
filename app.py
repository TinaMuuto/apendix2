import streamlit as st
import pandas as pd
import requests
from io import BytesIO
from openpyxl import load_workbook

st.title("Prislist App")
st.write("Indtast varenumre (én pr. linje) for at hente data fra masterdatafilen og udfylde template-filen.")

# Hent masterdatafilen fra GitHub via requests og BytesIO
@st.cache_data
def load_master_data():
    master_url = "https://raw.githubusercontent.com/DIT_BRUGERNAVN/DIT_REPO/DIT_BRANCH/Muuto_Master_Data_CON_January_2025_DKK.xlsx"
    response = requests.get(master_url)
    response.raise_for_status()  # Tjek for fejl
    df = pd.read_excel(BytesIO(response.content))
    return df

# Hent template-filen fra GitHub via requests og BytesIO
@st.cache_data
def load_template():
    template_url = "https://raw.githubusercontent.com/DIT_BRUGERNAVN/DIT_REPO/DIT_BRANCH/Price-template.xlsx"
    response = requests.get(template_url)
    response.raise_for_status()
    wb = load_workbook(filename=BytesIO(response.content))
    return wb

# Indlæs data
master_data = load_master_data()
template_wb = load_template()
template_ws = template_wb.active

# Input: Én varenummer per linje
input_numbers = st.text_area("Indtast varenumre (én pr. linje):", "")

if st.button("Udfyld Template"):
    if input_numbers:
        # Lav en liste af varenumre og fjern evt. tomme linjer
        numbers_list = [num.strip() for num in input_numbers.splitlines() if num.strip()]
        
        # Filtrer masterdata ud fra kolonnen "Varenummer" (sørg for at alle værdier behandles som strenge)
        filtered_data = master_data[master_data["Varenummer"].astype(str).isin(numbers_list)]
        
        if not filtered_data.empty:
            # Forvent, at række 8 i templaten indeholder kolonneoverskrifter
            header_row = [cell.value for cell in template_ws[8]]
            start_row = 9  # Data indsættes fra række 9

            # Ryd evt. tidligere data (hvis der er indhold under række 8)
            max_row = template_ws.max_row
            if max_row >= start_row:
                for row in template_ws.iter_rows(min_row=start_row, max_row=max_row):
                    for cell in row:
                        cell.value = None

            # Udfyld templaten baseret på header-match: for hver række i de filtrerede data
            for i, row_data in enumerate(filtered_data.itertuples(index=False), start=start_row):
                for j, header in enumerate(header_row, start=1):
                    if header in filtered_data.columns:
                        value = getattr(row_data, header)
                    else:
                        value = ""
                    template_ws.cell(row=i, column=j, value=value)

            st.success("Template er udfyldt. Download den udfyldte fil nedenfor.")
            
            # Gem den opdaterede workbook til en BytesIO-buffer og giv mulighed for download
            output = BytesIO()
            template_wb.save(output)
            output.seek(0)
            
            st.download_button(
                label="Download udfyldt template",
                data=output,
                file_name="Udfyldt_Price_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Ingen matchende varenumre fundet i masterdatafilen.")
    else:
        st.warning("Indtast venligst mindst ét varenummer.")
