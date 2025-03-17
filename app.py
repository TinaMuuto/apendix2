import streamlit as st
import pandas as pd

st.title("Prislist App")
st.write("Indtast varenumre for at hente data fra masterdatafilen.")

# Funktion til at indlæse masterdata med caching, så filen ikke genindlæses hver gang.
@st.cache
def load_master_data():
    # Sørg for, at filen 'Muuto_Master_Data_CON_January_2025_DKK.xlsx' ligger i samme mappe som appen,
    # eller angiv den fulde sti.
    df = pd.read_excel("Muuto_Master_Data_CON_January_2025_DKK.xlsx")
    return df

# Indlæs masterdata
master_data = load_master_data()

# Giv mulighed for at uploade en masterdatafil, hvis ønsket:
uploaded_file = st.file_uploader("Upload masterdatafilen (valgfrit):", type=["xlsx"])
if uploaded_file is not None:
    master_data = pd.read_excel(uploaded_file)

# Indtast varenumre (adskilt med komma)
input_numbers = st.text_area("Indtast varenumre (komma-separeret):", "")

if st.button("Søg"):
    if input_numbers:
        # Omdan input til en liste og fjern eventuelle ekstra mellemrum
        numbers_list = [num.strip() for num in input_numbers.split(",")]
        # Filtrer masterdata; her antages det, at kolonnen med varenumre hedder "Varenummer"
        result = master_data[master_data["Varenummer"].astype(str).isin(numbers_list)]
        if not result.empty:
            st.write("Resultater:")
            st.dataframe(result)
        else:
            st.write("Ingen matchende varenumre fundet.")
    else:
        st.write("Indtast venligst mindst ét varenummer.")

