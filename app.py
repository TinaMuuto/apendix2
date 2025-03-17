import streamlit as st
import pandas as pd

st.title("Prislist App")
st.write("Indtast varenumre for at hente data fra masterdatafilen.")

# Udskift URL'en med den direkte raw-link til din masterdatafil på GitHub.
@st.cache_data
def load_master_data():
    url = "https://raw.githubusercontent.com/DIT_BRUGERNAVN/DIT_REPO/DIT_BRANCH/Muuto_Master_Data_CON_January_2025_DKK.xlsx"
    df = pd.read_excel(url)
    return df

master_data = load_master_data()

# Lad brugeren indsætte varenumre, ét per linje.
input_numbers = st.text_area("Indtast varenumre (én pr. linje):", "")

if st.button("Søg"):
    if input_numbers:
        # Split inputtet ved linjeskift og fjern tomme linjer
        numbers_list = [num.strip() for num in input_numbers.splitlines() if num.strip()]
        result = master_data[master_data["Varenummer"].astype(str).isin(numbers_list)]
        if not result.empty:
            st.write("Resultater:")
            st.dataframe(result)
        else:
            st.write("Ingen matchende varenumre fundet.")
    else:
        st.write("Indtast venligst mindst ét varenummer.")
