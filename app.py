import streamlit as st
import pandas as pd

# Ensure openpyxl is installed
try:
    import openpyxl
except ImportError:
    st.error("Missing optional dependency 'openpyxl'. Use pip or conda to install openpyxl.")
    st.stop()

# Load the Excel file
file_path = '/Users/macbook/PycharmProjects/pythonProjectcarzim/car.xlsx'
car_data = pd.read_excel(file_path)

# # Display the column names for debugging
# st.write("### Column Names")
# st.write(car_data.columns)

# Title of the app
st.title("Car Data Search, Filter, and Comparison")

# Display the dataframe with a bordered style
st.markdown("""
<style>
    .dataframe {
        border: 2px solid #4CAF50;
    }
</style>
""", unsafe_allow_html=True)
st.write("## Car Data")
st.dataframe(car_data)

# Search feature with separation and bordered style
st.markdown("---")
st.write("## Search")
search_column = st.selectbox("Select column to search", car_data.columns, key='search_col')
search_term = st.text_input("Enter search term", key='search_term')
if search_term:
    search_results = car_data[car_data[search_column].astype(str).str.contains(search_term, case=False, na=False)]
    st.write(f"### Search results for '{search_term}' in '{search_column}'")
    st.dataframe(search_results)

# Filter feature with separation and bordered style
st.markdown("---")
st.write("## Filter")
filter_column = st.selectbox("Select column to filter", car_data.columns, key='filter_col')
unique_values = car_data[filter_column].dropna().unique()
filter_values = st.multiselect(f"Select values to filter '{filter_column}'", unique_values, key='filter_vals')
if filter_values:
    filter_results = car_data[car_data[filter_column].isin(filter_values)]
    st.write(f"### Filter results for '{filter_column}' in {filter_values}")
    st.dataframe(filter_results)

# Match feature with separation and bordered style
st.markdown("---")
st.write("## Match")
match_columns = st.multiselect("Select columns to match", car_data.columns, key='match_cols')
if match_columns:
    match_value = st.text_input("Enter value to match", key='match_val')
    if match_value:
        match_results = car_data[car_data[match_columns].astype(str).apply(lambda row: match_value in row.values, axis=1)]
        st.write(f"### Match results for '{match_value}' in columns {match_columns}")
        st.dataframe(match_results)

# Comparison and discrepancy analysis
st.markdown("---")
st.write("## Comparison and Discrepancy Analysis")

# Columns to compare
compare_cols = [
    "C",
    "E",
    "F",
    "H"
]

# Use the exact column names from the file
selected_compare_col = st.selectbox("Select column to compare", compare_cols, key='compare_col')

# Show discrepancies between selected columns
st.write(f"### Discrepancies in '{selected_compare_col}'")
try:
    discrepancies = car_data[car_data["C"] != car_data["E"]]
    st.dataframe(discrepancies)
except KeyError:
    st.error("Please ensure the column names are correct and match the ones in the Excel file.")

# Display mismatched names
st.write("### Mismatched Names")
try:
    mismatched_names = car_data[
        (car_data["C"] != car_data["E"]) |
        (car_data["F"] != car_data["H"])
    ]
    st.dataframe(mismatched_names)
except KeyError:
    st.error("Please ensure the column names are correct and match the ones in the Excel file.")

# Discrepancies in Facturiers and Payments
st.write("### Discrepancies in Facturiers and Payments")
try:
    facture_discrepancies = car_data[car_data["I"].isnull() | car_data["Q"].isnull()]
    st.dataframe(facture_discrepancies)
except KeyError:
    st.error("Please ensure the column names are correct and match the ones in the Excel file.")

# Discrepancies in Payments
st.write("### Discrepancies in Payments")
try:
    payment_discrepancies = car_data[car_data["R"] == 'NON']
    st.dataframe(payment_discrepancies)
except KeyError:
    st.error("Please ensure the column names are correct and match the ones in the Excel file.")

# Run the Streamlit app
# Use the command below to run this script
# streamlit run app.py
