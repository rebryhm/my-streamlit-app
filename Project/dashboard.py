import streamlit as st
import pandas as pd

# Load data
@st.cache_data
def load_data(file_name):
    return pd.read_excel(file_name)

# File path
file_path = "Random_Model_Names_Dimensions.xlsx"
data = load_data(file_path)

# Streamlit App
st.title("Model Details Dashboard")
st.write("Search and compare model details from the Excel file.")

# Display full dataset
if st.checkbox("Show Full Dataset"):
    st.write(data)

# Search for a single model
st.header("Search for a Model")
model_name = st.text_input("Enter the Model Name")
if model_name:
    filtered_data = data[data["Model Name"].str.contains(model_name, case=False, na=False)]
    if not filtered_data.empty:
        st.write(f"Details for '{model_name}':")
        st.write(filtered_data)
    else:
        st.error(f"No details found for model '{model_name}'.")

# Compare two models
st.header("Compare Two Models")
col1, col2 = st.columns(2)

with col1:
    model1 = st.selectbox("Select First Model", options=data["Model Name"].dropna().unique(), key="model1")
    model1_details = data[data["Model Name"] == model1]

with col2:
    model2 = st.selectbox("Select Second Model", options=data["Model Name"].dropna().unique(), key="model2")
    model2_details = data[data["Model Name"] == model2]

st.subheader(f"Comparison of {model1} and {model2}")
col1, col2 = st.columns(2)

with col1:
    st.write(f"Details for {model1}:")
    st.write(model1_details)

with col2:
    st.write(f"Details for {model2}:")
    st.write(model2_details)

# Create a button to download the comparison sheet
if st.button("Get Comparison Sheet"):
    # Combine details for both models into a column-wise format
    comparison_sheet = pd.concat([model1_details.T, model2_details.T], axis=1)
    comparison_sheet.columns = [model1, model2]  # Set column names for clarity

    # Convert to Excel and provide download link
    comparison_sheet_file = "comparison_sheet.xlsx"
    comparison_sheet.to_excel(comparison_sheet_file, index=True)
    with open(comparison_sheet_file, "rb") as file:
        st.download_button(
            label="Download Comparison Sheet",
            data=file,
            file_name="comparison_sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
