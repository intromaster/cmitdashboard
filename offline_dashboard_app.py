
import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Load data
@st.cache_data
def load_data():
    return pd.read_excel("offline_computers_by_location.xlsx")

df = load_data()

# Sidebar filters
st.sidebar.title("Filters")
companies = df["Company Name"].unique().tolist()
selected_companies = st.sidebar.multiselect("Select Companies", companies, default=companies)

# Filtered Data
filtered_df = df[df["Company Name"].isin(selected_companies)]

# Main title
st.title("Offline Computers Dashboard")
st.markdown("Showing computers offline for more than 30 days, grouped by location and company.")

# Chart
fig = px.bar(
    filtered_df,
    x="Location",
    y="Offline Devices",
    color="Company Name",
    title="Computers Offline >30 Days by Location",
    text="Offline Devices",
    barmode="group"
)
fig.update_traces(textposition='outside')
fig.update_layout(xaxis_tickangle=-45)
st.plotly_chart(fig)

# Download filtered data using in-memory buffer
buffer = io.BytesIO()
filtered_df.to_excel(buffer, index=False, engine='openpyxl')
buffer.seek(0)

st.download_button(
    label="Download Filtered Data as Excel",
    data=buffer,
    file_name="filtered_offline_computers.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
