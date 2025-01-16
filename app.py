import streamlit as st
import pandas as pd
from scheduler import dfm  # Import the processed `dfm` from the backend
from visualisation import visualisation  # Import Gantt chart visualization
from results import results
from modify import modify
from product_list_change import product_list_change
from product_catalogue import product_catalogue

# Set page configuration
st.set_page_config(
    page_title="Machine Production Scheduler",
    page_icon="🦾",
    layout="wide"
)

# Main Title
st.title("Machine Production Scheduler")

# File Download Button
@st.cache_data
def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

csv_file = convert_df_to_csv(dfm)
st.download_button(
    label="📥 Download Current File",
    data=csv_file,
    file_name="Machine_Production_Schedule.csv",
    mime="text/csv"
)

# Add Tabs Below
tabs = st.tabs([
    "Visualisation",  
    "Modify",
    "Product List Change", 
    "Product Catalogue", 
    # "Similarity Catalogue", 
    # "Results",
    ])

# Tab Content
with tabs[0]:  # Visualisation Tab
    visualisation(dfm,st)
    
with tabs[1]:
    modify()

with tabs[2]:
    product_list_change()

with tabs[3]:
    product_catalogue()

# with tabs[4]:
#     similarity()

# with tabs[5]:
#     results()
