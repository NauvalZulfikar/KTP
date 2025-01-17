import pandas as pd
import streamlit as st
import datetime as dt
from scheduler import dfm, calculate_machine_utilization, component_waiting_df, product_waiting_df, late_products

# Use `dfm` from scheduler.py and ensure connection
if "dfm" in st.session_state:
    dfm = st.session_state.dfm  # Store dfm in session state

def product_catalogue():
    df_list = ['Order Processing Date', 'Promised Delivery Date', 'Start Time', 'End Time']

    # Use a temporary DataFrame for display purposes
    display_df = dfm.drop(columns=['Status', 'wait_time', 'legend'], errors='ignore')

    # Format date columns in `display_df` only for display
    for col in df_list:
        if col in display_df.columns and pd.api.types.is_datetime64_any_dtype(display_df[col]):
            display_df[col] = display_df[col].dt.strftime('%Y-%m-%d %H:%M')

    # Display the DataFrame
    st.write(display_df.sort_values(by=['Start Time', 'End Time']))

    st.subheader("Production Scheduling Results")

    # Create two columns
    col1, col2 = st.columns(2)

    # Machine Utilization in the first column
    with col1:
        st.subheader("Machine Utilization")
        if "machine_utilization_df" not in st.session_state:
            st.session_state.machine_utilization_df = calculate_machine_utilization(st.session_state.dfm)
        st.write(st.session_state.machine_utilization_df)

        st.subheader("Component Waiting Time")
        if "component_waiting_df" not in st.session_state:
            st.session_state.component_waiting_df = component_waiting_df
        st.write(st.session_state.component_waiting_df)

    # Product Waiting Time and Late Products in the second column
    with col2:
        st.subheader("Late Products")
        if "late_products_df" not in st.session_state:
            st.session_state.late_products_df = late_products(st.session_state.dfm)
        st.write(st.session_state.late_products_df)

        st.subheader("Product Waiting Time")
        if "product_waiting_df" not in st.session_state:
            st.session_state.product_waiting_df = product_waiting_df
        st.write(st.session_state.product_waiting_df)
