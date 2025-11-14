import streamlit as st

def render():
    st.header("Linh")
    st.info("Container for Linh's charts and tables.")
    with st.container():
        st.write("Placeholder for Linh")
        placeholder = st.empty()
    return {"placeholder": placeholder}