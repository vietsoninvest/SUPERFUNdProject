import streamlit as st

def render():
    st.header("Son")
    st.info("Container for Son's charts and tables.")
    with st.container():
        st.write("Placeholder for Son")
        placeholder = st.empty()
    return {"placeholder": placeholder}