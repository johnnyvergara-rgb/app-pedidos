# app_simple.py
import streamlit as st

st.title("Mi primera app simple en Streamlit")
st.write("Esta es la app de ejemplo que estoy aprendiendo a desplegar 🚀")

nombre = st.text_input("Escribe tu nombre:")

if nombre:
    st.success(f"Hola {nombre}, la app funciona correctamente!")
