#Home Page for the streamlit website for the Exam Timetabling System
import streamlit as st
st.set_page_config(page_title="Exam Timetabling System", layout="wide")
st.title("Welcome to the Exam Timetabling System")
st.markdown("""
Use the sidebar to:
- **Generate Timetable**: Upload your data and generate a new exam timetable.
- **Check Timetable**: Upload a timetable file to check for constraint violations.
""") 