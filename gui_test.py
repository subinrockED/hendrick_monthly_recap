import streamlit as st

st.title("My Python Script")
input_text = st.text_input("Enter some text")
if st.button("Run"):
    # Your script code here
    result = input_text.upper()  # Example transformation
    st.write("Result:", result)