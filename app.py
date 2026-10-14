import streamlit as st
from MyProject import ask_system 


st.title("Student Data Analyzer")


user_input = st.text_input("Enter your query here:")



if st.button("Analyze"):
    if user_input:
        result = ask_system(user_input)
        
        
        if hasattr(result, 'savefig'): 
            st.pyplot(result)
            
        
        elif isinstance(result, str):
            if "❌" in result or "Error" in result:
                st.error(result)  
            elif "✅" in result:
                st.success(result) 
            else:
                st.write(result)   