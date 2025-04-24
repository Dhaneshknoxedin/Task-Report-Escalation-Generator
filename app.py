import streamlit as st
from processor import process_excel

# Page setup
st.set_page_config(page_title="Escalation Task Report", layout="centered")
st.title("📊 Task Report Escalation Generator")

# File uploader
uploaded_file = st.file_uploader("Upload Raw Task Excel File", type=[".xlsx", ".xls"])

if uploaded_file is not None:
    st.success("✅ File uploaded successfully!")

    try:
        # Process the uploaded Excel file
        output, df = process_excel(uploaded_file)

        # Show preview
        st.subheader("🔍 Preview of Processed Data")
        st.dataframe(df.head(10))

        # Download button for final Excel report
        st.download_button(
            label="📥 Download Processed Report",
            data=output,
            file_name="Task_Report_Escalation_Live.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("⬆️ Please upload an Excel file to begin.")
