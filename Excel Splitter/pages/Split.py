import streamlit as st
import pandas as pd
import os
import zipfile

st.set_page_config(
    page_title="Excel Tools",
    page_icon="⚒️",
    layout="wide",
)

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}

            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

def main():
    st.markdown("<h1><u>Excel Splitter and Customizer</u><h1>", unsafe_allow_html=True)
    #st.write("Enter data in the table below:")

    # Create a placeholder to collect user input data
    #path = st.text_input(label='Provide path of Excel file', placeholder='E:\Python\Python Programs\Manish Project\demo.xlsx')
    data_input = st.file_uploader("Choose an Excel or csv file", type=['xlsx', 'csv'])


    if data_input:
        # Convert the Excel input to a DataFrame
        df = pd.read_excel(data_input)


        col1,col2 = st.columns(2)
        with col1:
            st.subheader('First 10 data')
            st.write(df.head(10))
        with col2:
            st.subheader('Last 10 data')
            st.write(df.tail(10))

        st.write("---")

        # Allow the user to choose how many Excel files to create
        st.markdown("<h3 style='text-align: center;'>Select the number of Excel files to create</h3>", unsafe_allow_html=True)
        num_excel_files = st.slider("", min_value=1, max_value=10, value=2)

        # Dictionary to store DataFrames for each Excel file
        excel_dataframes = {}

        # Allow the user to choose columns for each Excel file and save them to dictionary
        for i in range(num_excel_files):
            selected_columns = st.multiselect(f"Select Columns for Excel File {i+1}", df.columns)

            if selected_columns:
                # Save the selected columns to the respective DataFrame
                df_selected = df[selected_columns]

                col3,col4 = st.columns(2)

                # Allow the user to choose the sort order
                with col3:
                    sort_order = st.selectbox(f"Select Sort Order for Excel File {i + 1}", ["Ascending", "Descending"])

                # Allow the user to choose the column for sorting
                with col4:
                    sort_column = st.selectbox(f"Select Column for Sorting for Excel File {i + 1}", selected_columns)

                if sort_order == "Ascending":
                    df_selected = df_selected.sort_values(by=sort_column, ascending=True)
                else:
                    df_selected = df_selected.sort_values(by=sort_column, ascending=False)

                excel_dataframes[f"excel_file_{i + 1}.xlsx"] = df_selected
            st.write("---")

        # Create a temporary directory to store Excel files
        temp_dir = "temp_excel_files"
        os.makedirs(temp_dir, exist_ok=True)



        # Save the DataFrames to Excel files in the temporary directory
        excel_file_paths = []
        for file_name, df_selected in excel_dataframes.items():
            excel_file_path = os.path.join(temp_dir, file_name)
            df_selected.to_excel(excel_file_path, index=False)

            excel_file_paths.append(excel_file_path)
            #st.divider()

        #download_path = st.text_input("Choose a directory to save the files", value=os.path.abspath(os.path.curdir))
        # Create a zip archive containing the Excel files
        if st.button(label='Confirm'):
            # Allow the user to choose the download location

            #zip_file_path = "excel_files.zip"
            download_path = os.path.abspath(os.path.curdir)
            zip_file_path = os.path.join(download_path, "excel_files.zip")
            with zipfile.ZipFile(zip_file_path, 'w') as zipf:
                for excel_file_path in excel_file_paths:
                    zipf.write(excel_file_path, os.path.basename(excel_file_path))

            with open("excel_files.zip", "rb") as fp:
                btn = st.download_button(
                    label="Download ZIP",
                    data=fp,
                    file_name="myfile.zip",
                    mime="application/zip"
                )

            # Display a success message
            st.success("Download successful!")



if __name__ == "__main__":
    main()
