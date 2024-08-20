import os
import pandas as pd
import streamlit as st
import win32com.client as win32
import pythoncom

# Function to initialize COM
def initialize_com():
    try:
        pythoncom.CoInitialize()
    except Exception as e:
        st.error(f"COM initialization failed: {str(e)}")

# Streamlit app setup
st.title("Data Splitter and Email Sender")

# Upload files
input_file = st.file_uploader("Upload the Excel file to split", type=["xlsx"])
distribution_list_file = st.file_uploader("Upload the Distribution List", type=["xlsx"])

# Get output folder path
output_folder = st.text_input("Enter the output folder path")

# Columns to split by
columns_to_split = st.multiselect(
    "Select columns to split by",
    ["Sales Manager L3", "Sales Manager L2", "Sales Manager L1", "AM Team Member", "AM Team L2", "AM Team Lead"]
)

# Function to send email
def send_email(to_address, subject, body, attachment_path):
    try:
        # Initialize COM
        initialize_com()
        
        # Initialize Outlook
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to_address
        mail.Subject = subject
        mail.Body = body
        mail.Attachments.Add(attachment_path)
        mail.Send()
        return True
    except Exception as e:
        st.error(f"Failed to send email: {str(e)}")
        return False

if st.button("Split Data and Send Emails"):
    if input_file and distribution_list_file and output_folder and columns_to_split:
        # Read the input files
        df = pd.read_excel(input_file)
        distribution_df = pd.read_excel(distribution_list_file)

        # Ensure the output folder exists
        os.makedirs(output_folder, exist_ok=True)
        
        # Add a column for the flag
        distribution_df['Sent_Flag'] = 'Not Sent'

        # Split and save
        for column in columns_to_split:
            if column in df.columns:
                # Create a folder for each column
                column_folder = os.path.join(output_folder, column)
                os.makedirs(column_folder, exist_ok=True)

                # Get unique values for the column
                unique_values = df[column].unique()

                for value in unique_values:
                    # Filter the dataframe
                    filtered_df = df[df[column] == value]

                    # Create a filename
                    filename = f"{value}.xlsx"
                    file_path = os.path.join(column_folder, filename)

                    # Save the filtered dataframe to a new Excel file
                    filtered_df.to_excel(file_path, index=False)

                    # Match with the Distribution List
                    matching_rows = distribution_df[distribution_df['Designation'] == column]
                    
                    for _, row in matching_rows.iterrows():
                        email_id = row['Email_ID']
                        name = row['Name']

                        # Match the file name with Name column
                        if value == name:
                            # Create an email
                            subject = f"Attached: {value} Data ({column})"
                            body = f"Dear {value},\n\nPlease find the attached file.\n\nBest regards,\nYour Name"

                            # Send the email and update the flag
                            if send_email(email_id, subject, body, file_path):
                                distribution_df.loc[distribution_df['Name'] == value, 'Sent_Flag'] = 'Sent'
                            else:
                                distribution_df.loc[distribution_df['Name'] == value, 'Sent_Flag'] = 'Failed'

        # Save the updated Distribution List with the Sent_Flag column
        output_distribution_list_path = os.path.join(output_folder, "Distribution_list_with_flags.xlsx")
        distribution_df.to_excel(output_distribution_list_path, index=False)
        st.success(f"Emails have been sent and the distribution list has been updated at {output_distribution_list_path}")
    else:
        st.error("Please ensure all inputs are provided.")
