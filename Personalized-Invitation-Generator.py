Import pandas as pd
from docx import Document
import os

# Load the Excel file
excel_file = '/Users/namesurname/Desktop/example.xlsx'
df = pd.read_excel(excel_file)

# Check the DataFrame for names
print(df.head())  # Ensure this prints your data

# Assuming the column with names is called 'Name'
names = df['Name']

# Load the invitation letter template
template_file = '/Users/namesurname/Desktop/example.docx'
if not os.path.exists(template_file):
    print(f"Template file not found: {template_file}")

# Specify the folder to save the files
output_folder = '/Users/namesurname/Desktop/foldername'  # Update this line

# Create the folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Loop through the names and create personalized letters
for name in names:
    print(f'Creating invitation for {name}')  # Debugging line
    # Load the template
    doc = Document(template_file)
    
    # Find and replace the placeholder with the name in the document
    for paragraph in doc.paragraphs:
        if 'Dear' in paragraph.text:
            paragraph.text = paragraph.text.replace('Dear', f'Dear {name}, ')
    
    # Save each personalized invitation in the specified folder
    file_path = os.path.join(output_folder, f'invitation_for_{name}.docx')
    doc.save(file_path)

print("Personalized invitations created and saved successfully!")
