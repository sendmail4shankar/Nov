import pdfplumber
import pandas as pd
import os

# Path to the folder containing PDFs
folder_path = r"C:\Users\EGLADMIN\Downloads\WIC4130035"

# List to store extracted data from all PDFs
all_data = []

# Loop through each file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):  # Only process PDF files
        pdf_path = os.path.join(folder_path, filename)

        # Open and read the PDF
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()  # Extract text from the page

                if text:  # If the page contains text
                    lines = text.split('\n')  # Split the text into lines

                    # Process each line to extract the details
                    for line in lines:
                        if "WIC" in line:  # Look for lines containing "WIC"
                            # Split the line into parts (adjust based on your PDF format)
                            parts = line.split()
                            try:
                                yadi_number = parts[0]
                                card_number = parts[1]
                                name = parts[2]
                                address = " ".join(parts[3:-3])  # Adjust indices if address spans multiple parts
                                age = parts[-3]
                                gender = parts[-2]

                                # Append the extracted details
                                all_data.append([filename, yadi_number, card_number, name, address, age, gender])
                            except IndexError:
                                # Skip lines that don't match the expected format
                                print(f"Skipped line in {filename}: {line}")

# Define column names for the DataFrame
columns = ["Filename", "Yadi Number", "Card Number", "Name", "Address", "Age", "Gender"]

# Convert the extracted data into a DataFrame
df = pd.DataFrame(all_data, columns=columns)

# Save the data to an Excel file
output_path = r"C:\Users\EGLADMIN\Downloads\extracted_data_from_folder.xlsx"
df.to_excel(output_path, index=False)

print(f"Data successfully extracted and saved to {output_path}")
