import pandas as pd
import pdfplumber
import re

# Input PDF and Output Excel file paths
input_pdf = "data1.pdf"
output_excel = "output.xlsx"

# Step 1: Extract text from the PDF
def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"  # Extract text from each page
    return text

# Step 2: Parse the extracted text into a DataFrame
def parse_text_to_dataframe(text):
    lines = text.strip().split('\n')
    data = []

    for line in lines:
        if line.strip():  # Skip empty lines
            # Split the line into columns based on spaces, using regex to split by multiple spaces
            columns = re.split(r'\s{2,}', line.strip())  # Split on two or more spaces
            
            # Optionally, skip lines that don't have enough columns
            if len(columns) >= 14:  # Adjust this condition based on your needs
                data.append(columns)

    # Create a DataFrame from the parsed data
    df = pd.DataFrame(data)
    return df

# Step 3: Process the DataFrame and set proper column names
def process_dataframe(df):
    # Define the expected column names
    columns = [
        'S.No', 
        'State/UT',
        'General Elector Men', 
        'General Elector Women', 
        'General Elector Third Gender', 
        'General Elector Total',
        'NRI Elector Men', 
        'NRI Elector Women', 
        'NRI Elector Third Gender', 
        'NRI Elector Total',
        'Service Elector Men', 
        'Service Elector Women', 
        'Service Elector Total',
        'Grand Total'
    ]
    
    # Debugging: Check the shape and content of the DataFrame
    print("DataFrame shape before processing:", df.shape)
    print("DataFrame head before processing:\n", df.head())

    # Adjust the number of columns in the DataFrame to match the expected columns
    if len(df.columns) > len(columns):
        df = df.iloc[:, :len(columns)]  # Keep only the expected number of columns

    # Assign the column names
    df.columns = columns[:df.shape[1]]  # Assign based on the number of columns present

    # Convert appropriate columns to numeric, handling errors using .loc
    for col in df.columns[2:]:  # Skip S.No and State/UT
        df.loc[:, col] = pd.to_numeric(df[col], errors='coerce').fillna('')

    return df

# Step 4: Write the processed DataFrame to Excel
with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
    text = extract_text_from_pdf(input_pdf)

    # Print the extracted text for debugging
    print("Extracted text from PDF:\n", text)  # Print the raw text
    
    df = parse_text_to_dataframe(text)
    processed_df = process_dataframe(df)
    
    # Save to a single sheet, including column headers in the output
    processed_df.to_excel(writer, sheet_name='Data', index=False)

print(f"PDF successfully converted to Excel: {output_excel}")
