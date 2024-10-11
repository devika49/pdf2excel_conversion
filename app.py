import pandas as pd
import pdfplumber
import re

# Input PDF, Output JSON, and Excel file paths
input_pdf = "data2.pdf"
output_json = "output.json"
output_excel = "output.xlsx"

# Step 1: Extract text from the PDF
def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:  # Check if text was extracted from the page
                text += page_text + "\n"
    
    return text if text.strip() else None

# Step 2: Parse the extracted text into a DataFrame
def parse_text_to_dataframe(text):
    if text is None:
        raise ValueError("No text found in the PDF.")
        
    lines = text.strip().split('\n')
    data = []
    start_data_found = False  # Flag to indicate when to start capturing data

    for line in lines:
        if line.strip():  # Skip empty lines
            # Check if the line starts with a serial number (digit)
            if re.match(r'^\d+', line):
                start_data_found = True  # Start capturing data from here
            
            if start_data_found:  # Only capture data if the flag is set
                # Extract all numeric data
                numeric_data = re.findall(r'\d+', line)  # Find all numeric values
                
                # Capture the state name (second column), which may include spaces and special characters
                state_name_match = re.search(r'([^\d]+)', line)  # Match everything until the first digit
                state_name = state_name_match.group(1).strip() if state_name_match else ""
                
                # Combine the state name with numeric data
                # Use only the first numeric data as a leading identifier (assuming it's 'S.No')
                s_no = numeric_data[0] if numeric_data else ''
                
                # Check if S.No is greater than 100
                if int(s_no) >= 100:
                    # Shift values down the column
                    row = {
                        'S.No': '',  # Empty for S.No > 100
                        'State/UT': state_name,
                        'General Elector Men': s_no,  # Move S.No to General Elector Men
                        'General Elector Women': numeric_data[1] if len(numeric_data) > 1 else '',
                        'General Elector Third Gender': numeric_data[2] if len(numeric_data) > 2 else '',
                        'General Elector Total': numeric_data[3] if len(numeric_data) > 3 else '',
                        'NRI Elector Men': numeric_data[4] if len(numeric_data) > 4 else '',
                        'NRI Elector Women': numeric_data[5] if len(numeric_data) > 5 else '',
                        'NRI Elector Third Gender': numeric_data[6] if len(numeric_data) > 6 else '',
                        'NRI Elector Total': numeric_data[7] if len(numeric_data) > 7 else '',
                        'Service Elector Men': numeric_data[8] if len(numeric_data) > 8 else '',
                        'Service Elector Women': numeric_data[9] if len(numeric_data) > 9 else '',
                        'Service Elector Total': numeric_data[10] if len(numeric_data) > 10 else '',
                        'Grand Total': numeric_data[11] if len(numeric_data) > 11 else '',
                        'Additional Info': None  # Add any additional info if needed
                    }
                else:
                    # Regular processing for S.No < 100
                    row = {
                        'S.No': s_no,
                        'State/UT': state_name,
                        'General Elector Men': numeric_data[1] if len(numeric_data) > 1 else '',
                        'General Elector Women': numeric_data[2] if len(numeric_data) > 2 else '',
                        'General Elector Third Gender': numeric_data[3] if len(numeric_data) > 3 else '',
                        'General Elector Total': numeric_data[4] if len(numeric_data) > 4 else '',
                        'NRI Elector Men': numeric_data[5] if len(numeric_data) > 5 else '',
                        'NRI Elector Women': numeric_data[6] if len(numeric_data) > 6 else '',
                        'NRI Elector Third Gender': numeric_data[7] if len(numeric_data) > 7 else '',
                        'NRI Elector Total': numeric_data[8] if len(numeric_data) > 8 else '',
                        'Service Elector Men': numeric_data[9] if len(numeric_data) > 9 else '',
                        'Service Elector Women': numeric_data[10] if len(numeric_data) > 10 else '',
                        'Service Elector Total': numeric_data[11] if len(numeric_data) > 11 else '',
                        'Grand Total': numeric_data[12] if len(numeric_data) > 12 else '',
                        'Additional Info': None  # Add any additional info if needed
                    }

                # Append the row data
                data.append(row)

    # Create a DataFrame from the parsed data
    df = pd.DataFrame(data)

    # Keep 'S.No' as 'TOTAL' without any changes
    df.loc[df['S.No'] == 'TOTAL', 'S.No'] = 'TOTAL'

    return df

# Step 3: Process the DataFrame and set proper column names
def process_dataframe(df):
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
        'Grand Total',
        'Additional Info'  # New column for combined text data
    ]
    
    # Adjust the number of columns in the DataFrame to match the expected columns
    if len(df.columns) > len(columns):
        df = df.iloc[:, :len(columns)]  # Keep only the expected number of columns

    # Assign the column names
    df.columns = columns[:df.shape[1]]  # Assign based on the number of columns present

    # Convert appropriate columns to numeric, handling errors using .loc
    for col in df.columns[2:-1]:  # Skip S.No, State/UT, and the last column
        df.loc[:, col] = pd.to_numeric(df[col], errors='coerce').fillna('')

    return df

# Step 4: Write the processed DataFrame to JSON and Excel
text = extract_text_from_pdf(input_pdf)

# Check if text was successfully extracted
if text is None:
    print(f"No text was extracted from the PDF: {input_pdf}")
else:
    df = parse_text_to_dataframe(text)
    processed_df = process_dataframe(df)

    # Filter out rows with invalid state names (subheadings and total)
    processed_df = processed_df[processed_df['State/UT'].str.match(r'^[A-Z ]+([& #][A-Z ]+)*$', na=False)]

    # Save DataFrame to JSON format
    processed_df.to_json(output_json, orient='records', indent=4)
    
    # Convert the JSON to Excel
    df_from_json = pd.read_json(output_json)
    
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        df_from_json.to_excel(writer, sheet_name='Data', index=False)
        
    # Print key-value pairs from the JSON
    json_dict = df_from_json.to_dict(orient='records')
    for record in json_dict:
        for key, value in record.items():
            print(f"{key}: {value}")
        print("\n" + "-"*50 + "\n")
