import pandas as pd
import sys

# Cleanup function to remove empty rows and identify the header row
def clean_up_excel(df):
    for index, row in df.iterrows():
        if not row.isnull().all():  # Found a row where not all cells are empty
            df = df.loc[index:]
            df.columns = df.iloc[0]  # Set the first non-empty row as header
            df = df[1:]  # Remove the header row from the data
            break
    return df

# Function to read and clean the Excel file
def read_and_clean_excel(input_file):
    df = pd.read_excel(input_file, header=None)  # Load without assuming any header
    df_cleaned = clean_up_excel(df)
    return df_cleaned

# Function to process the data
def process_data(df):
    # Filter data where 'IS Seriennummer' == 1
    df_filtered = df[df['IS Seriennummer'] == 1]
    
    # Calculate total sum by customer
    total_sum_by_customer = df.groupby('Other Ref2')['Charge (eur)'].sum().reset_index(name='Gesamtsumme')
    
    # Calculate module count by customer
    module_count_by_customer = df_filtered.groupby('Other Ref2')['Qty'].sum().reset_index(name='Anzahl der Module')
    
    # Merge total sum and module count
    merged_data = pd.merge(total_sum_by_customer, module_count_by_customer, on='Other Ref2', how='left').fillna(0)
    
    # Categorization based on module count
    def categorize(row):
        modules = row['Anzahl der Module']
        if 1 <= modules <= 7:
            return 'Kategorie 1'
        elif 8 <= modules <= 20:
            return 'Kategorie 2'
        elif 21 <= modules <= 30:
            return 'Kategorie 3'
        elif 31 <= modules <= 40:
            return 'Kategorie 4'
        elif 41 <= modules <= 60:
            return 'Kategorie 5'
        else:
            return 'Nicht kategorisiert'

    merged_data['Durchschnittliche Kosten DC'] = merged_data.apply(categorize, axis=1)
    
    return merged_data

# Function for basic validation (e.g., check for missing values)
def validate_data(df):
    validation_results = {}
    required_columns = ['IS Seriennummer', 'Other Ref2', 'Charge (eur)', 'Qty']
    for column in required_columns:
        if column not in df.columns:
            validation_results[column] = 'Missing'
        else:
            # Check for rows with missing values in this column
            missing_values = df[column].isnull().sum()
            validation_results[column] = f'Missing Rows: {missing_values}'
    return pd.DataFrame(list(validation_results.items()), columns=['Column', 'Validation Result'])

# Main function to run the app
def run_app(input_file, output_file):
    df = read_and_clean_excel(input_file)
    processed_data = process_data(df)
    validation_results = validate_data(df)
    
    # Write processed data and validation results to a new Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        processed_data.to_excel(writer, sheet_name='Processed Data', index=False)
        validation_results.to_excel(writer, sheet_name='Validation Results', index=False)

# Check if the script is running directly (not being imported)
if __name__ == "__main__":
    # Check if an argument was provided
    if len(sys.argv) != 2:
        print("Usage: python excel_calculater.py <input_excel_file>")
        sys.exit(1)

    # Get the input file from the command line arguments
    input_file = sys.argv[1]
    output_file = input_file.rsplit('.', 1)[0] + '_bearbeitet.xlsx'

    # Run the app
    run_app(input_file, output_file)