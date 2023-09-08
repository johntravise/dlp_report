import argparse
import pandas as pd
import os

def load_data(file_paths):
    dfs = []
    for file_path in file_paths:
        _, file_extension = os.path.splitext(file_path)
        if file_extension == '.csv':
            df = pd.read_csv(file_path)
        elif file_extension in ['.xls', '.xlsx']:
            df = pd.read_excel(file_path)
        else:
            raise ValueError(f"Unsupported file type: {file_extension}")
        dfs.append(df)
    return pd.concat(dfs)

def find_top_senders(df):
    email_counts = df['Sender Email Address'].value_counts()
    top_10_senders = email_counts.head(10)
    return top_10_senders.reset_index().rename(columns={'index': 'Sender Email Address', 'Sender Email Address': 'Count'})

def analyze_pci(df):
    pci_data = df[df['Matched Rules'].str.contains('PCI', na=False)]
    pci_count = len(pci_data)
    unique_pci_senders = pci_data['Sender Email Address'].unique()

    result = pd.DataFrame(columns=['Sender', 'Subject', 'Found Text'])
    for sender in unique_pci_senders:
        sender_data = pci_data[pci_data['Sender Email Address'] == sender]
        for _, row in sender_data.iterrows():
            result = pd.concat([result, pd.DataFrame({'Sender': [sender], 'Subject': [row['Subject']], 'Message ID': [row['Message ID']]})], ignore_index=True)
    
    return result

def quarantine_analysis(df):
    total_quarantined = len(df)
    total_released = df['Quarantined Restored Date'].notna().sum()
    total_not_released = total_quarantined - total_released

    result = pd.DataFrame({
        'Total Quarantined': [total_quarantined],
        'Total Released': [total_released],
        'Total Not Released': [total_not_released]
    })
    
    return result

def analyze_quarantine_by_brand(df):
    df['Sender Domain'] = df['Sender Email Address'].str.split('@').str[1]
    quarantined_by_brand = df['Sender Domain'].value_counts()
    released_by_brand = df[df['Quarantined Restored Date'].notna()]['Sender Domain'].value_counts()

    result = pd.concat([quarantined_by_brand, released_by_brand], axis=1, keys=['Quarantined', 'Released'])
    return result.reset_index().rename(columns={'index': 'Brand'})

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Analyze email data.')
    parser.add_argument('files', type=str, nargs='+', help='One or more input Excel or CSV files followed by the output Excel file')

    args = parser.parse_args()

    # Split input and output files
    input_files = args.files[:-1]
    output_file = args.files[-1]

    df = load_data(input_files)

    with pd.ExcelWriter(output_file) as writer:
        find_top_senders(df).to_excel(writer, sheet_name='Top Senders', index=False)
        analyze_pci(df).to_excel(writer, sheet_name='PCI Analysis', index=False)
        quarantine_analysis(df).to_excel(writer, sheet_name='Quarantine Analysis', index=False)
        analyze_quarantine_by_brand(df).to_excel(writer, sheet_name='Quarantine by Brand', index=False)


def analyze_by_department(df):
    dept_counts = df['Owner Department'].value_counts()
    return dept_counts.reset_index().rename(columns={'index': 'Department', 'Owner Department': 'Count'})

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Analyze email data.')
    parser.add_argument('files', type=str, nargs='+', help='One or more input Excel or CSV files followed by the output Excel file')

    args = parser.parse_args()

    # Split input and output files
    input_files = args.files[:-1]
    output_file = args.files[-1]

    df = load_data(input_files)
    
    with pd.ExcelWriter(output_file) as writer:
        find_top_senders(df).to_excel(writer, sheet_name='Top Senders', index=False)
        analyze_pci(df).to_excel(writer, sheet_name='PCI Analysis', index=False)
        quarantine_analysis(df).to_excel(writer, sheet_name='Quarantine Analysis', index=False)
        analyze_quarantine_by_brand(df).to_excel(writer, sheet_name='Quarantine by Brand', index=False)
        analyze_by_department(df).to_excel(writer, sheet_name='Department Analysis', index=False)
