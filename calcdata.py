import argparse
import pandas as pd


def load_data(file_paths):
    df = pd.concat((pd.read_excel(file) for file in file_paths))
    return df


def find_top_senders(df, output_file):
    email_counts = df['Sender Email Address'].value_counts()
    top_10_senders = email_counts.head(10)

    with open(output_file, 'w') as f:
        f.write("********** TOP TEN SENDERS BEGIN **********\n")
        f.write("Top 10 senders:\n")
        for sender, count in top_10_senders.items():
            f.write(f"{sender}: {count}\n")
        f.write("********** TOP TEN SENDERS END **********\n\n\n")


def analyze_pci(df, output_file):
    pci_data = df[df['Matched Rules'].str.contains('PCI', na=False)]
    pci_count = len(pci_data)
    unique_pci_senders = pci_data['Sender Email Address'].unique()

    with open(output_file, 'a') as f:
        f.write("********** PCI BEGIN **********\n")
        f.write(f"\nPCI related emails: {pci_count}\n")
        f.write(f"Unique senders of PCI related emails: {len(unique_pci_senders)}\n\n\n")
        f.write("List of unique senders of PCI related emails with subject and Message ID:\n")
        for sender in unique_pci_senders:
            sender_data = pci_data[pci_data['Sender Email Address'] == sender]
            for _, row in sender_data.iterrows():
                f.write(f"Sender: {sender}, Subject: {row['Subject']}, Message ID: {row['Message ID']}\n")
        f.write("********** PCI END **********\n\n\n")


def quarantine_analysis(df, output_file):
    total_quarantined = len(df)
    total_released = df['Quarantined Restored Date'].notna().sum()
    total_not_released = total_quarantined - total_released

    with open(output_file, 'a') as f:
        f.write("********** QUARANTINE/RELEASE BEGIN **********\n")
        f.write(f"\nTotal emails quarantined: {total_quarantined}\n")
        f.write(f"Total emails released from quarantine: {total_released}\n")
        f.write(f"Total emails not released from quarantine: {total_not_released}\n")
        f.write("********** QUARANTINE/RELEASE END **********\n\n\n")


def analyze_quarantine_by_brand(df, output_file):
    df['Sender Domain'] = df['Sender Email Address'].str.split('@').str[1]
    quarantined_by_brand = df['Sender Domain'].value_counts()
    released_by_brand = df[df['Quarantined Restored Date'].notna()]['Sender Domain'].value_counts()

    with open(output_file, 'a') as f:
        f.write("********** QUARANTINE BY BRAND BEGIN **********\n")        
        f.write("\nEmails quarantined by brand:\n")
        f.write(quarantined_by_brand.to_string())
        f.write("\nEmails released from quarantine by brand:\n")
        f.write(released_by_brand.to_string())
        f.write("\n********** QUARANTINE BY BRAND END **********\n\n\n")    


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Analyze email data.')
    parser.add_argument('input_files', type=str, nargs='+', help='Input Excel files')
    parser.add_argument('output_file', type=str, help='Output text file')

    args = parser.parse_args()

    df = load_data(args.input_files)
    find_top_senders(df, args.output_file)
    analyze_pci(df, args.output_file)
    quarantine_analysis(df, args.output_file)
    analyze_quarantine_by_brand(df, args.output_file)
