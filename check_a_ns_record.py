import subprocess
import openpyxl
import argparse

def scan_domain(domain):
    try:
        # Querying for multiple DNS record types
        result_a = subprocess.run(['dig', '+short', 'A', domain], capture_output=True, text=True)
        result_ns = subprocess.run(['dig', '+short', 'NS', domain], capture_output=True, text=True)
        
        # Collecting A records
        a_records = result_a.stdout.strip().split('\n') if result_a.stdout.strip() else []
        
        # Collecting NS records
        ns_records = result_ns.stdout.strip().split('\n') if result_ns.stdout.strip() else []
        
        return a_records, ns_records, result_a.stdout.strip().split('\n') + result_ns.stdout.strip().split('\n')
    except Exception as e:
        return str(e), [], []  # If an error occurs, return the error message

def read_domains_from_excel(input_file):
    workbook = openpyxl.load_workbook(input_file)
    sheet = workbook.active
    domains = [cell.value for cell in sheet['A'] if cell.value]
    return domains[1:]  # Skip header row

def write_results_to_excel(domains, a_records, ns_records, results, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'DNS Scan Results'

    # Add header row with dynamic columns for A Record and multiple NS Records
    header = ['Domain', 'A Record']
    max_ns_records = max(len(ns) for ns in ns_records)
    header.extend([f'NS Record {i+1}' for i in range(max_ns_records)])
    sheet.append(header)

    # Write data rows
    for domain, a_record, ns_record, result in zip(domains, a_records, ns_records, results):
        row = [domain, ', '.join(a_record)]  # A record may have multiple IPs
        row.extend(ns_record)  # Add all NS records for each domain
        sheet.append(row)

    # Save to output file
    workbook.save(output_file)

def main(input_file, output_file):
    domains = read_domains_from_excel(input_file)
    a_records = []
    ns_records = []
    results = []

    for domain in domains:
        a_record, ns_record, result = scan_domain(domain)
        a_records.append(a_record)
        ns_records.append(ns_record)
        results.append(result)

    write_results_to_excel(domains, a_records, ns_records, results, output_file)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Scan domains using dig and save results to Excel.')
    parser.add_argument('--input', required=True, help='Input Excel file containing domains to scan')
    parser.add_argument('--output', required=True, help='Output Excel file to save results')

    args = parser.parse_args()
    main(args.input, args.output)
