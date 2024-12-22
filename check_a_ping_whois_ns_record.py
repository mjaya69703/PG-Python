import subprocess
import openpyxl
import argparse
import re

def scan_domain(domain):
    try:
        # Querying for multiple DNS record types
        result_a = subprocess.run(['dig', '+short', 'A', domain], capture_output=True, text=True)
        result_ns = subprocess.run(['dig', '+short', 'NS', domain], capture_output=True, text=True)
        
        # Collecting A records
        a_records = result_a.stdout.strip().split('\n') if result_a.stdout.strip() else []
        
        # Collecting NS records
        ns_records = result_ns.stdout.strip().split('\n') if result_ns.stdout.strip() else []
        
        # Perform ping for each A record and extract hostname and IP
        ping_results = []
        for ip in a_records:
            ping_result = subprocess.run(['ping', '-c', '1', ip], capture_output=True, text=True)
            if ping_result.returncode == 0:
                # Extracting the hostname and IP address from the ping output
                match = re.search(r'from (.+?) $$(\d+\.\d+\.\d+\.\d+)$$', ping_result.stdout)
                if match:
                    hostname = match.group(1)
                    ping_results.append(f"{hostname} ({ip})")
                else:
                    ping_results.append(f"{ip} (not reachable)")
            else:
                ping_results.append(f"{ip} (not reachable)")

        # Perform WHOIS lookup
        whois_result = subprocess.run(['whois', domain], capture_output=True, text=True)
        whois_output = whois_result.stdout
        # Extracting status information from WHOIS output
        status_match = re.search(r'Status:\s*(.+)', whois_output)
        whois_status = status_match.group(1).strip() if status_match else "N/A"

        return a_records, ns_records, ping_results, whois_status, result_a.stdout.strip().split('\n') + result_ns.stdout.strip().split('\n')
    except Exception as e:
        return [], [], [], str(e), []  # If an error occurs, return the error message

def read_domains_from_excel(input_file):
    workbook = openpyxl.load_workbook(input_file)
    sheet = workbook.active
    domains = [cell.value for cell in sheet['A'] if cell.value]
    return domains[1:]  # Skip header row

def write_results_to_excel(domains, a_records, ns_records, ping_results, whois_statuses, results, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'DNS Scan Results'

    # Add header row with dynamic columns for A Record, Ping Result, WHOIS Status, and multiple NS Records
    header = ['Domain', 'A Record', 'Ping Result', 'WHOIS Status']
    max_ns_records = max(len(ns) for ns in ns_records)
    header.extend([f'NS Record {i+1}' for i in range(max_ns_records)])
    sheet.append(header)

    # Write data rows
    for domain, a_record, ping_result, whois_status, ns_record, result in zip(domains, a_records, ping_results, whois_statuses, ns_records, results):
        row = [domain, ', '.join(a_record), ', '.join(ping_result), whois_status]  # A record may have multiple IPs
        row.extend(ns_record)  # Add all NS records for each domain
        sheet.append(row)

    # Save to output file
    workbook.save(output_file)

def main(input_file, output_file):
    domains = read_domains_from_excel(input_file)
    a_records = []
    ns_records = []
    ping_results = []
    whois_statuses = []
    results = []

    for domain in domains:
        a_record, ns_record, ping_result, whois_status, result = scan_domain(domain)
        a_records.append(a_record)
        ns_records.append(ns_record)
        ping_results.append(ping_result)
        whois_statuses.append(whois_status)
        results.append(result)

    write_results_to_excel(domains, a_records, ns_records, ping_results, whois_statuses, results, output_file)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Scan domains using dig and save results to Excel.')
    parser.add_argument('--input', required=True, help='Input Excel file containing domains to scan')
    parser.add_argument('--output', required=True, help='Output Excel file to save results')

    args = parser.parse_args()
    main(args.input, args.output)
