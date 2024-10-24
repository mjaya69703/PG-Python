import subprocess
import openpyxl
import argparse

def scan_domain(domain):
    try:
        result = subprocess.run(['dig', '+short', domain], capture_output=True, text=True)
        return result.stdout.strip().split('\n')  # Mengembalikan hasil sebagai daftar
    except Exception as e:
        return [str(e)]

def read_domains_from_excel(input_file):
    workbook = openpyxl.load_workbook(input_file)
    sheet = workbook.active
    domains = [cell.value for cell in sheet['A'] if cell.value]
    return domains[1:]  # Mengabaikan header

def write_results_to_excel(domains, results, output_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'DNS Scan Results'

    # Menambahkan header
    header = ['Domain']
    max_results = max(len(result) for result in results)
    header.extend([f'Hasil dig {i+1}' for i in range(max_results)])
    sheet.append(header)

    for domain, result in zip(domains, results):
        row = [domain] + result + [''] * (max_results - len(result))  # Menambahkan kolom kosong jika hasil kurang
        sheet.append(row)

    workbook.save(output_file)

def main(input_file, output_file):
    domains = read_domains_from_excel(input_file)
    results = []

    for domain in domains:
        result = scan_domain(domain)
        results.append(result)

    write_results_to_excel(domains, results, output_file)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Scan domains using dig and save results to Excel.')
    parser.add_argument('--input', required=True, help='Input Excel file containing domains to scan')
    parser.add_argument('--output', required=True, help='Output Excel file to save results')

    args = parser.parse_args()
    main(args.input, args.output)
