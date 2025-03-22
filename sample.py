# Example usage of psimport.py - For Palmsense PStouch data conversion

import os
from psimport import VoltammetryImporter

def main():
    """Example of using the VoltammetryImporter class"""
    # Define the path to the input CSV file
    input_file = "seconde misure.csv"
    
    # Create a directory for results if it doesn't already exist
    output_dir = "analysis_results"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Initialize the importer
    importer = VoltammetryImporter(input_file)
    
    # Load the data
    print(f"Loading file {input_file}...")
    if not importer.load_file():
        print("Error loading file. Exiting.")
        return
    
    # Show information about loaded scans
    scan_count = importer.get_scan_count()
    scan_names = importer.get_scan_names()
    
    print(f"Loaded {scan_count} scans:")
    for i, name in enumerate(scan_names):
        scan_data = importer.get_scan_data(i)
        points = len(scan_data['potential'])
        print(f"  {i}: {name} ({points} points)")
    
    # Export all scans to Excel
    excel_file = os.path.join(output_dir, "complete_voltammetry.xlsx")
    print(f"Exporting to Excel: {excel_file}")
    if importer.export_to_excel(excel_file):
        print("  Export completed successfully!")
    else:
        print("  Error exporting to Excel")
    
    # Export each scan individually to CSV format
    print("Exporting each scan to CSV...")
    for i in range(scan_count):
        csv_file = os.path.join(output_dir, f"scan_{i+1}.csv")
        if importer.export_to_csv(csv_file, scan_index=i):
            print(f"  Scan {i+1} exported to: {csv_file}")
        else:
            print(f"  Error exporting scan {i+1}")
    
    # Export the first scan to TXT format
    txt_file = os.path.join(output_dir, "scan_1.txt")
    print(f"Exporting to TXT: {txt_file}")
    if importer.export_to_txt(txt_file, scan_index=0):
        print("  Export completed successfully!")
    else:
        print("  Error exporting to TXT")
    
    # Export the first scan to CHI format
    chi_file = os.path.join(output_dir, "scan_1.chi")
    print(f"Exporting to CHI format: {chi_file}")
    if importer.export_to_chi(chi_file, scan_index=0):
        print("  Export completed successfully!")
    else:
        print("  Error exporting to CHI format")
    
    print("\nAll operations completed!")

if __name__ == "__main__":
    main()