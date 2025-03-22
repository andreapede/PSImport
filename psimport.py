#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
psimport.py - Script for importing cyclic voltammetry data from Palmsense PStouch software
Specifically designed for converting CSV files exported from Palmsense PStouch software
Supports CSV files with UTF-16 encoding in the format of "seconde misure.csv"
Can export to various formats compatible with electrochemical analysis software
"""

import os
import sys
import argparse
import numpy as np
import pandas as pd
from datetime import datetime

class VoltammetryImporter:
    def __init__(self, input_file=None, encoding='utf-16'):
        """
        Initialize the cyclic voltammetry data importer for Palmsense PStouch files
        
        Args:
            input_file (str): Path to the CSV file exported from Palmsense PStouch to import
            encoding (str): File encoding (default: utf-16)
        """
        self.input_file = input_file
        self.encoding = encoding
        self.scans = []
        self.metadata = {}
        self.dates = []
        
    def load_file(self, input_file=None):
        """
        Load a CSV file containing cyclic voltammetry data
        
        Args:
            input_file (str, optional): Path to the CSV file to load
            
        Returns:
            bool: True if loading succeeded, False otherwise
        """
        if input_file:
            self.input_file = input_file
            
        if not self.input_file or not os.path.exists(self.input_file):
            print(f"Error: File {self.input_file} not found")
            return False
        
        try:
            # Read the file with the specified encoding
            with open(self.input_file, 'r', encoding=self.encoding) as f:
                lines = f.readlines()
            
            # Extract basic information
            if len(lines) > 0:
                header_parts = lines[0].strip().split(',')
                if len(header_parts) >= 2:
                    self.metadata['date_time'] = header_parts[1].strip()
            
            # Look for scan names and measurement dates
            scan_headers = []
            measurement_dates = []
            
            for i, line in enumerate(lines[:10]):  # Check only the first few lines
                if "Cyclic Voltammetry" in line:
                    scan_headers = line.strip().split(',')
                elif "Date and time measurement:" in line:
                    measurement_dates = line.strip().split(',')
            
            # Extract scan names
            scan_names = []
            for part in scan_headers:
                part = part.strip()
                if part and "Cyclic Voltammetry" in part:
                    scan_names.append(part)
            
            # Extract measurement dates
            for part in measurement_dates:
                part = part.strip()
                if part and "Date and time measurement:" not in part:
                    try:
                        # Attempt to convert the string to a datetime object
                        date_obj = datetime.strptime(part, "%Y-%m-%d %H:%M:%S")
                        self.dates.append(date_obj)
                    except (ValueError, TypeError):
                        # Ignore if not a valid date
                        pass
            
            # Look for the row with column headers (V, µA)
            column_header_index = -1
            for i, line in enumerate(lines[:20]):  # Check only the first few lines
                if 'V' in line and 'µA' in line:
                    column_header_index = i
                    break
            
            if column_header_index == -1:
                print("Error: Unable to find column headers (V, µA)")
                return False
                
            # Get column headers
            column_headers = lines[column_header_index].strip().split(',')
            scan_count = len(column_headers) // 2  # Each scan has two columns (V and µA)
            
            # Prepare data for each scan
            self.scans = []
            for i in range(scan_count):
                scan_name = scan_names[i] if i < len(scan_names) else f"Scan {i+1}"
                scan_date = self.dates[i] if i < len(self.dates) else None
                
                self.scans.append({
                    'name': scan_name,
                    'date': scan_date,
                    'potential': [],  # V values
                    'current': [],    # µA values
                    'metadata': {}
                })
            
            # Parse data rows (starting from the row after the headers)
            data_start_index = column_header_index + 1
            
            for line_index in range(data_start_index, len(lines)):
                line = lines[line_index].strip()
                if not line:
                    continue
                
                values = line.split(',')
                
                # Process each data point for each scan
                for scan_index in range(scan_count):
                    potential_index = scan_index * 2
                    current_index = potential_index + 1
                    
                    if potential_index < len(values) and current_index < len(values):
                        try:
                            potential = float(values[potential_index])
                            current = float(values[current_index])
                            
                            self.scans[scan_index]['potential'].append(potential)
                            self.scans[scan_index]['current'].append(current)
                        except ValueError:
                            # Ignore non-numeric values
                            pass
            
            # Verify that we actually loaded data
            if not self.scans or all(len(scan['potential']) == 0 for scan in self.scans):
                print("Error: No valid data found in the file")
                return False
                
            # Convert lists to NumPy arrays for better efficiency
            for scan in self.scans:
                scan['potential'] = np.array(scan['potential'])
                scan['current'] = np.array(scan['current'])
                
                # Add useful metadata
                scan['metadata']['points'] = len(scan['potential'])
                if len(scan['potential']) > 0:
                    scan['metadata']['potential_range'] = [
                        float(np.min(scan['potential'])), 
                        float(np.max(scan['potential']))
                    ]
                    scan['metadata']['current_range'] = [
                        float(np.min(scan['current'])), 
                        float(np.max(scan['current']))
                    ]
            
            return True
                
        except Exception as e:
            print(f"Error loading file: {str(e)}")
            return False
    
    def get_scan_count(self):
        """Returns the number of loaded scans"""
        return len(self.scans)
    
    def get_scan_names(self):
        """Returns the list of scan names"""
        return [scan['name'] for scan in self.scans]
    
    def get_scan_data(self, index):
        """
        Returns data for a specific scan
        
        Args:
            index (int): Scan index (0-based)
            
        Returns:
            dict: Dictionary with scan data
        """
        if 0 <= index < len(self.scans):
            return self.scans[index]
        return None
    
    def export_to_csv(self, output_file, scan_index=None):
        """
        Exports data from one or all scans to CSV format
        
        Args:
            output_file (str): Path to the output file
            scan_index (int, optional): Index of the scan to export.
                                       If None, exports all scans.
        
        Returns:
            bool: True if export succeeded, False otherwise
        """
        try:
            if scan_index is not None:
                # Export a single scan
                if 0 <= scan_index < len(self.scans):
                    scan = self.scans[scan_index]
                    df = pd.DataFrame({
                        'Potential (V)': scan['potential'],
                        'Current (µA)': scan['current']
                    })
                    df.to_csv(output_file, index=False)
                    return True
                else:
                    print(f"Error: Invalid scan index {scan_index}")
                    return False
            else:
                # Export all scans to a single CSV
                all_data = {}
                for i, scan in enumerate(self.scans):
                    scan_name = f"Scan_{i+1}"
                    all_data[f"{scan_name}_Potential_V"] = pd.Series(scan['potential'])
                    all_data[f"{scan_name}_Current_µA"] = pd.Series(scan['current'])
                
                df = pd.DataFrame(all_data)
                df.to_csv(output_file, index=False)
                return True
                
        except Exception as e:
            print(f"Error exporting to CSV: {str(e)}")
            return False
    
    def export_to_excel(self, output_file):
        """
        Exports data to Excel format with separate sheets for each scan
        
        Args:
            output_file (str): Path to the Excel output file
            
        Returns:
            bool: True if export succeeded, False otherwise
        """
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Sheet with general metadata
                metadata_df = pd.DataFrame([
                    ["Original file", os.path.basename(self.input_file)],
                    ["Import date", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                    ["Number of scans", len(self.scans)]
                ], columns=["Property", "Value"])
                
                metadata_df.to_excel(writer, sheet_name='Metadata', index=False)
                
                # For each scan, create a separate worksheet
                for i, scan in enumerate(self.scans):
                    # Create a dataframe for this scan
                    df = pd.DataFrame({
                        'Potential (V)': scan['potential'],
                        'Current (µA)': scan['current']
                    })
                    
                    # Write the dataframe to a worksheet
                    sheet_name = f"Scan_{i+1}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=3)
                    
                    # Get the worksheet to add metadata
                    worksheet = writer.sheets[sheet_name]
                    worksheet.cell(row=1, column=1, value=f"Name: {scan['name']}")
                    
                    if scan['date']:
                        worksheet.cell(row=2, column=1, value=f"Measurement date: {scan['date'].strftime('%Y-%m-%d %H:%M:%S')}")
                
                # Also create a summary sheet with all data
                all_data = {}
                for i, scan in enumerate(self.scans):
                    all_data[f"Scan_{i+1}_Potential_V"] = pd.Series(scan['potential'])
                    all_data[f"Scan_{i+1}_Current_µA"] = pd.Series(scan['current'])
                
                all_df = pd.DataFrame(all_data)
                all_df.to_excel(writer, sheet_name='All_Scans', index=False)
                
            return True
                
        except Exception as e:
            print(f"Error exporting to Excel: {str(e)}")
            return False
    
    def export_to_txt(self, output_file, scan_index=0, delimiter='\t'):
        """
        Exports a scan to plain text format (column format)
        Useful for importing into other analysis software
        
        Args:
            output_file (str): Path to the output file
            scan_index (int): Index of the scan to export (default: 0)
            delimiter (str): Delimiter to use (default: tab)
            
        Returns:
            bool: True if export succeeded, False otherwise
        """
        try:
            if 0 <= scan_index < len(self.scans):
                scan = self.scans[scan_index]
                with open(output_file, 'w', encoding='utf-8') as f:
                    # Write header
                    f.write(f"Potential (V){delimiter}Current (µA)\n")
                    
                    # Write data
                    for i in range(len(scan['potential'])):
                        f.write(f"{scan['potential'][i]}{delimiter}{scan['current'][i]}\n")
                        
                return True
            else:
                print(f"Error: Invalid scan index {scan_index}")
                return False
                
        except Exception as e:
            print(f"Error exporting to TXT: {str(e)}")
            return False
            
    def export_to_chi(self, output_file, scan_index=0):
        """
        Exports a scan to CHI format (for CH Instruments software)
        
        Args:
            output_file (str): Path to the output file
            scan_index (int): Index of the scan to export (default: 0)
            
        Returns:
            bool: True if export succeeded, False otherwise
        """
        try:
            if 0 <= scan_index < len(self.scans):
                scan = self.scans[scan_index]
                
                # The CHI format is a proprietary format, but we can create a
                # similar text file that might be compatible with some software
                with open(output_file, 'w', encoding='utf-8') as f:
                    # Write header
                    f.write("CH Instruments Data Format\n")
                    f.write(f"Technique: Cyclic Voltammetry\n")
                    f.write(f"File: {os.path.basename(self.input_file)}\n")
                    if scan['date']:
                        f.write(f"Date: {scan['date'].strftime('%m/%d/%Y')}\n")
                        f.write(f"Time: {scan['date'].strftime('%H:%M:%S')}\n")
                    f.write(f"Scan: {scan['name']}\n")
                    f.write(f"Points: {len(scan['potential'])}\n")
                    f.write("Header end\n")
                    
                    # Write data
                    for i in range(len(scan['potential'])):
                        # Convert current from µA to A (format often used by CHI)
                        current_A = scan['current'][i] * 1e-6
                        f.write(f"{scan['potential'][i]}\t{current_A:.12e}\n")
                        
                return True
            else:
                print(f"Error: Invalid scan index {scan_index}")
                return False
                
        except Exception as e:
            print(f"Error exporting to CHI format: {str(e)}")
            return False

def main():
    """Main function for command-line usage"""
    parser = argparse.ArgumentParser(description='Import and convert cyclic voltammetry data')
    parser.add_argument('input_file', help='Input CSV file')
    parser.add_argument('-o', '--output', help='Output file (default: output.xlsx)', default='output.xlsx')
    parser.add_argument('-f', '--format', choices=['csv', 'excel', 'txt', 'chi'], 
                        help='Output format (default: excel)', default='excel')
    parser.add_argument('-s', '--scan', type=int, help='Index of the scan to export (0-based)', default=0)
    parser.add_argument('-e', '--encoding', help='Input file encoding (default: utf-16)', default='utf-16')
    
    args = parser.parse_args()
    
    importer = VoltammetryImporter(args.input_file, encoding=args.encoding)
    if not importer.load_file():
        sys.exit(1)
    
    print(f"Loaded {importer.get_scan_count()} scans")
    for i, name in enumerate(importer.get_scan_names()):
        print(f"  {i}: {name}")
    
    success = False
    if args.format == 'csv':
        success = importer.export_to_csv(args.output, args.scan)
    elif args.format == 'excel':
        success = importer.export_to_excel(args.output)
    elif args.format == 'txt':
        success = importer.export_to_txt(args.output, args.scan)
    elif args.format == 'chi':
        success = importer.export_to_chi(args.output, args.scan)
    
    if success:
        print(f"File exported successfully: {args.output}")
        return 0
    else:
        print("Export failed")
        return 1

if __name__ == "__main__":
    sys.exit(main())