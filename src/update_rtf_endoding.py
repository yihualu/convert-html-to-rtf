"""Fix RTF strings in CSV column D - clean encoding and formatting."""

import csv
import re

def fix_rtf_string(rtf_content):
    """
    Fix RTF string encoding - apply proper header with deflang1036.
    """
    if not rtf_content or not rtf_content.startswith('{\\rtf'):
        return rtf_content
    
    # Replace header with standard format including deflang1036
    rtf = rtf_content
    rtf = re.sub(r'\\rtf1\\ansi.*?\\deff\d+.*?(?=\\|{)', r'\\rtf1\\ansi\\ansicpg1252\\deff0\\deflang1036\n', rtf)
    
    return rtf

def process_csv_file(input_file, output_file):
    """Process CSV file and fix all RTF strings in column D."""
    print(f"Reading {input_file}...")
    
    rows = []
    with open(input_file, 'r', encoding='utf-8', newline='') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    print(f"Total rows: {len(rows)}")
    
    # Fix RTF strings in column D (index 3)
    fixed_count = 0
    for i, row in enumerate(rows):
        if len(row) > 3 and row[3]:  # Column D exists and has content
            original = row[3]
            fixed = fix_rtf_string(original)
            if original != fixed:
                row[3] = fixed
                fixed_count += 1
                if i % 100 == 0:
                    print(f"  Processed row {i}... ({fixed_count} fixed so far)")
    
    print(f"\nWriting {output_file}...")
    print(f"Total RTF strings fixed: {fixed_count}")
    
    with open(output_file, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(rows)
    
    print("Done!")

if __name__ == '__main__':
    input_file = 'exported_medical_record_note_in_html_rich_text_format_v4_with_rtf_full_output_20251215.csv'
    output_file = 'exported_medical_record_note_in_html_rich_text_format_v4_with_rtf_fixed_deflang1036.csv'
    
    process_csv_file(input_file, output_file)
