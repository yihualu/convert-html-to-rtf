"""Convert HTML to RTF using Microsoft Word COM interface.

This script uses Word's built-in HTML to RTF conversion which preserves
formatting including colors, bold, italic, etc.
"""

import csv
import os
import tempfile
from pathlib import Path
import argparse
import time
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 package is required for Word automation")
    print("Install it with: pip install pywin32")
    exit(1)


def extract_rtf_content(rtf_full: str) -> str:
    """
    Extract just the content portion from full RTF, keeping color table but removing metadata.
    
    Args:
        rtf_full: Full RTF document with headers, font tables, etc.
        
    Returns:
        Minimal RTF with color table and content
    """
    if not rtf_full or not rtf_full.strip():
        return ""
    
    # Extract color table (needed for \cf codes to work)
    colortbl_match = re.search(r'\{\\colortbl[^}]*\}', rtf_full)
    colortbl = colortbl_match.group(0) if colortbl_match else ''
    
    # Find where themedata starts (this marks end of actual content)
    themedata_match = re.search(r'\{\\?\*\\themedata', rtf_full)
    if themedata_match:
        content_section = rtf_full[:themedata_match.start()]
    else:
        content_section = rtf_full
    
    # Find the actual text content within this section
    # Look for the paragraph with actual text (after all the \pard formatting)
    # Pattern: \pard ... {text content}\par
    
    # Find the last \pard before content
    pard_matches = list(re.finditer(r'\\pard[^{]*', content_section))
    if not pard_matches:
        return rtf_full
    
    last_pard = pard_matches[-1]
    content_start = last_pard.end()
    
    # Get everything from after \pard to end of content section
    content = content_section[content_start:].strip()
    
    # Remove only the outermost trailing } (the closing of the main document)
    if content.endswith('}'):
        content = content[:-1].strip()
    
    # Fix Word's bug: insert \cf0 to reset color after colored text groups end
    # Pattern: }{\rtlch... without \cf means color should reset
    content = re.sub(r'(\})\{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\dbch', r'\1{\\rtlch\\fcs1 \\af0 \\ltrch\\fcs0 \\cf0\\dbch', content)
    
    # Wrap in minimal RTF structure with color table
    if colortbl:
        return '{\\rtf1\\ansi\\deff0 ' + colortbl + '{\\pard ' + content + '}}'
    else:
        return '{\\rtf1\\ansi\\deff0 {\\pard ' + content + '}}'


def html_to_rtf_word(html_content: str) -> str:
    """
    Convert HTML to RTF using Microsoft Word.
    
    Args:
        html_content: HTML string to convert
        
    Returns:
        RTF string with full formatting preserved
    """
    if not html_content or not html_content.strip():
        return ""
    
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    
    try:
        # Create temporary files
        with tempfile.TemporaryDirectory() as temp_dir:
            html_file = Path(temp_dir) / "temp.html"
            rtf_file = Path(temp_dir) / "temp.rtf"
            
            # Write HTML with basic structure so Word recognizes it as HTML
            html_doc = f'''<html>
<head><meta charset="utf-8"></head>
<body>
{html_content}
</body>
</html>'''
            html_file.write_text(html_doc, encoding='utf-8')
            
            # Initialize Word application
            word = None
            doc = None
            try:
                # Use DispatchEx for a new independent instance
                word = win32com.client.DispatchEx("Word.Application")
                # Try to set properties but don't fail if they can't be set
                try:
                    word.Visible = False
                except:
                    pass
                try:
                    word.DisplayAlerts = 0  # wdAlertsNone
                except:
                    pass
                
                # Open HTML file with format specified
                # wdOpenFormatAuto = 0, ConfirmConversions = False to auto-detect HTML
                doc = word.Documents.Open(str(html_file.absolute()), 
                                         ConfirmConversions=False,
                                         Format=0)  # wdOpenFormatAuto
                
                # Save as RTF (wdFormatRTF = 6)
                doc.SaveAs2(str(rtf_file.absolute()), FileFormat=6)
                
                # Read RTF content - return full RTF file
                rtf_full = rtf_file.read_text(encoding='utf-8', errors='ignore')
                
                # Extract just the content portion
                rtf_content = extract_rtf_content(rtf_full)
                
                return rtf_content
                
            except Exception as e:
                print(f"Error during Word conversion: {e}")
                return ""
            finally:
                # Clean up
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                    except:
                        pass
                if word:
                    try:
                        word.Quit()
                    except:
                        pass
    finally:
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()


def process_row(row_data, html_col_index, rtf_col_index):
    """Process a single row - designed to be called by thread pool."""
    row_num, row = row_data
    # Get HTML from specified column
    html = row[html_col_index] if len(row) > html_col_index else ""
    
    # Convert to RTF
    rtf = html_to_rtf_word(html)
    
    # Ensure row has enough columns for RTF output
    while len(row) <= rtf_col_index:
        row.append("")
    
    # Set RTF in specified column
    row[rtf_col_index] = rtf
    
    return row_num, row


def process_csv(input_csv: str, output_csv: str = None, num_rows: int = None, progress_every: int = 100, max_workers: int = 8, html_col: int = 3, rtf_col: int = 4):
    """
    Process CSV file, converting HTML from specified column to RTF in output column.
    
    Args:
        input_csv: Path to input CSV file
        output_csv: Path to output CSV file (default: test_output.csv)
        num_rows: Number of rows to process (default: all rows)
        progress_every: Print progress every N rows
        max_workers: Number of parallel Word instances (default: 8)
        html_col: Column number for HTML input (1-based, default: 3)
        rtf_col: Column number for RTF output (1-based, default: 4)
    """
    if not output_csv:
        # Always use the same output file - overwrite previous test
        output_csv = str(Path(__file__).parent / 'test_output.csv')
    
    # Convert 1-based column numbers to 0-based indices
    html_col_index = html_col - 1
    rtf_col_index = rtf_col - 1
    
    print(f"Converting HTML to RTF using Microsoft Word...")
    print(f"Input: {input_csv}")
    print(f"Output: {output_csv}")
    print(f"HTML column: {html_col}, RTF column: {rtf_col}")
    if num_rows:
        print(f"Rows to process: {num_rows}")
    print(f"Threads: {max_workers}")
    print()
    
    start_time = time.time()
    
    # Read all rows first
    with open(input_csv, 'r', encoding='utf-8', newline='') as infile:
        reader = csv.reader(infile)
        header = next(reader)
        
        # Collect rows to process
        rows_to_process = []
        for i, row in enumerate(reader):
            if num_rows and i >= num_rows:
                break
            rows_to_process.append((i, row))
    
    total_rows = len(rows_to_process)
    print(f"Processing {total_rows} rows with {max_workers} threads...")
    
    # Process rows in parallel
    results = {}
    completed = 0
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Submit all tasks
        future_to_row = {executor.submit(process_row, row_data, html_col_index, rtf_col_index): row_data[0] 
                        for row_data in rows_to_process}
        
        # Process completed tasks
        for future in as_completed(future_to_row):
            row_num, processed_row = future.result()
            results[row_num] = processed_row
            completed += 1
            
            if progress_every and completed % progress_every == 0:
                elapsed = time.time() - start_time
                rate = completed / elapsed
                remaining = (total_rows - completed) / rate if rate > 0 else 0
                print(f"Processed {completed}/{total_rows} rows ({completed*100//total_rows}%) - "
                      f"{rate:.1f} rows/sec - ETA: {remaining:.0f}s")
    
    # Write results in original order
    with open(output_csv, 'w', encoding='utf-8', newline='') as outfile:
        writer = csv.writer(outfile)
        writer.writerow(header)
        
        for i in range(total_rows):
            writer.writerow(results[i])
    
    elapsed = time.time() - start_time
    print(f"\nDone. Processed {total_rows} rows in {elapsed:.1f}s ({total_rows/elapsed:.1f} rows/sec)")
    print(f"Output: {output_csv}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert HTML to RTF using Microsoft Word (multi-threaded)"
    )
    parser.add_argument(
        "input_csv",
        help="Input CSV file path"
    )
    parser.add_argument(
        "--output",
        help="Output CSV file path (default: test_output.csv)"
    )
    parser.add_argument(
        "--num-rows",
        type=int,
        help="Number of rows to process (default: all rows)"
    )
    parser.add_argument(
        "--progress-every",
        type=int,
        default=100,
        help="Print progress every N rows (default: 100)"
    )
    parser.add_argument(
        "--threads",
        type=int,
        default=4,
        help="Number of concurrent Word instances (default: 4)"
    )
    parser.add_argument(
        "--html-col",
        type=int,
        default=3,
        help="Column number for HTML input (1-based, default: 3)"
    )
    parser.add_argument(
        "--rtf-col",
        type=int,
        default=4,
        help="Column number for RTF output (1-based, default: 4)"
    )
    
    args = parser.parse_args()
    
    # Check if Microsoft Word is available
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        print("✓ Microsoft Word is available")
        print()
    except Exception as e:
        print("ERROR: Microsoft Word is not available")
        print(f"Details: {e}")
        print()
        print("This script requires:")
        print("  1. Microsoft Word installed")
        print("  2. pywin32 package: pip install pywin32")
        exit(1)
    
    process_csv(args.input_csv, args.output, args.num_rows, args.progress_every, args.threads, args.html_col, args.rtf_col)
