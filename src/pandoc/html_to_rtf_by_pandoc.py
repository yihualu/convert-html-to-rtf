"""Convert HTML to RTF using Pandoc.

This script uses Pandoc's HTML to RTF conversion with preprocessing and 
post-processing to preserve formatting including colors, bold, italic, etc.
"""

import argparse
import csv
import re
from pathlib import Path
from typing import Optional, List, Tuple
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pypandoc


def extract_color_spans(html: str) -> List[Tuple[str, str, Optional[str]]]:
    """
    Extract text spans with their colors from HTML.
    Returns list of (text, fg_color_hex, bg_color_hex) tuples.
    Colors are like 'FF2712' or None.
    """
    # Find all spans with color styles (foreground and/or background)
    color_spans = []
    
    # Combine patterns - look for spans with either or both colors
    combined_pattern = r'<span\s+style="([^"]*)"[^>]*>(.*?)</span>'
    for match in re.finditer(combined_pattern, html, re.IGNORECASE | re.DOTALL):
        style = match.group(1)
        text = match.group(2)
        
        # Extract colors from style
        fg_color = None
        bg_color = None
        
        # Look for foreground color (more strict pattern to avoid false matches)
        fg_match = re.search(r'(?:^|;)\s*color\s*:\s*#([0-9A-Fa-f]{6})', style, re.IGNORECASE)
        if fg_match:
            fg_color = fg_match.group(1)
        
        bg_match = re.search(r'background-color\s*:\s*#([0-9A-Fa-f]{6})', style, re.IGNORECASE)
        if bg_match:
            bg_color = bg_match.group(1)
        
        # Only add if we found at least one color
        if fg_color or bg_color:
            color_spans.append((text, fg_color, bg_color))
    
    return color_spans


def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """Convert hex color like 'FF2712' to RGB tuple (255, 39, 18)."""
    hex_color = hex_color.strip('#')
    return (
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16)
    )


def preprocess_html_for_pandoc(html: str) -> Tuple[str, List[Tuple[str, str, Optional[str]]]]:
    """
    Convert inline CSS styles to semantic HTML tags that Pandoc recognizes.
    Also extracts color information for post-processing.
    
    Returns: (processed_html, color_spans)
    """
    if not html:
        return html, []
    
    # First, extract color spans before we process them
    color_spans = extract_color_spans(html)
    
    # Convert strikethrough to <del> or <s> tags
    html = re.sub(
        r'<span\s+style="[^"]*text-decoration:\s*line-through[^"]*"[^>]*>(.*?)</span>',
        r'<del>\1</del>',
        html,
        flags=re.IGNORECASE | re.DOTALL
    )
    
    # Pattern to match span tags with style attributes
    def replace_span_with_styles(match):
        full_match = match.group(0)
        style_attr = match.group(1) if match.group(1) else ""
        content = match.group(2)
        
        # Check for various style properties
        is_bold = bool(re.search(r'font-weight\s*:\s*(bold|700|800|900)', style_attr, re.IGNORECASE))
        is_italic = bool(re.search(r'font-style\s*:\s*italic', style_attr, re.IGNORECASE))
        is_underline = bool(re.search(r'text-decoration\s*:\s*underline', style_attr, re.IGNORECASE))
        is_strikethrough = bool(re.search(r'text-decoration\s*:\s*line-through', style_attr, re.IGNORECASE))
        has_color = bool(re.search(r'(?:^|;)\s*color\s*:\s*#[0-9A-Fa-f]{6}', style_attr, re.IGNORECASE))
        has_bgcolor = bool(re.search(r'background-color\s*:\s*#[0-9A-Fa-f]{6}', style_attr, re.IGNORECASE))
        
        # Wrap content with appropriate tags
        if is_bold:
            content = f"<strong>{content}</strong>"
        if is_italic:
            content = f"<em>{content}</em>"
        if is_underline:
            content = f"<u>{content}</u>"
        if is_strikethrough:
            content = f"<del>{content}</del>"
        
        # If we added formatting OR has color, return just the content
        # (color will be handled in post-processing)
        if is_bold or is_italic or is_underline or is_strikethrough or has_color or has_bgcolor:
            return content
        else:
            # Keep the original span if it has no recognized styles
            return full_match
    
    # Match <span style="...">content</span>
    html = re.sub(
        r'<span\s+style="([^"]*)"[^>]*>(.*?)</span>',
        replace_span_with_styles,
        html,
        flags=re.IGNORECASE | re.DOTALL
    )
    
    return html, color_spans


def apply_colors_to_rtf(rtf: str, color_spans: List[Tuple[str, str, Optional[str]]]) -> str:
    """
    Post-process RTF to add color formatting.
    This adds color table entries and wraps matching text with color codes.
    Supports both foreground (cf) and background (highlight) colors.
    """
    if not color_spans or not rtf:
        return rtf
    
    # Check if RTF already has a color table
    has_colortbl = '\\colortbl' in rtf
    
    # Build color table entries for both foreground and background colors
    unique_colors = {}
    color_index = 1  # RTF color indices start at 1
    
    for text, fg_color, bg_color in color_spans:
        if fg_color and fg_color not in unique_colors:
            unique_colors[fg_color] = color_index
            color_index += 1
        if bg_color and bg_color not in unique_colors:
            unique_colors[bg_color] = color_index
            color_index += 1
    
    # If no color table exists, we can't easily inject one without full RTF parsing
    # For now, we'll work with the existing structure
    if not unique_colors:
        return rtf
    
    # Insert colors into existing color table or create one
    color_table_entries = []
    for hex_color, idx in unique_colors.items():
        r, g, b = hex_to_rgb(hex_color)
        color_table_entries.append(f"\\red{r}\\green{g}\\blue{b};")
    
    # Replace or enhance color table
    if has_colortbl:
        # Find and replace the color table
        def replace_colortbl(match):
            existing = match.group(1)
            new_colors = "".join(color_table_entries)
            # Append new colors to existing table
            return f"{{\\colortbl{existing}{new_colors}}}"
        
        rtf = re.sub(r'\{\\colortbl([^}]*)\}', replace_colortbl, rtf, count=1)
    else:
        # Need to insert color table
        color_table = "{\\colortbl;" + "".join(color_table_entries) + "}"
        
        # Check if this is a fragment or full RTF document
        is_full_doc = rtf.startswith('{\\rtf1')
        
        if is_full_doc:
            # Try to insert after font table if it exists
            if '\\fonttbl' in rtf:
                # Find the end of the font table (handle nested braces)
                match = re.search(r'\{\\fonttbl', rtf)
                if match:
                    pos = match.start()
                    brace_count = 0
                    i = pos
                    while i < len(rtf):
                        if rtf[i] == '{':
                            brace_count += 1
                        elif rtf[i] == '}':
                            brace_count -= 1
                            if brace_count == 0:
                                # Insert color table right after this position
                                rtf = rtf[:i+1] + '\n' + color_table + rtf[i+1:]
                                break
                        i += 1
            else:
                # No font table, insert after RTF header
                rtf = re.sub(
                    r'(\{\\rtf1[^\{]*)',
                    lambda m: m.group(1) + '\n' + color_table,
                    rtf,
                    count=1
                )
        else:
            # Fragment - wrap with minimal RTF structure including color table
            rtf_header = "{\\rtf1\\ansi\\deff0\n" + color_table + "\n"
            rtf = rtf_header + rtf + "}"
    
    # Now wrap text spans with color codes
    # Sort by length (longest first) to handle nested/overlapping spans better
    sorted_spans = sorted(color_spans, key=lambda x: len(x[0]), reverse=True)
    
    for text, fg_color, bg_color in sorted_spans:
        if not text:
            continue
            
        # Build color formatting codes
        color_codes = []
        if fg_color and fg_color in unique_colors:
            fg_idx = unique_colors[fg_color]
            color_codes.append(f'\\cf{fg_idx}')
        if bg_color and bg_color in unique_colors:
            bg_idx = unique_colors[bg_color]
            color_codes.append(f'\\highlight{bg_idx}')
        
        if not color_codes:
            continue
        
        # Create the RTF color wrapper
        color_code_start = '{{' + ' '.join(color_codes) + ' '
        color_code_end = '}}'
        
        # Try to find and replace the text
        # Need to handle: HTML entities, <br/> tags, and RTF unicode escapes
        
        # Common HTML entity conversions (including numeric entities)
        html_entities = {
            '&gt;': '>', '&lt;': '<', '&amp;': '&', '&quot;': '"',
            '&agrave;': 'à', '&aacute;': 'á', '&acirc;': 'â', '&atilde;': 'ã', '&auml;': 'ä',
            '&egrave;': 'è', '&eacute;': 'é', '&ecirc;': 'ê', '&euml;': 'ë',
            '&igrave;': 'ì', '&iacute;': 'í', '&icirc;': 'î', '&iuml;': 'ï',
            '&ograve;': 'ò', '&oacute;': 'ó', '&ocirc;': 'ô', '&otilde;': 'õ', '&ouml;': 'ö',
            '&ugrave;': 'ù', '&uacute;': 'ú', '&ucirc;': 'û', '&uuml;': 'ü',
            '&ccedil;': 'ç', '&ntilde;': 'ñ',
            '&reg;': '®', '&copy;': '©', '&trade;': '™',
            '&euro;': '€', '&pound;': '£', '&yen;': '¥',
            '&rsquo;': ''', '&lsquo;': ''', '&rdquo;': '"', '&ldquo;': '"',
            '&mdash;': '—', '&ndash;': '–',
        }
        
        # Convert numeric HTML entities (&#9; -> tab, etc.)
        import re as regex_module
        converted = text
        # Replace &#NNN; with corresponding character
        converted = regex_module.sub(r'&#(\d+);', lambda m: chr(int(m.group(1))), converted)
        
        # Convert named entities
        for entity, char in html_entities.items():
            converted = converted.replace(entity, char)
        
        # Create text variants to try
        text_variants = []
        
        # Variant 1: Original text with HTML entities
        text_variants.append(text)
        
        # Variant 2: Converted text with entities replaced
        text_variants.append(converted)
        
        # Variant 3: Replace <br/> tags with RTF line breaks
        # In RTF, <br/> becomes \\line (literal backslash + "line")
        rtf_with_line = converted.replace('<br/>', '\\line ').replace('<br>', '\\line ')
        text_variants.append(rtf_with_line)
        
        # Variant 4: Also try just removing <br/> (in case it becomes a space or nothing)
        no_br = converted.replace('<br/>', ' ').replace('<br>', ' ')
        text_variants.append(no_br)
        
        # Variant 5: Convert Unicode characters to RTF escapes (with space+?)
        rtf_escaped = ''
        for char in converted:
            code = ord(char)
            if code > 127:  # Non-ASCII
                rtf_escaped += f'\\u{code} ?'
            else:
                rtf_escaped += char
        text_variants.append(rtf_escaped)
        
        # Variant 5b: RTF escapes without the space (e.g., \u8217' instead of \u8217 ?)
        rtf_escaped_nospace = ''
        for char in converted:
            code = ord(char)
            if code > 127:
                rtf_escaped_nospace += f'\\u{code}'
            else:
                rtf_escaped_nospace += char
        text_variants.append(rtf_escaped_nospace)
        
        # Variant 6: RTF escaped with \\line
        rtf_escaped_with_line = rtf_escaped.replace('<br/>', '\\line ').replace('<br>', '\\line ')
        text_variants.append(rtf_escaped_with_line)
        
        # Variant 6b: RTF escaped no-space with \\line
        rtf_escaped_nospace_with_line = rtf_escaped_nospace.replace('<br/>', '\\line ').replace('<br>', '\\line ')
        text_variants.append(rtf_escaped_nospace_with_line)
        
        # Variant 7: RTF escaped without <br/>
        rtf_escaped_no_br = rtf_escaped.replace('<br/>', ' ').replace('<br>', ' ')
        text_variants.append(rtf_escaped_no_br)
        
        # Variant 7b: RTF escaped no-space without <br/>
        rtf_escaped_nospace_no_br = rtf_escaped_nospace.replace('<br/>', ' ').replace('<br>', ' ')
        text_variants.append(rtf_escaped_nospace_no_br)
        
        # Variant 8: Trimmed versions of all above
        all_variants = [text, converted, rtf_with_line, no_br, rtf_escaped, rtf_escaped_nospace,
                       rtf_escaped_with_line, rtf_escaped_nospace_with_line, 
                       rtf_escaped_no_br, rtf_escaped_nospace_no_br]
        for v in all_variants:
            text_variants.append(v.strip())
        
        # Variant 9: Normalize whitespace (multiple spaces -> single space)
        # This helps with HTML that has double spaces but RTF normalizes to single
        import re as regex
        for v in all_variants:
            normalized = regex.sub(r'\s+', ' ', v)
            text_variants.append(normalized)
            text_variants.append(normalized.strip())
        
        # Try each variant
        replaced = False
        for variant in text_variants:
            if not variant:
                continue
            if variant in rtf and not replaced:
                rtf = rtf.replace(variant, color_code_start + variant + color_code_end, 1)
                replaced = True
                break
    
    return rtf


def html_to_rtf(html_fragment: str) -> str:
    """
    Convert an HTML fragment to RTF using pypandoc (pandoc backend).
    Preprocesses HTML to convert CSS styles to semantic tags.
    Post-processes RTF to add color formatting.
    """
    if not html_fragment or not html_fragment.strip():
        return ""

    try:
        # Preprocess HTML to convert inline styles to tags and extract colors
        preprocessed_html, color_spans = preprocess_html_for_pandoc(html_fragment)
        
        rtf_output = pypandoc.convert_text(
            preprocessed_html,
            to="rtf",
            format="html",
        )
        
        # Post-process to apply colors
        if color_spans:
            rtf_output = apply_colors_to_rtf(rtf_output, color_spans)
        
        return rtf_output
    except Exception as e:
        print(f"Warning: Conversion failed: {e}")
        return ""


def process_csv(
    input_csv: str,
    output_csv: str = None,
    num_rows: int = None,
    progress_every: int = 100,
    html_col: int = 3,
    rtf_col: int = 4,
    encoding: str = "utf-8",
) -> None:
    """
    Convert the HTML in specified column to RTF using pandoc.

    Args:
        input_csv: Path to input CSV file
        output_csv: Path to output CSV file (default: test_output.csv)
        num_rows: Number of rows to process (default: all rows)
        progress_every: Print progress every N rows (default: 100)
        html_col: Column number for HTML input (1-based, default: 3)
        rtf_col: Column number for RTF output (1-based, default: 4)
        encoding: CSV encoding (default: utf-8)
    """
    if not output_csv:
        output_csv = str(Path(__file__).parent / 'test_output.csv')
    
    # Convert 1-based column numbers to 0-based indices
    html_col_index = html_col - 1
    rtf_col_index = rtf_col - 1
    
    input_path = Path(input_csv)
    output_path = Path(output_csv)

    # Ensure pandoc is available (fail fast)
    try:
        ver = pypandoc.get_pandoc_version()
        print(f"✓ Pandoc {ver} is available")
    except Exception as e:
        print(f"ERROR: Pandoc is not available or not found")
        print(f"Details: {e}")
        print()
        print("Install pandoc from: https://pandoc.org")
        raise SystemExit(1)
    
    print(f"Converting HTML to RTF using pandoc...")
    print(f"Input: {input_csv}")
    print(f"Output: {output_csv}")
    print(f"HTML column: {html_col}, RTF column: {rtf_col}")
    if num_rows:
        print(f"Rows to process: {num_rows}")
    print(f"Encoding: {encoding}")
    print()
    
    start_time = time.time()

    processed = 0
    total_rows = 0
    
    # First pass: count rows (excluding header)
    with input_path.open("r", newline="", encoding=encoding, errors="replace") as f:
        reader = csv.reader(f)
        next(reader)  # Skip header
        for _ in reader:
            total_rows += 1
            if num_rows and total_rows >= num_rows:
                total_rows = num_rows
                break
    
    print(f"Processing {total_rows} rows...")
    
    # Second pass: process rows
    with input_path.open("r", newline="", encoding=encoding, errors="replace") as inf, \
         output_path.open("w", newline="", encoding=encoding) as outf:
        reader = csv.reader(inf)
        writer = csv.writer(outf)
        
        # Write header
        header = next(reader)
        writer.writerow(header)
        
        for idx, row in enumerate(reader):
            # Make sure the row has the target HTML column
            while len(row) <= html_col_index:
                row.append("")

            html_content = row[html_col_index] or ""
            rtf_content = html_to_rtf(html_content)

            # Ensure row has enough columns for RTF output
            while len(row) <= rtf_col_index:
                row.append("")
            row[rtf_col_index] = rtf_content

            writer.writerow(row)
            processed += 1

            if progress_every and processed % progress_every == 0:
                elapsed = time.time() - start_time
                rate = processed / elapsed if elapsed > 0 else 0
                remaining = (total_rows - processed) / rate if rate > 0 else 0
                print(f"Processed {processed}/{total_rows} rows ({processed*100//total_rows}%) - "
                      f"{rate:.1f} rows/sec - ETA: {remaining:.0f}s")

            if num_rows and processed >= num_rows:
                break
    
    elapsed = time.time() - start_time
    print(f"\nDone. Processed {processed} rows in {elapsed:.1f}s ({processed/elapsed:.1f} rows/sec)")
    print(f"Output: {output_csv}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert HTML to RTF using Pandoc (single-threaded)"
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
    parser.add_argument(
        "--encoding",
        type=str,
        default="utf-8",
        help="CSV encoding (default: utf-8)"
    )
    
    args = parser.parse_args()
    process_csv(args.input_csv, args.output, args.num_rows, args.progress_every, 
                args.html_col, args.rtf_col, args.encoding)
