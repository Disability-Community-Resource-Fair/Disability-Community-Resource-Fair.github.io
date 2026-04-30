#!/usr/bin/env python3
"""
Generate 2026 Disability Community Resource Fair Vendor Directory

This script:
1. Reads 2026 Vendor Requests from fair_logistics.xlsx
2. Extracts vendor requests CSV
3. Reads vendor details from _posts markdown files
4. Combines vendor request numbers with detailed vendor information
5. Generates a comprehensive vendor directory markdown with:
   - Vendor listings organized by table number
   - Full service descriptions (no truncation)
   - Service category index (10 main categories)
   - Age group index (5 age ranges)
"""

import os
import re
import frontmatter
import csv
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

def extract_vendors_from_excel(excel_path, output_csv_path):
    """Extract vendors from fair_logistics.xlsx 2026 Vendor Requests sheet."""
    
    def extract_cell_value(cell, shared_strings):
        """Extract value from a cell, handling shared strings."""
        value_elem = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
        
        if value_elem is not None:
            cell_type = cell.get('t')
            text_value = value_elem.text or ''
            
            if cell_type == 's':
                try:
                    idx = int(text_value)
                    return shared_strings[idx] if idx < len(shared_strings) else ''
                except:
                    return text_value
            else:
                return text_value
        
        rich_text = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is')
        if rich_text is not None:
            text_parts = rich_text.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
            return ''.join([t.text or '' for t in text_parts])
        
        return ''
    
    with zipfile.ZipFile(excel_path, 'r') as zip_ref:
        # Load shared strings
        shared_strings = []
        try:
            shared_strings_xml = zip_ref.read('xl/sharedStrings.xml')
            ss_root = ET.fromstring(shared_strings_xml)
            string_items = ss_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')
            
            for si in string_items:
                t_elem = si.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                if t_elem is not None:
                    shared_strings.append(t_elem.text or '')
                else:
                    text_parts = si.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}r/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                    shared_strings.append(''.join([t.text or '' for t in text_parts]))
        except Exception as e:
            print(f"Warning: Error loading shared strings: {e}")
        
        # Read the 2026 Vendor Requests sheet
        worksheet_xml = zip_ref.read('xl/worksheets/sheet10.xml')
        ws_root = ET.fromstring(worksheet_xml)
        
        rows = ws_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
        
        all_data = []
        for row in rows:
            cells = row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
            row_data = []
            for cell in cells:
                value = extract_cell_value(cell, shared_strings)
                row_data.append(value)
            if any(row_data):
                all_data.append(row_data)
        
        # Save to CSV
        with open(output_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerows(all_data)
        
        return len(all_data)

def read_vendor_requests(csv_path):
    """Read vendor requests from CSV."""
    vendor_requests = {}
    
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) >= 2 and row[0].strip() and row[1].strip():
                vendor_num = row[0].strip()
                vendor_name = row[1].strip()
                
                # Skip header rows that may contain 'VENDOR REQUESTS' or similar
                if 'VENDOR REQUESTS' in vendor_num.upper() or vendor_name.strip().upper() == 'ARRIVED':
                    continue
                # Preserve the original table identifier (could be a range like '19-21', 'Lobby', or '3 & 4')
                # Store as string so we don't drop valid entries that are non-integer.
                vendor_requests[vendor_name] = vendor_num
    
    return vendor_requests

def read_vendor_details(posts_dir):
    """Read vendor details from markdown files in _posts."""
    vendor_details = {}
    
    md_files = sorted([f for f in os.listdir(posts_dir) if f.endswith('.md')])
    
    for md_file in md_files:
        filepath = os.path.join(posts_dir, md_file)
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                post = frontmatter.load(f)
            
            vendor_name = md_file.split('-', 3)[-1].replace('.md', '')
            
            metadata = {
                'title': post.metadata.get('title', vendor_name),
                'categories': post.metadata.get('categories', []),
                'tags': post.metadata.get('tags', []),
                'content': post.content,
                'filename': md_file
            }
            
            vendor_details[vendor_name] = metadata
        except Exception as e:
            pass
    
    return vendor_details

def read_vendor_map_entries(vendor_map_path):
    """Read vendor labels from the vendor map workbook drawing layer."""
    vendor_map_entries = {}

    if not os.path.exists(vendor_map_path):
        return []

    drawing_path = 'xl/drawings/drawing1.xml'
    ns = {
        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    }

    with zipfile.ZipFile(vendor_map_path, 'r') as zip_ref:
        if drawing_path not in zip_ref.namelist():
            return []

        root = ET.fromstring(zip_ref.read(drawing_path))
        for shape in root.findall('.//xdr:sp', ns):
            text_parts = [t.text or '' for t in shape.findall('.//a:t', ns)]
            raw_text = ''.join(text_parts).strip()
            if not raw_text:
                continue

            match = re.match(r'^(Lobby|\d+\s*&\s*\d+|\d+(?:-\d+)?)\s+(.*\S)\s*$', raw_text)
            if not match:
                continue

            number_text = match.group(1).replace(' ', '')
            display_name = match.group(2).strip()
            normalized_name = normalize_name(display_name)

            if not display_name or not normalized_name:
                continue

            entry = vendor_map_entries.get(normalized_name)
            if entry is None:
                entry = {
                    'display_name': display_name,
                    'normalized_name': normalized_name,
                    'number_values': set(),
                }
                vendor_map_entries[normalized_name] = entry

            entry['number_values'].update(parse_table_number_set(number_text))

    return list(vendor_map_entries.values())

def clean_category_name(category):
    """Clean category name by replacing underscores with spaces."""
    return category.replace('_', ' ').strip()

def parse_categories(categories_str):
    """Parse categories string into individual category names.
    
    Categories are SPACE-delimited (not comma-delimited).
    Within each category, replace underscores with spaces.
    Commas within categories are kept as-is.
    """
    if isinstance(categories_str, str):
        # Split by spaces to get individual categories
        cats = [c.strip() for c in categories_str.split()]
    else:
        cats = categories_str if isinstance(categories_str, list) else []
    
    # Clean up each category
    cleaned = []
    for cat in cats:
        if cat:
            # Replace underscores with spaces within the category
            cat = cat.replace('_', ' ')
            if cat:
                cleaned.append(cat)
    
    return cleaned

def parse_tags(tags_str):
    """Parse tags string into individual tag names."""
    if isinstance(tags_str, str):
        tags = [t.strip() for t in tags_str.split()]
    else:
        tags = tags_str if isinstance(tags_str, list) else []
    
    cleaned = []
    for tag in tags:
        if tag:
            tag = tag.replace('_', ' ')
            if tag:
                cleaned.append(tag)
    
    return cleaned

def normalize_name(s):
    if not s:
        return ''
    ns = s.lower().replace('&', 'and')
    ns = re.sub(r'[^a-z0-9\s]', '', ns)
    return ' '.join(ns.split())

def names_match(left, right, threshold=0.85):
    left_norm = normalize_name(left)
    right_norm = normalize_name(right)
    if not left_norm or not right_norm:
        return False
    if left_norm == right_norm:
        return True
    if left_norm in right_norm or right_norm in left_norm:
        return True

    import difflib

    return difflib.SequenceMatcher(a=left_norm, b=right_norm).ratio() >= threshold

def parse_table_number_set(number_text):
    if number_text is None:
        return set()

    text = str(number_text).strip()
    if not text:
        return set()

    if text.lower() == 'lobby':
        return {'Lobby'}

    normalized = text.replace(' ', '')
    parts = re.split(r'(?:&|,|/|;|\band\b)', normalized, flags=re.IGNORECASE)
    numbers = set()

    for part in parts:
        token = part.strip()
        if not token:
            continue

        range_match = re.fullmatch(r'(\d+)-(\d+)', token)
        if range_match:
            start = int(range_match.group(1))
            end = int(range_match.group(2))
            if start <= end:
                numbers.update(range(start, end + 1))
            else:
                numbers.update(range(end, start + 1))
            continue

        if token.isdigit():
            numbers.add(int(token))
            continue

        numbers.add(token)

    return numbers

def format_table_number_set(number_values):
    if not number_values:
        return ''

    def sort_key(value):
        if isinstance(value, int):
            return (0, value)
        return (1, str(value))

    return ', '.join(str(value) for value in sorted(number_values, key=sort_key))

def generate_vendor_directory(vendor_requests, vendor_details, output_path, output_missing_path=None, vendor_map_path=None):
    """Generate the vendor directory markdown file."""
    
    # Match vendor requests with details
    vendors_combined = []
    
    for req_name, req_num in vendor_requests.items():
        detail = None
        for detail_name, detail_info in vendor_details.items():
            req_lower = req_name.lower().replace('&', 'and').replace(' ', '')
            detail_lower = detail_name.lower().replace('&', 'and').replace(' ', '')
            detail_title_lower = detail_info['title'].lower().replace('&', 'and').replace(' ', '')
            
            if req_lower == detail_lower or req_lower == detail_title_lower:
                detail = detail_info
                break
        
        vendors_combined.append({
            'number': req_num,
            'request_name': req_name,
            'detail': detail
        })
    
    def sort_key_number(val):
        # Try to parse an integer from the start of the identifier (handles '19-21', '3 & 4')
        try:
            s = str(val)
            # grab leading digits
            import re
            m = re.match(r"\s*(\d+)", s)
            if m:
                return (0, int(m.group(1)))
        except:
            pass
        # fallback: sort as string after numeric ones
        return (1, str(val))

    vendors_combined.sort(key=lambda x: sort_key_number(x['number']))
    
    # Collect all categories and tags and build number->name map
    all_categories = defaultdict(set)
    all_tags = defaultdict(set)
    num_to_name = {}
    
    def normalize_name(s):
        if not s:
            return ''
        ns = s.lower().replace('&', 'and')
        # keep alphanumerics and spaces
        import re
        ns = re.sub(r'[^a-z0-9\s]', '', ns)
        ns = ' '.join(ns.split())
        return ns

    def names_match(left, right):
        left_norm = normalize_name(left)
        right_norm = normalize_name(right)
        if not left_norm or not right_norm:
            return False
        if left_norm == right_norm:
            return True
        if left_norm in right_norm or right_norm in left_norm:
            return True

        import difflib
        return difflib.SequenceMatcher(a=left_norm, b=right_norm).ratio() >= 0.85
    
    for v in vendors_combined:
        # Map table number to display name (use detail title if available)
        display_name = v['detail']['title'] if v['detail'] else v['request_name']
        num_to_name[v['number']] = display_name

        if v['detail']:
            cats = parse_categories(v['detail']['categories'])
            for cat in cats:
                if cat:
                    all_categories[cat].add(v['number'])
            
            tags = parse_tags(v['detail']['tags'])
            for tag in tags:
                if tag:
                    all_tags[tag].add(v['number'])
    
    # Generate markdown content
    md_content = """# Disability Community Resource Fair 2026 - Vendor Directory

This is the comprehensive vendor directory for the 2026 Disability Community Resource Fair. Vendors are listed by table number, with an index organized by service categories and age groups served.

---

## Vendor Listings by Table Number

"""
    
    for v in vendors_combined:
        title = v['detail']['title'] if v['detail'] else v['request_name']
        
        md_content += f"\n### {v['number']}. {title}\n\n"
        
        if v['detail']:
            # Add categories
            cats = parse_categories(v['detail']['categories'])
            if cats:
                cats_str = ', '.join([c for c in cats if c])
                md_content += f"**Services:** {cats_str}\n\n"
            
            # Add tags (age groups)
            tags = parse_tags(v['detail']['tags'])
            if tags:
                tags_str = ', '.join([t for t in tags if t])
                md_content += f"**Age Groups:** {tags_str}\n\n"
            
            # Add full content (no truncation)
            content = v['detail']['content'].strip()
            if content:
                md_content += f"{content}\n\n"
    
    # Add index sections
    md_content += "\n---\n\n## Index by Service Category\n\n"
    
    for category in sorted(all_categories.keys()):
        vendors = sorted(all_categories[category])
        if vendors:
            md_content += f"\n### {category}\n"
            entries = [f"{n}. {num_to_name.get(n, '')}" for n in vendors]
            md_content += f"**Vendors:** {', '.join(entries)}\n"
    
    md_content += "\n---\n\n## Index by Age Group Served\n\n"
    
    for tag in sorted(all_tags.keys()):
        vendors = sorted(all_tags[tag])
        if vendors:
            md_content += f"\n### {tag}\n"
            entries = [f"{n}. {num_to_name.get(n, '')}" for n in vendors]
            md_content += f"**Vendors:** {', '.join(entries)}\n"
    
    # Write the directory to file (HTML if requested)
    if str(output_path).lower().endswith('.html'):
        html_parts = ["<!doctype html>", "<html lang=\"en\">", "<head>", "<meta charset=\"utf-8\">",
                      "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">",
                      f"<title>Disability Community Resource Fair 2026 - Vendor Directory</title>",
                      "<style>",
                      "body{font-family:Arial,Helvetica,sans-serif;margin:20px;color:#111}",
                      "h1,h2,h3{margin:0 0 8px 0}",
                      "section{margin-bottom:18px}",
                      "table{border-collapse:collapse;width:100%}",
                      "th,td{border:1px solid #333;padding:6px;text-align:left;vertical-align:top}",
                      "@media print{body{margin:6mm} a[href]:after{content:\"\"}}",
                      "</style>",
                      "</head>", "<body>"]

        html_parts.append("<h1>Disability Community Resource Fair 2026 - Vendor Directory</h1>")
        html_parts.append("<p>Vendors are listed by table number, with services and full descriptions.</p>")
        html_parts.append("<hr>")
        html_parts.append("<section id=\"listings\">")

        for v in vendors_combined:
            title = v['detail']['title'] if v['detail'] else v['request_name']
            html_parts.append(f"<h3>{v['number']}. {title}</h3>")
            if v['detail']:
                cats = parse_categories(v['detail']['categories'])
                if cats:
                    cats_str = ', '.join([c for c in cats if c])
                    html_parts.append(f"<p><strong>Services:</strong> {cats_str}</p>")
                tags = parse_tags(v['detail']['tags'])
                if tags:
                    tags_str = ', '.join([t for t in tags if t])
                    html_parts.append(f"<p><strong>Age Groups:</strong> {tags_str}</p>")
                content = v['detail']['content'].strip()
                if content:
                    # simple paragraph wrapping; preserve newlines
                    paragraphs = [f"<p>{p.strip()}</p>" for p in content.split('\n\n') if p.strip()]
                    html_parts.extend(paragraphs)

        html_parts.append("</section>")

        # Index by service category
        html_parts.append("<hr>")
        html_parts.append("<section id=\"index-services\">")
        html_parts.append("<h2>Index by Service Category</h2>")
        for category in sorted(all_categories.keys()):
            vendors = sorted(all_categories[category])
            if vendors:
                html_parts.append(f"<h3>{category}</h3>")
                html_parts.append('<p><strong>Vendors:</strong> ' + ', '.join([f"{n}. {num_to_name.get(n, '')}" for n in vendors]) + '</p>')
        html_parts.append("</section>")

        # Index by age group
        html_parts.append("<section id=\"index-age\">")
        html_parts.append("<h2>Index by Age Group Served</h2>")
        for tag in sorted(all_tags.keys()):
            vendors = sorted(all_tags[tag])
            if vendors:
                html_parts.append(f"<h3>{tag}</h3>")
                html_parts.append('<p><strong>Vendors:</strong> ' + ', '.join([f"{n}. {num_to_name.get(n, '')}" for n in vendors]) + '</p>')
        html_parts.append("</section>")

        html_parts.append("</body></html>")

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(html_parts))
    else:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(md_content)

    # Generate missing vendors report (two-way):
    # 1) Vendors present in requests but without a markdown file
    # 2) Vendors present in _posts but not referenced in the requests/logistics
    missing = [v for v in vendors_combined if not v['detail']]

    vendor_map_conflicts = []
    if vendor_map_path:
        vendor_map_entries = read_vendor_map_entries(vendor_map_path)
        logistics_entries = []

        for req_name, req_num in vendor_requests.items():
            logistics_entries.append({
                'display_name': req_name,
                'normalized_name': normalize_name(req_name),
                'number_values': parse_table_number_set(req_num),
            })

        for map_entry in vendor_map_entries:
            matching_logistics = [entry for entry in logistics_entries if names_match(map_entry['display_name'], entry['display_name'])]
            best_match = None
            best_score = 0.0

            if matching_logistics:
                import difflib

                for entry in matching_logistics:
                    score = difflib.SequenceMatcher(a=normalize_name(map_entry['display_name']), b=normalize_name(entry['display_name'])).ratio()
                    if score > best_score:
                        best_score = score
                        best_match = entry

            if best_match is not None:
                if map_entry['number_values'] != best_match['number_values']:
                    vendor_map_conflicts.append({
                        'map_name': map_entry['display_name'],
                        'map_numbers': format_table_number_set(map_entry['number_values']),
                        'logistics_name': best_match['display_name'],
                        'logistics_numbers': format_table_number_set(best_match['number_values']),
                        'conflict_type': 'vendor matched, table numbers differ',
                    })
                continue

            overlapping_logistics = [entry for entry in logistics_entries if map_entry['number_values'] & entry['number_values']]
            if overlapping_logistics:
                overlap_entry = sorted(overlapping_logistics, key=lambda entry: entry['display_name'])[0]
                vendor_map_conflicts.append({
                    'map_name': map_entry['display_name'],
                    'map_numbers': format_table_number_set(map_entry['number_values']),
                    'logistics_name': overlap_entry['display_name'],
                    'logistics_numbers': format_table_number_set(overlap_entry['number_values']),
                    'conflict_type': 'table number overlaps a different vendor',
                })

    # Build normalized name maps for comparison
    logistics_names = set()
    logistics_norm_to_original = {}
    for req_name in vendor_requests.keys():
        n = normalize_name(req_name)
        logistics_names.add(n)
        logistics_norm_to_original[n] = req_name

    posts_names = set()
    posts_norm_to_original = {}
    for pd_name, pd in vendor_details.items():
        # prefer the title if present
        display = pd.get('title') or pd_name
        n = normalize_name(display)
        posts_names.add(n)
        posts_norm_to_original[n] = (display, pd.get('filename'))

    matched_posts_norm = set()
    for p_norm in posts_names:
        display, _filename = posts_norm_to_original.get(p_norm, ('', ''))
        if any(names_match(req_name, display) for req_name in vendor_requests.keys()):
            matched_posts_norm.add(p_norm)

    posts_not_in_logistics_norm = sorted([n for n in posts_names if n not in matched_posts_norm])

    # Attempt fuzzy reconciliation for missing logistics entries vs posts
    import difflib
    reconciled = []
    unresolved_missing = []
    for m in missing:
        req = m['request_name']
        req_norm = normalize_name(req)
        # find best match in posts
        best = None
        best_score = 0.0
        for p_norm in posts_names:
            score = difflib.SequenceMatcher(a=req_norm, b=p_norm).ratio()
            if score > best_score:
                best_score = score
                best = p_norm

        # consider it reconciled if similarity is high or one name is substring of the other
        if best and (best_score >= 0.72 or best in req_norm or req_norm in best):
            display, filename = posts_norm_to_original.get(best, ('', ''))
            reconciled.append({'request_number': m['number'], 'request_name': req, 'matched_post': display, 'filename': filename, 'score': round(best_score, 2)})
        else:
            unresolved_missing.append(m)

    # Replace missing with unresolved_missing for the report
    missing = unresolved_missing

    if output_missing_path is not None:
        # If the user requested an HTML report, produce a printer-friendly HTML
        if str(output_missing_path).lower().endswith('.html'):
            html = []
            html.append('<!doctype html>')
            html.append('<html lang="en">')
            html.append('<head>')
            html.append('<meta charset="utf-8">')
            html.append('<meta name="viewport" content="width=device-width,initial-scale=1">')
            html.append('<title>Missing Vendor Detail Files</title>')
            html.append('<style>')
            html.append('body{font-family:Arial,Helvetica,sans-serif;margin:20px;color:#111}');
            html.append('h1,h2{margin:0 0 10px 0}');
            html.append('table{border-collapse:collapse;width:100%;margin-top:8px}');
            html.append('th,td{border:1px solid #333;padding:6px;text-align:left}');
            html.append('@media print{body{margin:6mm} a[href]:after{content:""}}');
            html.append('</style>')
            html.append('</head>')
            html.append('<body>')
            html.append('<h1>Missing Vendor Detail Files</h1>')
            html.append('<p>This report lists two sets of mismatches between the fair logistics sheet and the <code>_posts</code> vendor detail files.</p>')

            # Section A
            html.append('<h2>A. In logistics but missing <code>_posts</code> file</h2>')
            html.append('<p>Vendors present in fair logistics but missing a corresponding <code>_posts</code> markdown file.</p>')
            html.append('<table>')
            html.append('<thead><tr><th style="width:15%">Table Number</th><th>Vendor Name</th></tr></thead>')
            html.append('<tbody>')
            for m in missing:
                if not m['request_name'] or m['request_name'].strip().lower() == 'empty':
                    continue
                html.append(f"<tr><td>{m['number']}</td><td>{m['request_name']}</td></tr>")
            html.append('</tbody></table>')

            # Section B
            html.append('<h2>B. In <code>_posts</code> but not in logistics</h2>')
            html.append('<p>Vendors that have a <code>_posts</code> detail file but were not found in the fair logistics vendor requests.</p>')
            html.append('<table>')
            html.append('<thead><tr><th>Vendor Filename</th><th>Vendor Name</th></tr></thead>')
            html.append('<tbody>')
            for n in posts_not_in_logistics_norm:
                display, filename = posts_norm_to_original.get(n, ('', ''))
                if not display or display.strip().lower() == 'empty':
                    continue
                html.append(f"<tr><td>{filename or ''}</td><td>{display}</td></tr>")
            html.append('</tbody></table>')

            if vendor_map_conflicts:
                html.append('<h2>C. In vendor_map but number conflicts with logistics</h2>')
                html.append('<p>Vendors that appear in both workbooks but have different table numbers, or share a table number with a different vendor.</p>')
                html.append('<table>')
                html.append('<thead><tr><th>Vendor Map Name</th><th>Vendor Map Number(s)</th><th>Fair Logistics Name</th><th>Fair Logistics Number(s)</th><th>Conflict Type</th></tr></thead>')
                html.append('<tbody>')
                for conflict in vendor_map_conflicts:
                    html.append(
                        '<tr>'
                        f"<td>{conflict['map_name']}</td>"
                        f"<td>{conflict['map_numbers']}</td>"
                        f"<td>{conflict['logistics_name']}</td>"
                        f"<td>{conflict['logistics_numbers']}</td>"
                        f"<td>{conflict['conflict_type']}</td>"
                        '</tr>'
                    )
                html.append('</tbody></table>')

            # Print-on-load script
            html.append('<script>window.addEventListener("load", function(){window.print();});</script>')
            html.append('</body></html>')

            with open(output_missing_path, 'w', encoding='utf-8') as r:
                r.write('\n'.join(html))
        else:
            # Fallback: write markdown as before
            rpt_lines = ["# Missing Vendor Detail Files\n\n",
                         "This report lists two sets of mismatches between the fair logistics sheet and the `_posts` vendor detail files.\n\n",
                         "## A. In logistics but missing `_posts` file\n\n",
                         "Vendors present in fair logistics but missing a corresponding `_posts` markdown file.\n\n",
                         "| Table Number | Vendor Name |\n",
                         "|---:|---|\n"]
            for m in missing:
                if not m['request_name'] or m['request_name'].strip().lower() == 'empty':
                    continue
                rpt_lines.append(f"| {m['number']} | {m['request_name']} |\n")

            rpt_lines.extend(["\n## B. In `_posts` but not in logistics\n\n",
                              "Vendors that have a `_posts` detail file but were not found in the fair logistics vendor requests.\n\n",
                              "| Vendor Filename | Vendor Name |\n",
                              "|---|---|\n"])

            for n in posts_not_in_logistics_norm:
                display, filename = posts_norm_to_original.get(n, ('', ''))
                if not display or display.strip().lower() == 'empty':
                    continue
                rpt_lines.append(f"| {filename or ''} | {display} |\n")

            if vendor_map_conflicts:
                rpt_lines.extend(["\n## C. In vendor_map but number conflicts with logistics\n\n",
                                  "Vendors that appear in both workbooks but have different table numbers, or share a table number with a different vendor.\n\n",
                                  "| Vendor Map Name | Vendor Map Number(s) | Fair Logistics Name | Fair Logistics Number(s) | Conflict Type |\n",
                                  "|---|---|---|---|---|\n"])

                for conflict in vendor_map_conflicts:
                    rpt_lines.append(
                        f"| {conflict['map_name']} | {conflict['map_numbers']} | {conflict['logistics_name']} | {conflict['logistics_numbers']} | {conflict['conflict_type']} |\n"
                    )

            with open(output_missing_path, 'w', encoding='utf-8') as r:
                r.writelines(rpt_lines)

    return {
        'total_vendors': len(vendors_combined),
        'service_categories': len(all_categories),
        'age_groups': len(all_tags),
        'file_size': len(md_content),
        'missing_count': len(missing),
        'missing_list': missing,
        'posts_not_in_logistics_count': len(posts_not_in_logistics_norm),
        'posts_not_in_logistics_list': posts_not_in_logistics_norm,
        'vendor_map_conflicts_count': len(vendor_map_conflicts),
        'vendor_map_conflicts_list': vendor_map_conflicts,
    }

def main():
    """Main execution function."""
    import sys
    
    # Get base directory (parent of this script)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, '_data')
    posts_dir = os.path.join(base_dir, '_posts')
    
    # File paths
    fair_logistics_path = os.path.join(data_dir, 'fair_logistics.xlsx')
    vendor_requests_csv = os.path.join(data_dir, '2026_vendor_requests.csv')
    vendor_directory_md = os.path.join(data_dir, '2026_vendor_directory.html')
    vendor_missing_report = os.path.join(data_dir, '2026_missing_vendors.md')
    vendor_map_path = os.path.join(data_dir, 'vendor_map.xlsx')
    
    print("=" * 70)
    print("2026 Disability Community Resource Fair - Vendor Directory Generator")
    print("=" * 70)
    
    # Step 1: Extract vendors from Excel
    print("\n[1/4] Extracting vendors from fair_logistics.xlsx...")
    try:
        count = extract_vendors_from_excel(fair_logistics_path, vendor_requests_csv)
        print(f"      ✓ Extracted {count} rows to {vendor_requests_csv}")
    except Exception as e:
        print(f"      ✗ Error: {e}")
        sys.exit(1)
    
    # Step 2: Read vendor requests
    print("\n[2/4] Reading vendor requests from CSV...")
    try:
        vendor_requests = read_vendor_requests(vendor_requests_csv)
        print(f"      ✓ Loaded {len(vendor_requests)} vendor requests")
    except Exception as e:
        print(f"      ✗ Error: {e}")
        sys.exit(1)
    
    # Step 3: Read vendor details
    print("\n[3/4] Reading vendor details from markdown files...")
    try:
        vendor_details = read_vendor_details(posts_dir)
        print(f"      ✓ Loaded {len(vendor_details)} vendor detail files")
    except Exception as e:
        print(f"      ✗ Error: {e}")
        sys.exit(1)
    
    # Step 4: Generate vendor directory
    print("\n[4/4] Generating vendor directory...")
    try:
        stats = generate_vendor_directory(vendor_requests, vendor_details, vendor_directory_md, output_missing_path=vendor_missing_report, vendor_map_path=vendor_map_path)
        print(f"      ✓ Generated {vendor_directory_md}")
        print(f"\n      Statistics:")
        print(f"        - Vendors: {stats['total_vendors']}")
        print(f"        - Service Categories: {stats['service_categories']}")
        print(f"        - Age Groups: {stats['age_groups']}")
        print(f"        - File Size: {stats['file_size']:,} bytes")
        print(f"        - Missing vendor detail files: {stats.get('missing_count', 0)}")
        print(f"        - Vendor map conflicts: {stats.get('vendor_map_conflicts_count', 0)}")
    except Exception as e:
        print(f"      ✗ Error: {e}")
        sys.exit(1)
    
    print("\n" + "=" * 70)
    print("✓ Vendor directory generation complete!")
    print("=" * 70)

if __name__ == '__main__':
    main()
