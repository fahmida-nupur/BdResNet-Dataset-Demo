import os
import requests
import time
import re
from dotenv import load_dotenv
load_dotenv()  # loads variables from .env file

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

HAS_DRIVE = False


def reconstruct_abstract(inverted_index):
    """
    Reconstruct full abstract text from OpenAlex's abstract_inverted_index.
    """
    if not inverted_index or not isinstance(inverted_index, dict):
        return ""
    
    position_word_map = []
    for word, positions in inverted_index.items():
        for pos in positions:
            position_word_map.append((pos, word))
            
    position_word_map.sort(key=lambda x: x[0])
    return " ".join([word for pos, word in position_word_map])


def clean_text(val):
    """
    Clean non-printable control characters to prevent Excel XML corruption
    and trim text to Excel cell length limit.
    """
    if isinstance(val, str):
        val = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]', '', val)
        if len(val) > 32000:
            val = val[:32000] + "... [TRUNCATED]"
    return val


def sanitize_sheet_name(name):
    r"""
    Sanitize Excel sheet names by removing invalid characters (\, /, ?, *, :, [, ])
    and restricting length to 31 characters.
    """
    clean_name = re.sub(r'[\\/*?:\[\]]', '_', name)
    return clean_name[:31]


def fetch_papers_for_institution(institution_id, email, max_papers=100):
    """
    Fetch up to max_papers for a single institution from OpenAlex API and return list of records.
    Set max_papers=None or float('inf') to fetch ALL papers.
    """
    base_url = "https://api.openalex.org/works"
    
    inst_id_clean = institution_id.split("/")[-1].upper()
    filter_param = f"authorships.institutions.lineage:{inst_id_clean}"
    
    cursor = "*"
    per_page = min(200, max_papers) if max_papers else 200
    papers = []
    
    print(f"\n[+] Fetching up to {max_papers if max_papers else 'all'} papers for Institution: {inst_id_clean}...")
    
    while cursor:
        if max_papers and len(papers) >= max_papers:
            break

        params = {
            'filter': filter_param,
            'per-page': per_page,
            'cursor': cursor,
            'mailto': email
        }
        
        headers = {'User-Agent': f'OpenAlexPaperFetcher/1.0 (mailto:{email})'}
        
        try:
            response = requests.get(base_url, params=params, headers=headers, timeout=30)
            
            if response.status_code != 200:
                print(f"Error HTTP {response.status_code}: {response.text}")
                break
                
            data = response.json()
            results = data.get('results', [])
            meta = data.get('meta', {})
            total_count = meta.get('count', 0)
            
            if not results:
                break
                
            for work in results:
                if max_papers and len(papers) >= max_papers:
                    break

                authorships = work.get('authorships', [])
                author_data = []
                for a in authorships:
                    if a.get('author'):
                        name = a['author'].get('display_name', '')
                        institutions = a.get('institutions', [])
                        affil = '; '.join(
                            inst.get('display_name', '')
                            for inst in institutions
                            if inst.get('display_name')
                        )
                        author_data.append((name, affil))

                abstract_text = reconstruct_abstract(work.get('abstract_inverted_index'))

                # --- Primary location / source metadata ---
                primary_loc = work.get('primary_location') or {}
                source = primary_loc.get('source') or {}
                venue_name      = source.get('display_name', '')
                publisher_name  = source.get('host_organization_name', '')
                source_type     = source.get('type', '')        # e.g. journal, conference, repository
                issn_l          = source.get('issn_l', '')
                is_in_doaj      = source.get('is_in_doaj', '')

                # --- Work type ---
                work_type = work.get('type', '')  # e.g. article, book-chapter, proceedings-article

                # --- Open Access ---
                oa_info    = work.get('open_access') or {}
                oa_status  = oa_info.get('oa_status', '')   # gold, green, bronze, closed, hybrid

                # --- Bibliographic details ---
                biblio     = work.get('biblio') or {}
                volume     = biblio.get('volume', '')
                issue      = biblio.get('issue', '')
                first_page = biblio.get('first_page', '')
                last_page  = biblio.get('last_page', '')

                # --- Metrics / ranking proxies ---
                fwci = work.get('fwci', '')
                cnp  = work.get('citation_normalized_percentile') or {}
                cnp_value      = cnp.get('value', '')
                is_top_1_pct   = cnp.get('is_in_top_1_percent', '')
                is_top_10_pct  = cnp.get('is_in_top_10_percent', '')

                # --- Primary topic hierarchy ---
                primary_topic = work.get('primary_topic') or {}
                topic_name    = primary_topic.get('display_name', '')
                subfield_name = (primary_topic.get('subfield') or {}).get('display_name', '')
                field_name    = (primary_topic.get('field') or {}).get('display_name', '')
                domain_name   = (primary_topic.get('domain') or {}).get('display_name', '')

                # --- Keywords (top 10 by score, semicolon-separated) ---
                kw_list  = work.get('keywords') or []
                keywords = '; '.join(
                    kw.get('display_name', '')
                    for kw in sorted(kw_list, key=lambda k: k.get('score', 0), reverse=True)[:10]
                    if kw.get('display_name')
                )

                raw_row = {
                    'Landing Page URL':        work.get('id', ''),
                    'Title':                   work.get('title', ''),
                    'Abstract':                abstract_text,
                    'DOI':                     work.get('doi', ''),
                    'Publication Year':        work.get('publication_year', ''),
                    'Publication Date':        work.get('publication_date', ''),
                    'Citation Count':          work.get('cited_by_count', 0),
                    # Venue / publisher
                    'Venue / Journal Name':    venue_name,
                    'ISSN':                    issn_l,
                    'Publisher':               publisher_name,
                    'Source Type':             source_type,   # journal | conference | repository | book | …
                    'Work Type':               work_type,     # article | book-chapter | proceedings-article | …
                    'In DOAJ':                 is_in_doaj,
                    # Open access
                    'OA Status':               oa_status,     # gold | green | bronze | hybrid | closed
                    # Bibliographic
                    'Volume':                  volume,
                    'Issue':                   issue,
                    'First Page':              first_page,
                    'Last Page':               last_page,
                    # Metrics / ranking proxies
                    'FWCI':                    fwci,
                    'Citation Percentile':     cnp_value,
                    'Top 1% Cited':            is_top_1_pct,
                    'Top 10% Cited':           is_top_10_pct,
                    # Topic hierarchy
                    'Primary Topic':           topic_name,
                    'Subfield':                subfield_name,
                    'Field':                   field_name,
                    'Domain':                  domain_name,
                    # Keywords
                    'Keywords':                keywords,
                }

                # Clean string fields; keep numeric fields as-is
                cleaned_row = {
                    k: (clean_text(v) if isinstance(v, str) else v)
                    for k, v in raw_row.items()
                }
                # Store authors as structured list (expanded later into columns)
                cleaned_row['_authors'] = [
                    (clean_text(name), clean_text(affil))
                    for name, affil in author_data
                ]
                papers.append(cleaned_row)
                
            print(f"    Progress: {len(papers)} / {min(total_count, max_papers if max_papers else total_count)} papers downloaded...")
            cursor = meta.get('next_cursor')
            time.sleep(0.1)
            
        except Exception as e:
            print(f"Error occurred: {e}")
            break
            
    print(f"    Total papers collected for {inst_id_clean}: {len(papers)}")
    return papers


def process_institutions(institutions, email, excel_output="all_institutions_data.xlsx", max_papers_per_inst=100):
    """
    Fetch metadata for multiple institutions and add a separate Sheet
    for each institution into a single Excel workbook.

    Skip logic: if the Excel file already exists and a sheet for an institution
    is already present, that institution is skipped (no re-download).
    Only new institutions are fetched and appended.
    """
    if not HAS_PANDAS:
        print("[!] pandas is not installed. Cannot create Excel file.")
        return

    # --- Load existing sheets to determine what to skip ---
    existing_sheets = {}
    if os.path.exists(excel_output):
        try:
            existing_wb = pd.ExcelFile(excel_output, engine='openpyxl')
            for sheet in existing_wb.sheet_names:
                existing_sheets[sheet] = pd.read_excel(existing_wb, sheet_name=sheet)
            existing_wb.close()
            print(f"Found existing workbook '{excel_output}' with {len(existing_sheets)} sheet(s): {list(existing_sheets.keys())}")
        except Exception as e:
            print(f"Could not read existing workbook ({e}). Starting fresh.")

    # --- Open writer for output ---
    try:
        excel_writer = pd.ExcelWriter(excel_output, engine='openpyxl')
    except PermissionError:
        print(f"\nERROR: '{excel_output}' is open in Excel. Please close Excel and re-run.")
        return
    except Exception as e:
        print(f"\nCould not initialize ExcelWriter: {e}")
        return

    # --- Copy existing sheets into the new writer first ---
    for sheet_name, df_existing in existing_sheets.items():
        df_existing.to_excel(excel_writer, sheet_name=sheet_name, index=False)

    # --- Process each institution ---
    for inst_name, inst_id in institutions.items():
        sheet_title = sanitize_sheet_name(inst_name)

        # Skip if sheet already exists
        if sheet_title in existing_sheets:
            print(f"Skipping '{inst_name}' — sheet already exists.")
            continue

        papers = fetch_papers_for_institution(inst_id, email, max_papers=max_papers_per_inst)

        if not papers:
            print(f"No papers found for '{inst_name}'. Skipping...")
            continue


        max_authors = max((len(p.get('_authors', [])) for p in papers), default=0)

        rows = []
        for paper in papers:
            row = {k: v for k, v in paper.items() if k != '_authors'}
            for i, (name, affil) in enumerate(paper.get('_authors', []), start=1):
                row[f'Author {i}'] = name
                row[f'Author {i} Affiliation'] = affil
            rows.append(row)

        df = pd.DataFrame(rows)
        df.to_excel(excel_writer, sheet_name=sheet_title, index=False)
        print(f"  --> Added to Excel Sheet: '{sheet_title}' ({len(papers)} rows, up to {max_authors} authors)")

    if excel_writer:
        try:
            excel_writer.close()
            print(f"\n[+] Excel workbook saved successfully: '{excel_output}'!")
        except PermissionError:
            print(f"\n[!] ERROR: Could not save '{excel_output}' because it is open in Excel. Close Excel and try again.")
        except Exception as e:
            print(f"\n[!] Failed to save Excel workbook: {e}")


if __name__ == "__main__":
    INSTITUTIONS = {
        "Green University of Bangladesh": "I2802576217",
        "American International University - Bangladesh": "I103434671",
        "North South University": "I157386601",
        "University of Dhaka": "I205746353",
        "BRAC University": "I5518804",
        "Chittagong Medical College": "I2653708",
        "Khulna University of Engineering and Technology": "I1307585291",
        "Shahjalal University of Science and Technology": "I133190216",
        "Chandpur Science and Technology University": "I4402554010"
    }

    YOUR_EMAIL = os.getenv("OPENALEX_EMAIL", "")
    if not YOUR_EMAIL:
        print("WARNING: OPENALEX_EMAIL not set in .env file. Copy .env.example to .env and fill in your email.")

    EXCEL_OUTPUT = "all_institutions_data.xlsx"

    process_institutions(
        institutions=INSTITUTIONS,
        email=YOUR_EMAIL,
        excel_output=EXCEL_OUTPUT,
        max_papers_per_inst=100
    )
