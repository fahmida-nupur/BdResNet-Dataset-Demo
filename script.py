import os
import requests
# import csv  # CSV export disabled; uncomment to re-enable CSV saving
import time
import re

from dotenv import load_dotenv
load_dotenv()  # loads variables from .env file

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

# --- Google Drive Integration (disabled) ---
# Uncomment the block below to enable Drive sync.
# try:
#     from googleapiclient.discovery import build
#     from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
#     from google.oauth2 import service_account
#     HAS_DRIVE = True
# except ImportError:
#     HAS_DRIVE = False
# --- End Google Drive Integration ---
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

                raw_row = {
                    'Landing Page URL': work.get('primary_location', {}).get('landing_page_url', ''),
                    'Title': work.get('title', ''),
                    'Abstract': abstract_text,
                    'DOI': work.get('doi', ''),
                    'Publication Year': work.get('publication_year', ''),
                    'Citation Count': work.get('cited_by_count', 0),
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

        # --- For CSV Export  ---
        # Uncomment the block below to re-enable saving each institution as a separate CSV file.
        # clean_file_name = re.sub(r'[\\/*?:\[\]\s]', '_', inst_name)
        # csv_filename = f"{clean_file_name}_papers.csv"
        # fieldnames = papers[0].keys()
        # try:
        #     with open(csv_filename, mode='w', newline='', encoding='utf-8-sig') as f:
        #         writer = csv.DictWriter(f, fieldnames=fieldnames)
        #         writer.writeheader()
        #         writer.writerows(papers)
        #     print(f"  --> Saved CSV: '{csv_filename}' ({len(papers)} rows)")
        # except PermissionError:
        #     print(f"  [!] Could not save '{csv_filename}' (file open in another program).")




        # --- Expand _authors into individual columns at the far right ---
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


# --- Google Drive Functions (disabled) ---
# Uncomment the blocks below to enable Drive sync.
#
# SCOPES = ['https://www.googleapis.com/auth/drive']
#
#
# def get_drive_service():
#     """
#     Build and return an authenticated Google Drive service client
#     using the Service Account JSON key defined in the .env file.
#     Returns None if credentials are not configured or unavailable.
#     """
#     if not HAS_DRIVE:
#         print("[!] Google API libraries not installed. Run: pip install -r requirements.txt")
#         return None
#
#     json_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
#     if not json_path or not os.path.exists(json_path):
#         print("[!] GOOGLE_SERVICE_ACCOUNT_JSON not set or file not found. Skipping Drive sync.")
#         return None
#
#     try:
#         creds = service_account.Credentials.from_service_account_file(json_path, scopes=SCOPES)
#         return build('drive', 'v3', credentials=creds)
#     except Exception as e:
#         print(f"[!] Failed to authenticate with Google Drive: {e}")
#         return None
#
#
# def download_from_drive(service, folder_id, filename, dest_path):
#     """
#     Download the latest version of `filename` from a Drive folder to `dest_path`.
#     Returns the file ID if found, else None.
#     """
#     try:
#         query = f"name='{filename}' and '{folder_id}' in parents and trashed=false"
#         results = service.files().list(q=query, fields="files(id, name)").execute()
#         files = results.get('files', [])
#
#         if not files:
#             print(f"[i] No existing '{filename}' found in Drive. Starting fresh.")
#             return None
#
#         file_id = files[0]['id']
#         print(f"[i] Downloading latest '{filename}' from Google Drive...")
#         import io
#         request = service.files().get_media(fileId=file_id)
#         fh = io.FileIO(dest_path, 'wb')
#         downloader = MediaIoBaseDownload(fh, request)
#         done = False
#         while not done:
#             _, done = downloader.next_chunk()
#         fh.close()
#         print(f"[+] Download complete: '{dest_path}'")
#         return file_id
#
#     except Exception as e:
#         print(f"[!] Drive download failed: {e}")
#         return None
#
#
# def upload_to_drive(service, folder_id, file_path, existing_file_id=None):
#     """
#     Upload `file_path` to a Drive folder.
#     If `existing_file_id` is given, the existing file is replaced (updated in-place).
#     """
#     filename = os.path.basename(file_path)
#     mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
#     media = MediaFileUpload(file_path, mimetype=mime, resumable=True)
#
#     try:
#         if existing_file_id:
#             # Update the existing file (keeps same file ID and sharing settings)
#             service.files().update(
#                 fileId=existing_file_id,
#                 media_body=media
#             ).execute()
#             print(f"[+] Updated '{filename}' on Google Drive.")
#         else:
#             # Upload as a new file
#             metadata = {'name': filename, 'parents': [folder_id]}
#             service.files().create(
#                 body=metadata,
#                 media_body=media,
#                 fields='id'
#             ).execute()
#             print(f"[+] Uploaded '{filename}' to Google Drive.")
#     except Exception as e:
#         print(f"[!] Drive upload failed: {e}")
# --- End Google Drive Functions ---


if __name__ == "__main__":
    INSTITUTIONS = {
        "Green University of Bangladesh": "I2802576217",
        "American International University - Bangladesh": "I103434671",
        "North South University": "I157386601",
        "University of Dhaka": "I205746353",
        "BRAC University": "I5518804"



    }

    YOUR_EMAIL = os.getenv("OPENALEX_EMAIL", "")
    if not YOUR_EMAIL:
        print("WARNING: OPENALEX_EMAIL not set in .env file. Copy .env.example to .env and fill in your email.")

    EXCEL_OUTPUT = "all_institutions_data.xlsx"


    # --- Google Drive Settings (disabled) ---  
    # FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "")  # Drive disabled

    # --- Drive Step 1: Download latest Excel from Drive (disabled) ---
    # drive_service = get_drive_service()
    # existing_file_id = None
    # if drive_service and FOLDER_ID:
    #     existing_file_id = download_from_drive(
    #         drive_service, FOLDER_ID, EXCEL_OUTPUT, EXCEL_OUTPUT
    #     )
    # else:
    #     if not FOLDER_ID:
    #         print("[i] GOOGLE_DRIVE_FOLDER_ID not set. Running in local-only mode.")
    # --- End Drive Step 1 ---

    # ── Step 2: Fetch new institutions and update Excel ────────────────────────
    # 100 papers per institution limit (set max_papers_per_inst=None to fetch all papers)


    process_institutions(
        institutions=INSTITUTIONS,
        email=YOUR_EMAIL,
        excel_output=EXCEL_OUTPUT,
        max_papers_per_inst=100
    )

    # --- Drive Step 3: Upload updated Excel back to Drive (disabled) ---
    # if drive_service and FOLDER_ID and os.path.exists(EXCEL_OUTPUT):
    #     upload_to_drive(drive_service, FOLDER_ID, EXCEL_OUTPUT, existing_file_id)
    # --- End Drive Step 3 ---
