from google.oauth2 import service_account
from googleapiclient.discovery import build
import json
import os

# -----------------------
# Step 1: Set up config
# -----------------------
SERVICE_ACCOUNT_FILE = 'service.json'
DOCUMENT_ID = '1Hfq2he3cEkS71PrLmkrt1biX-M08wbU8WoRqjdCcDdI'

SCOPES = ['https://www.googleapis.com/auth/documents']
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
docs_service = build('docs', 'v1', credentials=credentials)

# -----------------------
# Step 3: Get table cell indexes
# -----------------------
def get_table_cell_indexes(doc_service, doc_id):
    doc = doc_service.documents().get(documentId=doc_id).execute()

    # Find the first table
    table = next((c for c in doc['body']['content'] if 'table' in c), None)
    if not table:
        raise Exception("No table found.")

    # Extract start indexes of each cell
    cell_indexes = []
    for row in table['table']['tableRows']:
        row_indexes = []
        for cell in row['tableCells']:
            # Ensure cell has content
            content = cell.get('content', [])
            if content and 'startIndex' in content[0]:
                row_indexes.append(content[0]['startIndex'])
            else:
                row_indexes.append(None)
        cell_indexes.append(row_indexes)

    return cell_indexes


cell_indexes = get_table_cell_indexes(docs_service, DOCUMENT_ID)

filename = "cell_indexes_output.json"
file_path = os.path.join(os.path.dirname(__file__), filename)

with open(file_path, "w") as f:
    json.dump(cell_indexes, f, indent=4)

print(f"cell_indexes saved to {file_path}")


