## Auto Document Generator

Auto Document Generator is a Python project that builds **beautiful, formatted Google Docs reports automatically** using the Google Docs API and pandas data structures.  
It includes a reusable `docbuilder` library plus ready‑to‑run scripts that generate a complete, styled report from dummy or real data.

### Features

- **High-level Google Docs client** (`docbuilder.GoogleDocsClient`) for:
  - Creating / copying documents
  - Inserting and styling text, headings, titles, lists
  - Inserting and styling tables from `pandas.DataFrame`
  - Adding page / section breaks and images
- **Report generator script** (`fill_data.py`) that:
  - Builds a full “Company Performance Report” with multiple sections
  - Inserts several styled tables from pandas data
  - Applies consistent typography and colors
  - Can optionally copy a template before populating it
  - Can optionally send a Slack notification with the document link
- **Slack integration** (`slack_alerts.alerting`) to push a message to one or more Slack channels with the generated Doc URL.

---

### Project Structure

- `docbuilder/` – Reusable library for working with Google Docs
  - `auth.py` – Service account authentication and Docs/Drive client creation
  - `docs_client.py` – High-level wrapper around the Google Docs API
  - `styles.py` – Text, paragraph, table styles, colors, fonts, etc.
  - `elements.py` – Object model for paragraphs, tables, lists, etc.
  - `utils.py` – Request builders, parsers, helpers
- `fill_data.py` – Main entrypoint that generates a sample Google Docs report
- `get_doc_content.py` – Utility to inspect table cell indices in an existing Doc
- `slack_alerts/alerting.py` – Helper to send Slack notifications with a Doc link
- `requirements.txt` – Python dependencies

---

### Prerequisites

- **Python**: 3.9+ recommended
- **Google Cloud project** with:
  - Google Docs API enabled
  - Google Drive API enabled
- **Service account** with access to the target Google Doc:
  1. Create a service account in Google Cloud.
  2. Generate a JSON key and download it.
  3. Save it next to this project as `service.json` (or anywhere and pass the path via CLI).
  4. Share your Google Docs template with the service account email (e.g. `service-account@project-id.iam.gserviceaccount.com`) with at least **Editor** access.

> **Security tip:** Do **not** commit `service.json` to Git. It is already ignored by `.gitignore`.

---

### Installation

From the project root:

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
```

---

### Configuration

You can configure the script via **environment variables** or **CLI flags**:

- `GOOGLE_SERVICE_ACCOUNT_FILE` – Path to service account JSON (default: `service.json`)
- `DOC_TEMPLATE_ID` – Google Docs template ID to use by default

You can find the document ID in the URL:  
`https://docs.google.com/document/d/<DOCUMENT_ID>/edit`

---

### Usage – Generate a Sample Report

The main script is `fill_data.py`. It generates a complete, styled “Company Performance Report” into a Google Doc.

Basic run (uses defaults and environment variables if set):

```bash
python fill_data.py
```

With explicit options:

```bash
python fill_data.py \
  --service-account-file ./service.json \
  --template-id YOUR_TEMPLATE_DOC_ID \
  --document-title "Auto-generated Performance Report" \
  --copy-template
```

Key options:

- `--service-account-file / -s` – Path to service account JSON.
- `--template-id / -t` – ID of an existing Google Doc to use as a template.
- `--document-title / -d` – Optional title to set on the target document.
- `--copy-template / -c` – If set, the script first **copies** the template and then populates the copy (recommended for reports).

---

### Optional: Slack Notification

You can send a Slack message with the final document link once generation is complete.

```bash
python fill_data.py \
  --template-id YOUR_TEMPLATE_DOC_ID \
  --slack-channel "#my-team-channel" \
  --slack-channel "#another-channel"
```

Options:

- `--slack-channel` – Slack channel to notify (can be used multiple times).
- `--slack-endpoint` – Override the default Slack endpoint defined in `slack_alerts/alerting.py`.

The helper in `slack_alerts.alerting` sends a message like:

> “Please find the Delivery+ Deployment Safety MBR document for the current month below: \<link\>”

---

### Utility: Inspecting Table Cell Indices

`get_doc_content.py` is a small utility to inspect the start indices of each cell in the first table of a Google Doc and dump them to JSON:

```bash
python get_doc_content.py
```

It writes a file named `cell_indexes_output.json` in the project root, which can be useful when building more advanced table-manipulation logic.

