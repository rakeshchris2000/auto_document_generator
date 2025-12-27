"""
Google Docs API Client Module

This module contains the main client class that provides a high-level interface
for interacting with the Google Docs API, including document creation, content
manipulation, and formatting operations.
"""

from __future__ import annotations
import time
import pandas as pd
from typing import Dict, Any, List, Optional, Union, Tuple
from googleapiclient.errors import HttpError

from .auth import GoogleDocsAuth
from .elements import (
    DocumentElement, Paragraph, TextRun, Table, List as DocList,
    PageBreak, SectionBreak, DocumentBuilder, create_simple_table,
    create_bulleted_list, create_numbered_list
)
from .styles import (
    TextStyle, ParagraphStyle, TableCellStyle, Color, Dimension,
    ParagraphStyleType, TextAlignment, ListType, FontFamilies, FontSizes
)
from .utils import (
    RequestBuilder, DocumentParser, BatchRequestBuilder, generate_unique_id,
    validate_index_range, clean_text, ErrorHandler
)


class GoogleDocsClient:
    """
    High-level client for Google Docs API operations.
    
    This client provides easy-to-use methods for creating and manipulating
    Google Docs documents, including text formatting, tables, lists, and more.
    """
    
    def __init__(self, service_account_file: str):
        """
        Initialize the Google Docs client.
        
        Args:
            service_account_file (str): Path to service account JSON file
            
        Raises:
            Exception: If authentication fails
        """
        self.auth = GoogleDocsAuth(service_account_file)
        self.auth.authenticate()
        self.docs_service = self.auth.get_service()
        self.drive_service = self.auth.get_drive_service()
        self._current_document_id: Optional[str] = None
    
    def create_document(self, title: str = "Untitled Document") -> str:
        """
        Create a new Google Docs document.
        
        Args:
            title (str): Document title
            
        Returns:
            str: Document ID
            
        Raises:
            Exception: If document creation fails
        """
        try:
            document = {
                'title': title
            }
            
            doc = self.docs_service.documents().create(body=document).execute()
            document_id = doc.get('documentId')
            self._current_document_id = document_id
            
            return document_id
            
        except HttpError as e:
            raise Exception(f"Failed to create document: {e}")
    
    def get_document(self, document_id: str) -> Dict[str, Any]:
        """
        Retrieve a document by ID.
        
        Args:
            document_id (str): Document ID
            
        Returns:
            Dict[str, Any]: Document object
            
        Raises:
            Exception: If document retrieval fails
        """
        try:
            return self.docs_service.documents().get(documentId=document_id).execute()
        except HttpError as e:
            raise Exception(f"Failed to get document: {e}")

    def copy_document(self, document_id: str, title: str) -> str:
        """
        Copy a document.
        """
        try:
            return self.drive_service.files().copy(fileId=document_id, body={'name': title}).execute()
        except HttpError as e:
            raise Exception(f"Failed to copy document: {e}")        

    def update_document(self, document_id: str, requests: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Update a document with batch requests.
        
        Args:
            document_id (str): Document ID
            requests (List[Dict[str, Any]]): List of update requests
            
        Returns:
            Dict[str, Any]: Update response
            
        Raises:
            Exception: If update fails
        """
        if not requests:
            return {}
        
        try:
            body = {'requests': requests}
            return self.docs_service.documents().batchUpdate(
                documentId=document_id, body=body
            ).execute()
        except HttpError as e:
            raise Exception(f"Failed to update document: {e}")
    
    def get_document_content(self, document_id: str) -> str:
        """
        Get the text content of a document.
        
        Args:
            document_id (str): Document ID
            
        Returns:
            str: Document text content
        """
        document = self.get_document(document_id)
        return DocumentParser.extract_text_content(document)
    
    def insert_text(self, document_id: str, text: str, index: int = 1, 
                   style: Optional[TextStyle] = None) -> Dict[str, Any]:
        """
        Insert text at a specific position in the document.
        
        Args:
            document_id (str): Document ID
            text (str): Text to insert
            index (int): Position to insert at (1-based)
            style (TextStyle, optional): Text formatting
            
        Returns:
            Dict[str, Any]: Update response
        """
        text = clean_text(text)
        requests = []
        
        # Insert text first
        requests.append(RequestBuilder.insert_text(text, index))
        
        # Apply style if provided
        if style:
            start_index = index
            end_index = index + len(text)
            requests.append(RequestBuilder.update_text_style(start_index, end_index, style))
        
        return self.update_document(document_id, requests)
    
    def append_text(self, document_id: str, text: str, 
                   style: Optional[TextStyle] = None) -> Dict[str, Any]:
        """
        Append text to the end of the document.
        
        Args:
            document_id (str): Document ID
            text (str): Text to append
            style (TextStyle, optional): Text formatting
            
        Returns:
            Dict[str, Any]: Update response
        """
        document = self.get_document(document_id)
        end_index = document['body']['content'][-1]['endIndex'] - 1
        return self.insert_text(document_id, text, end_index, style)
    
    def replace_text(self, document_id: str, search_text: str, replace_text: str,
                    match_case: bool = False) -> Dict[str, Any]:
        """
        Replace all occurrences of text in the document.
        
        Args:
            document_id (str): Document ID
            search_text (str): Text to search for
            replace_text (str): Replacement text
            match_case (bool): Whether to match case
            
        Returns:
            Dict[str, Any]: Update response
        """
        request = {
            "replaceAllText": {
                "containsText": {
                    "text": search_text,
                    "matchCase": match_case
                },
                "replaceText": clean_text(replace_text)
            }
        }
        return self.update_document(document_id, [request])
    
    def format_text(self, document_id: str, start_index: int, end_index: int,
                   style: TextStyle) -> Dict[str, Any]:
        """
        Apply text formatting to a range of text.
        
        Args:
            document_id (str): Document ID
            start_index (int): Start of text range
            end_index (int): End of text range
            style (TextStyle): Text formatting to apply
            
        Returns:
            Dict[str, Any]: Update response
        """
        validate_index_range(start_index, end_index)
        request = RequestBuilder.update_text_style(start_index, end_index, style)
        return self.update_document(document_id, [request])
    
    def format_paragraph(self, document_id: str, start_index: int, end_index: int,
                        style: ParagraphStyle) -> Dict[str, Any]:
        """
        Apply paragraph formatting to a range.
        
        Args:
            document_id (str): Document ID
            start_index (int): Start of paragraph range
            end_index (int): End of paragraph range
            style (ParagraphStyle): Paragraph formatting to apply
            
        Returns:
            Dict[str, Any]: Update response
        """
        validate_index_range(start_index, end_index)
        request = RequestBuilder.update_paragraph_style(start_index, end_index, style)
        return self.update_document(document_id, [request])
    
    def insert_paragraph(self, document_id: str, text: str = "", index: int = 1,
                        text_style: Optional[TextStyle] = None,
                        paragraph_style: Optional[ParagraphStyle] = None) -> Dict[str, Any]:
        """
        Insert a paragraph with optional text and formatting.
        
        Args:
            document_id (str): Document ID
            text (str): Paragraph text
            index (int): Position to insert at
            text_style (TextStyle, optional): Text formatting
            paragraph_style (ParagraphStyle, optional): Paragraph formatting
            
        Returns:
            Dict[str, Any]: Update response
        """
        requests = []
        original_index = index
        
        # Insert text if provided
        if text:
            text = clean_text(text)
            requests.append(RequestBuilder.insert_text(text, index))
            
            # Apply text style if provided
            if text_style:
                requests.append(RequestBuilder.update_text_style(index, index + len(text), text_style))
            
            index += len(text)
        
        # Insert paragraph break
        requests.append(RequestBuilder.insert_text("\n", index))
        index += 1
        
        # Apply paragraph style if provided
        if paragraph_style:
            start_index = original_index
            end_index = index
            requests.append(RequestBuilder.update_paragraph_style(
                start_index, end_index, paragraph_style
            ))
        
        return self.update_document(document_id, requests)
    
    def add_heading(self, document_id: str, text: str, level: int = 1,
                   index: Optional[int] = None, style: Optional[TextStyle] = None) -> Dict[str, Any]:
        """
        Add a heading to the document.
        
        Args:
            document_id (str): Document ID
            text (str): Heading text
            level (int): Heading level (1-6)
            index (int, optional): Position to insert at. If None, appends to end.
            style (TextStyle, optional): Additional text formatting
            
        Returns:
            Dict[str, Any]: Update response
        """
        if level < 1 or level > 6:
            raise ValueError("Heading level must be between 1 and 6")
        
        heading_types = {
            1: ParagraphStyleType.HEADING_1,
            2: ParagraphStyleType.HEADING_2,
            3: ParagraphStyleType.HEADING_3,
            4: ParagraphStyleType.HEADING_4,
            5: ParagraphStyleType.HEADING_5,
            6: ParagraphStyleType.HEADING_6
        }
        
        paragraph_style = ParagraphStyle(
            namedStyleType=heading_types[level],
            space_below=Dimension.points(6)
        )
        
        if index is None:
            document = self.get_document(document_id)
            index = document['body']['content'][-1]['endIndex'] - 1
        
        return self.insert_paragraph(document_id, text, index, style, paragraph_style)
    
    def add_title(self, document_id: str, text: str, index: Optional[int] = None,
                 style: Optional[TextStyle] = None) -> Dict[str, Any]:
        """Add a title to the document."""
        paragraph_style = ParagraphStyle(namedStyleType=ParagraphStyleType.TITLE)
        
        if index is None:
            document = self.get_document(document_id)
            index = document['body']['content'][-1]['endIndex'] - 1
        
        return self.insert_paragraph(document_id, text, index, style, paragraph_style)
    
    def add_subtitle(self, document_id: str, text: str, index: Optional[int] = None,
                    style: Optional[TextStyle] = None) -> Dict[str, Any]:
        """Add a subtitle to the document."""
        paragraph_style = ParagraphStyle(namedStyleType=ParagraphStyleType.SUBTITLE)
        
        if index is None:
            document = self.get_document(document_id)
            index = document['body']['content'][-1]['endIndex'] - 1
        
        return self.insert_paragraph(document_id, text, index, style, paragraph_style)
    
    def insert_table(self, document_id: str, rows: int, columns: int,
                    index: Optional[int] = None) -> Dict[str, Any]:
        """
        Insert a table into the document.
        
        Args:
            document_id (str): Document ID
            rows (int): Number of rows
            columns (int): Number of columns
            index (int, optional): Position to insert at. If None, appends to end.
            
        Returns:
            Dict[str, Any]: Update response
        """
        if index is None:
            document = self.get_document(document_id)
            index = document['body']['content'][-1]['endIndex'] - 1
        
        request = RequestBuilder.insert_table(index, rows, columns)
        return self.update_document(document_id, [request])
    
    def insert_table_from_dataframe(self, document_id: str, df: pd.DataFrame,
                                 index: Optional[int] = None,
                                 header_style: Optional[TextStyle] = None,
                                 cell_style: Optional[TextStyle] = None ,
                                 table_cell_style: Optional[TableCellStyle] = None) -> Dict[str, Any]:
        """
        Insert a styled table into a Google Doc from a pandas DataFrame.

        Args:
            document_id (str): Document ID.
            df (pd.DataFrame): DataFrame to insert.
            index (int, optional): Document index to insert at.
            header_style (TextStyle, optional): Style for the header row.
            cell_style (TextStyle, optional): Style for the data rows.
            table_cell_style (TableCellStyle, optional): Style for the table cells.
        Returns:
            Dict[str, Any]: Response from the table insertion request.
        """
        if df.empty:
            raise ValueError("DataFrame is empty")

        # Convert DataFrame to 2D list
        data = [list(df.columns)] + df.astype(str).values.tolist()
        rows = len(data)
        columns = len(data[0])

        # Determine insertion index
        if index is None:
            document = self.get_document(document_id)
            index = document['body']['content'][-1]['endIndex'] - 1

        # Step 1: Insert table structure
        table_response = self.insert_table(document_id, rows, columns, index)

        time.sleep(1)  # Give API a moment to reflect the new table
        document = self.get_document(document_id)

        # Step 2: Locate the newly inserted table
        table_element = None
        table_start_index = None
        for element in reversed(document['body']['content']):
            if 'table' in element:
                table_element = element['table']
                table_start_index = element['startIndex']
                break

        if table_element is None:
            self.append_text(document_id, "\n\nTable Data:\n" + df.to_string(index=False))
            return table_response

        # Step 3: Build insert and style requests with offset tracking
        requests = []
        text_style_requests = []
        table_cell_style_requests = []

        table_rows = table_element.get('tableRows', [])

        inserted_texts = []

        for row_idx, row_data in enumerate(data):
            if row_idx >= len(table_rows):
                break

            table_row = table_rows[row_idx]
            table_cells = table_row.get('tableCells', [])

            for col_idx, cell_data in enumerate(row_data):
                if col_idx >= len(table_cells):
                    break

                text = str(cell_data).strip()
                if not text:
                    continue

                table_cell = table_cells[col_idx]
                cell_content = table_cell.get('content', [])

                for content_element in cell_content:
                    if 'paragraph' in content_element:
                        para_start = content_element['startIndex']
                        inserted_texts.append({
                            "text": clean_text(text),
                            "start_index": para_start,
                            "is_header": row_idx == 0
                        })
                        break

        # Sort to handle text insertions in reverse to prevent index shifts
        inserted_texts.sort(key=lambda x: x["start_index"])
        offset = 0

        for entry in inserted_texts:
            text = entry["text"]
            start_index = entry["start_index"] + offset
            end_index = start_index + len(text)
            is_header = entry["is_header"]

            # Insert text
            requests.append({
                "insertText": {
                    "location": {"index": start_index},
                    "text": text
                }
            })

            # Apply text style
            style = header_style if is_header else cell_style
            if style:
                text_style_requests.append({
                    "updateTextStyle": {
                        "range": {"startIndex": start_index, "endIndex": end_index},
                        "textStyle": style.to_dict(),
                        "fields": "*"
                    }
                })

            # Apply table cell style
            if table_cell_style and table_start_index is not None:
                # Table cell styles need to be applied differently - use proper API structure
                table_cell_style_requests.append({
                    "updateTableCellStyle": {
                        "tableRange": {
                            "tableCellLocation": {
                                "tableStartLocation": {"index": table_start_index},
                                "rowIndex": row_idx,
                                "columnIndex": col_idx
                            },
                            "rowSpan": 1,
                            "columnSpan": 1
                        },
                        "tableCellStyle": table_cell_style.to_dict(),
                        "fields": "*"
                    }
                })

            offset += len(text)

        # Combine all requests
        all_requests = requests + text_style_requests + table_cell_style_requests

        # Send the batch update
        try:
            if all_requests:
                self.update_document(document_id, all_requests)
            else:
                self.append_text(document_id, "\n\nTable Data:\n" + df.to_string(index=False))
        except Exception as e:
            print(f"Warning: Could not populate table cells directly: {e}")
            self.append_text(document_id, "\n\nTable Data:\n" + df.to_string(index=False))

        return table_response

    def replace_text_with_table_from_dataframe(self, document_id: str, placeholder: str, df: pd.DataFrame,
                                          header_style: Optional[TextStyle] = None,
                                          cell_style: Optional[TextStyle] = None,
                                          table_cell_style: Optional[TableCellStyle] = None) -> Dict[str, Any]:
        """
        Find and replace a placeholder text with a styled table created from a pandas DataFrame.

        Args:
            document_id (str): Document ID.
            placeholder (str): Placeholder text to find and replace (e.g., "{{table_placeholder}}").
            df (pd.DataFrame): DataFrame to insert as a table.
            header_style (TextStyle, optional): Style for the header row.
            cell_style (TextStyle, optional): Style for the data rows.
            table_cell_style (TableCellStyle, optional): Style for the table cells.

        Returns:
            Dict[str, Any]: Response from the batch update request.
        """
        import time

        if df.empty:
            raise ValueError("DataFrame is empty")
        
        df=df.reset_index(drop=True)

        # Fetch document content to find placeholder index
        document = self.get_document(document_id)
        content = document.get('body', {}).get('content', [])

        placeholder_start = None
        placeholder_end = None

        for element in content:
            if 'paragraph' in element:
                for elem in element['paragraph'].get('elements', []):
                    text_run = elem.get('textRun')
                    if text_run and placeholder in text_run.get('content', ''):
                        placeholder_start = elem['startIndex']
                        placeholder_end = elem['endIndex']
                        break
                if placeholder_start is not None:
                    break

        if placeholder_start is None:
            print(f"Placeholder '{placeholder}' not found in the document.")
            return {}

        # Prepare table data (header + rows)
        data = [list(df.columns)] + df.astype(str).values.tolist()
        rows = len(data)
        columns = len(data[0])

        # Batch request: delete placeholder AND insert table at placeholder_start
        batch_requests = [
            {
                "deleteContentRange": {
                    "range": {
                        "startIndex": placeholder_start,
                        "endIndex": placeholder_end
                    }
                }
            },
            {
                "insertTable": {
                    "rows": rows,
                    "columns": columns,
                    "location": {
                        "index": placeholder_start
                    }
                }
            }
        ]

        # Execute deletion + table insertion
        self.update_document(document_id, batch_requests)

        # Wait for the document to update
        time.sleep(1)

        # Fetch updated document and locate the new table
        updated_doc = self.get_document(document_id)
        tables = [
            (elem['startIndex'], elem['table'])
            for elem in updated_doc.get('body', {}).get('content', [])
            if 'table' in elem
        ]

        # Find table with startIndex >= placeholder_start, closest to placeholder_start
        table_start_index = None
        table_element = None

        for start_idx, table in sorted(tables, key=lambda x: x[0]):
            if start_idx >= placeholder_start:
                table_start_index = start_idx
                table_element = table
                break

        # If not found (edge case), fallback to last table
        if table_element is None and tables:
            table_start_index, table_element = tables[-1]


        if table_element is None:
            # Fallback: insert plain text table data if table not found
            self.append_text(document_id, "\n\nTable Data:\n" + df.to_string(index=False))
            return {}

        # Prepare requests to insert text and apply styles inside table cells
        text_insert_requests = []
        text_style_requests = []
        table_cell_style_requests = []

        table_rows = table_element.get('tableRows', [])

        inserted_texts = []

        for row_idx, row_data in enumerate(data):
            if row_idx >= len(table_rows):
                break

            table_row = table_rows[row_idx]
            table_cells = table_row.get('tableCells', [])

            for col_idx, cell_data in enumerate(row_data):
                if col_idx >= len(table_cells):
                    break

                text = str(cell_data).strip()
                if not text:
                    continue

                table_cell = table_cells[col_idx]
                cell_content = table_cell.get('content', [])

                # Find paragraph start index inside cell
                para_start = None
                for content_element in cell_content:
                    if 'paragraph' in content_element:
                        para_start = content_element['startIndex']
                        break

                if para_start is None:
                    continue

                inserted_texts.append({
                    "text": clean_text(text),
                    "start_index": para_start,
                    "is_header": row_idx == 0,
                    "row_idx": row_idx,
                    "col_idx": col_idx
                })

        # Sort insertions by start index to avoid shifting problems
        inserted_texts.sort(key=lambda x: x["start_index"])

        offset = 0
        for entry in inserted_texts:
            text = entry["text"]
            start_index = entry["start_index"] + offset
            end_index = start_index + len(text)
            is_header = entry["is_header"]
            row_idx = entry["row_idx"]
            col_idx = entry["col_idx"]

            # Insert text
            text_insert_requests.append({
                "insertText": {
                    "location": {"index": start_index},
                    "text": text
                }
            })

            # Apply text style
            style = header_style if is_header else cell_style
            if style:
                text_style_requests.append({
                    "updateTextStyle": {
                        "range": {"startIndex": start_index, "endIndex": end_index},
                        "textStyle": style.to_dict(),
                        "fields": "*"
                    }
                })

            # Apply table cell style
            if table_cell_style and table_start_index is not None:
                table_cell_style_requests.append({
                    "updateTableCellStyle": {
                        "tableRange": {
                            "tableCellLocation": {
                                "tableStartLocation": {"index": table_start_index},
                                "rowIndex": row_idx,
                                "columnIndex": col_idx
                            },
                            "rowSpan": 1,
                            "columnSpan": 1
                        },
                        "tableCellStyle": table_cell_style.to_dict(),
                        "fields": "*"
                    }
                })

            offset += len(text)

        # Combine all requests and send batch update
        all_requests = text_insert_requests + text_style_requests + table_cell_style_requests

        try:
            if all_requests:
                return self.update_document(document_id, all_requests)
            else:
                self.append_text(document_id, "\n\nTable Data:\n" + df.to_string(index=False))
                return {}
        except Exception as e:
            print(f"Warning: Failed to populate table cells: {e}")
            self.append_text(document_id, "\n\nTable Data:\n" + df.to_string(index=False))
            return {}

    def create_bulleted_list(self, document_id: str, items: List[str],
                           index: Optional[int] = None,
                           style: Optional[TextStyle] = None) -> Dict[str, Any]:
        """
        Create a bulleted list in the document.
        
        Args:
            document_id (str): Document ID
            items (List[str]): List items
            index (int, optional): Position to insert at
            style (TextStyle, optional): Text style for items
            
        Returns:
            Dict[str, Any]: Update response
        """
        # Simplified approach: just append the list items as text with bullet points
        list_text = "\n"
        for item in items:
            list_text += f"• {clean_text(item)}\n"
        list_text += "\n"
        
        return self.append_text(document_id, list_text, style)
    
    def create_numbered_list(self, document_id: str, items: List[str],
                           index: Optional[int] = None,
                           style: Optional[TextStyle] = None) -> Dict[str, Any]:
        """
        Create a numbered list in the document.
        
        Args:
            document_id (str): Document ID
            items (List[str]): List items
            index (int, optional): Position to insert at
            style (TextStyle, optional): Text style for items
            
        Returns:
            Dict[str, Any]: Update response
        """
        # Simplified approach: just append the list items as text with numbers
        list_text = "\n"
        for idx, item in enumerate(items, 1):
            list_text += f"{idx}. {clean_text(item)}\n"
        list_text += "\n"
        
        return self.append_text(document_id, list_text, style)
    
    def insert_page_break(self, document_id: str, index: Optional[int] = None) -> Dict[str, Any]:
        """Insert a page break."""
        if index is None:
            document = self.get_document(document_id)
            index = document['body']['content'][-1]['endIndex'] - 1
        
        request = RequestBuilder.insert_page_break(index)
        return self.update_document(document_id, [request])
    
    def insert_section_break(self, document_id: str, section_type: str = "NEXT_PAGE",
                           index: Optional[int] = None) -> Dict[str, Any]:
        """Insert a section break."""
        if index is None:
            document = self.get_document(document_id)
            index = document['body']['content'][-1]['endIndex'] - 1
        
        request = RequestBuilder.insert_section_break(index, section_type)
        return self.update_document(document_id, [request])
    
    def insert_image(self, document_id: str, image_uri: str,
                    index: Optional[int] = None,
                    width: Optional[Dimension] = None,
                    height: Optional[Dimension] = None) -> Dict[str, Any]:
        """
        Insert an image into the document.
        
        Args:
            document_id (str): Document ID
            image_uri (str): URI of the image
            index (int, optional): Position to insert at
            width (Dimension, optional): Image width
            height (Dimension, optional): Image height
            
        Returns:
            Dict[str, Any]: Update response
        """
        if index is None:
            document = self.get_document(document_id)
            index = document['body']['content'][-1]['endIndex'] - 1
        
        request = RequestBuilder.insert_inline_image(index, image_uri, width, height)
        return self.update_document(document_id, [request])
    
    def delete_content(self, document_id: str, start_index: int, end_index: int) -> Dict[str, Any]:
        """Delete content from the document."""
        validate_index_range(start_index, end_index)
        request = RequestBuilder.delete_content(start_index, end_index)
        return self.update_document(document_id, [request])
    
    def find_and_replace(self, document_id: str, find_text: str, replace_text: str,
                        match_case: bool = False) -> Dict[str, Any]:
        """Find and replace text in the document."""
        return self.replace_text(document_id, find_text, replace_text, match_case)
    
    def set_document_title(self, document_id: str, title: str) -> Dict[str, Any]:
        """Set the document title."""
        # Update document title through Drive API
        try:
            body = {'name': title}
            return self.drive_service.files().update(
                fileId=document_id, body=body
            ).execute()
        except HttpError as e:
            raise Exception(f"Failed to set document title: {e}")
    
    def share_document(self, document_id: str, email: str, 
                      role: str = "reader") -> Dict[str, Any]:
        """
        Share a document with a user.
        
        Args:
            document_id (str): Document ID
            email (str): Email address to share with
            role (str): Permission role (reader, writer, commenter)
            
        Returns:
            Dict[str, Any]: Share response
        """
        try:
            permission = {
                'type': 'user',
                'role': role,
                'emailAddress': email
            }
            
            return self.drive_service.permissions().create(
                fileId=document_id,
                body=permission,
                sendNotificationEmail=True
            ).execute()
        except HttpError as e:
            raise Exception(f"Failed to share document: {e}")
    
    def export_document(self, document_id: str, mime_type: str = 'application/pdf') -> bytes:
        """
        Export document in specified format.
        
        Args:
            document_id (str): Document ID
            mime_type (str): Export format MIME type
            
        Returns:
            bytes: Exported document content
        """
        try:
            return self.drive_service.files().export(
                fileId=document_id, mimeType=mime_type
            ).execute()
        except HttpError as e:
            raise Exception(f"Failed to export document: {e}")
    
    def build_document_from_elements(self, document_id: str, 
                                   elements: List[DocumentElement]) -> Dict[str, Any]:
        """
        Build a document from a list of elements using batch requests.
        
        Args:
            document_id (str): Document ID
            elements (List[DocumentElement]): Document elements to add
            
        Returns:
            Dict[str, Any]: Update response
        """
        batch_builder = BatchRequestBuilder()
        
        for element in elements:
            if isinstance(element, Paragraph):
                text_content = ""
                for elem in element.elements:
                    if isinstance(elem, TextRun):
                        text_content += elem.content
                
                batch_builder.insert_paragraph(
                    text_content,
                    element.elements[0].style if element.elements else None,
                    element.style
                )
            elif isinstance(element, Table):
                batch_builder.insert_table(len(element.rows), element.columns)
            elif isinstance(element, PageBreak):
                batch_builder.insert_page_break()
        
        requests = batch_builder.get_requests()
        return self.update_document(document_id, requests)
    
    def get_current_document_id(self) -> Optional[str]:
        """Get the current document ID."""
        return self._current_document_id
    
    def set_current_document_id(self, document_id: str) -> None:
        """Set the current document ID for operations."""
        self._current_document_id = document_id
    
    def retry_on_error(self, func, *args, max_retries: int = 3, **kwargs):
        """
        Retry a function call on certain errors.
        
        Args:
            func: Function to call
            *args: Function arguments
            max_retries (int): Maximum number of retries
            **kwargs: Function keyword arguments
            
        Returns:
            Function result
            
        Raises:
            Exception: If all retries fail
        """
        last_error = None
        
        for attempt in range(max_retries + 1):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                last_error = e
                
                if attempt == max_retries:
                    break
                
                if ErrorHandler.is_quota_exceeded_error(e):
                    delay = ErrorHandler.get_retry_delay(attempt + 1)
                    time.sleep(delay)
                elif ErrorHandler.is_permission_error(e) or ErrorHandler.is_not_found_error(e):
                    # Don't retry permission or not found errors
                    break
                else:
                    # Brief delay for other errors
                    time.sleep(0.5)
        
        raise last_error
