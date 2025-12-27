"""
Google Docs API Utilities Module

This module contains helper functions and utilities for working with the Google Docs API,
including request builders, response parsers, and common operations.
"""

from __future__ import annotations
import re
import uuid
from typing import Dict, Any, List, Optional, Union, Tuple
from datetime import datetime

from .styles import Color, Dimension, TextStyle, ParagraphStyle
from .elements import DocumentElement, Paragraph, TextRun, Table


class RequestBuilder:
    """Helper class for building Google Docs API requests."""
    
    @staticmethod
    def insert_text(text: str, location_index: int, text_style: Optional[TextStyle] = None) -> Dict[str, Any]:
        """
        Build an insertText request.
        
        Args:
            text (str): Text to insert
            location_index (int): Index where to insert text
            text_style (TextStyle, optional): Text formatting (Note: style must be applied separately)
            
        Returns:
            Dict[str, Any]: InsertText request
        """
        request = {
            "insertText": {
                "location": {"index": location_index},
                "text": text
            }
        }
        
        # Note: Text style cannot be applied directly in insertText request
        # It must be applied separately using updateTextStyle
        
        return request
    
    @staticmethod
    def update_text_style(start_index: int, end_index: int, text_style: TextStyle) -> Dict[str, Any]:
        """
        Build an updateTextStyle request.
        
        Args:
            start_index (int): Start of text range
            end_index (int): End of text range
            text_style (TextStyle): Text formatting to apply
            
        Returns:
            Dict[str, Any]: UpdateTextStyle request
        """
        return {
            "updateTextStyle": {
                "range": {
                    "startIndex": start_index,
                    "endIndex": end_index
                },
                "textStyle": text_style.to_dict(),
                "fields": "*"
            }
        }
    
    @staticmethod
    def update_paragraph_style(start_index: int, end_index: int, paragraph_style: ParagraphStyle) -> Dict[str, Any]:
        """
        Build an updateParagraphStyle request.
        
        Args:
            start_index (int): Start of paragraph range
            end_index (int): End of paragraph range
            paragraph_style (ParagraphStyle): Paragraph formatting to apply
            
        Returns:
            Dict[str, Any]: UpdateParagraphStyle request
        """
        return {
            "updateParagraphStyle": {
                "range": {
                    "startIndex": start_index,
                    "endIndex": end_index
                },
                "paragraphStyle": paragraph_style.to_dict(),
                "fields": "*"
            }
        }
    
    @staticmethod
    def insert_table(location_index: int, rows: int, columns: int) -> Dict[str, Any]:
        """
        Build an insertTable request.
        
        Args:
            location_index (int): Where to insert the table
            rows (int): Number of rows
            columns (int): Number of columns
            
        Returns:
            Dict[str, Any]: InsertTable request
        """
        return {
            "insertTable": {
                "location": {"index": location_index},
                "rows": rows,
                "columns": columns
            }
        }
    
    @staticmethod
    def insert_page_break(location_index: int) -> Dict[str, Any]:
        """
        Build an insertPageBreak request.
        
        Args:
            location_index (int): Where to insert the page break
            
        Returns:
            Dict[str, Any]: InsertPageBreak request
        """
        return {
            "insertPageBreak": {
                "location": {"index": location_index}
            }
        }
    
    @staticmethod
    def insert_section_break(location_index: int, section_type: str = "NEXT_PAGE") -> Dict[str, Any]:
        """
        Build an insertSectionBreak request.
        
        Args:
            location_index (int): Where to insert the section break
            section_type (str): Type of section break
            
        Returns:
            Dict[str, Any]: InsertSectionBreak request
        """
        return {
            "insertSectionBreak": {
                "location": {"index": location_index},
                "sectionType": section_type
            }
        }
    
    @staticmethod
    def delete_content(start_index: int, end_index: int) -> Dict[str, Any]:
        """
        Build a deleteContentRange request.
        
        Args:
            start_index (int): Start of range to delete
            end_index (int): End of range to delete
            
        Returns:
            Dict[str, Any]: DeleteContentRange request
        """
        return {
            "deleteContentRange": {
                "range": {
                    "startIndex": start_index,
                    "endIndex": end_index
                }
            }
        }
    
    @staticmethod
    def create_named_range(name: str, start_index: int, end_index: int) -> Dict[str, Any]:
        """
        Build a createNamedRange request.
        
        Args:
            name (str): Name for the range
            start_index (int): Start of range
            end_index (int): End of range
            
        Returns:
            Dict[str, Any]: CreateNamedRange request
        """
        return {
            "createNamedRange": {
                "name": name,
                "range": {
                    "startIndex": start_index,
                    "endIndex": end_index
                }
            }
        }
    
    @staticmethod
    def insert_inline_image(location_index: int, image_uri: str, 
                           width: Optional[Dimension] = None, 
                           height: Optional[Dimension] = None) -> Dict[str, Any]:
        """
        Build an insertInlineImage request.
        
        Args:
            location_index (int): Where to insert the image
            image_uri (str): URI of the image
            width (Dimension, optional): Image width
            height (Dimension, optional): Image height
            
        Returns:
            Dict[str, Any]: InsertInlineImage request
        """
        request = {
            "insertInlineImage": {
                "location": {"index": location_index},
                "uri": image_uri
            }
        }
        
        if width or height:
            object_size = {}
            if width:
                object_size["width"] = width.to_dict()
            if height:
                object_size["height"] = height.to_dict()
            
            request["insertInlineImage"]["objectSize"] = object_size
        
        return request


class DocumentParser:
    """Helper class for parsing Google Docs API responses."""
    
    @staticmethod
    def extract_text_content(document: Dict[str, Any]) -> str:
        """
        Extract all text content from a document.
        
        Args:
            document (Dict[str, Any]): Document object from API
            
        Returns:
            str: Extracted text content
        """
        content = document.get("body", {}).get("content", [])
        text_parts = []
        
        for element in content:
            if "paragraph" in element:
                paragraph = element["paragraph"]
                for para_element in paragraph.get("elements", []):
                    if "textRun" in para_element:
                        text_parts.append(para_element["textRun"]["content"])
        
        return "".join(text_parts)
    
    @staticmethod
    def extract_paragraphs(document: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Extract all paragraphs from a document.
        
        Args:
            document (Dict[str, Any]): Document object from API
            
        Returns:
            List[Dict[str, Any]]: List of paragraph objects
        """
        content = document.get("body", {}).get("content", [])
        paragraphs = []
        
        for element in content:
            if "paragraph" in element:
                paragraphs.append(element["paragraph"])
        
        return paragraphs
    
    @staticmethod
    def extract_tables(document: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Extract all tables from a document.
        
        Args:
            document (Dict[str, Any]): Document object from API
            
        Returns:
            List[Dict[str, Any]]: List of table objects
        """
        content = document.get("body", {}).get("content", [])
        tables = []
        
        for element in content:
            if "table" in element:
                tables.append(element["table"])
        
        return tables
    
    @staticmethod
    def find_text_ranges(document: Dict[str, Any], search_text: str) -> List[Tuple[int, int]]:
        """
        Find all occurrences of text in the document and return their ranges.
        
        Args:
            document (Dict[str, Any]): Document object from API
            search_text (str): Text to search for
            
        Returns:
            List[Tuple[int, int]]: List of (start_index, end_index) tuples
        """
        content = document.get("body", {}).get("content", [])
        ranges = []
        current_index = 1  # Documents start at index 1
        
        for element in content:
            if "paragraph" in element:
                paragraph = element["paragraph"]
                for para_element in paragraph.get("elements", []):
                    if "textRun" in para_element:
                        text_content = para_element["textRun"]["content"]
                        start_pos = 0
                        
                        while True:
                            pos = text_content.find(search_text, start_pos)
                            if pos == -1:
                                break
                            
                            start_index = current_index + pos
                            end_index = start_index + len(search_text)
                            ranges.append((start_index, end_index))
                            start_pos = pos + 1
                        
                        current_index += len(text_content)
        
        return ranges


class BatchRequestBuilder:
    """Helper class for building batched requests efficiently."""
    
    def __init__(self):
        self.requests: List[Dict[str, Any]] = []
        self._current_index = 1  # Documents start at index 1
    
    def add_request(self, request: Dict[str, Any]) -> 'BatchRequestBuilder':
        """Add a request to the batch."""
        self.requests.append(request)
        return self
    
    def insert_text(self, text: str, text_style: Optional[TextStyle] = None) -> 'BatchRequestBuilder':
        """Add an insertText request at the current position."""
        request = RequestBuilder.insert_text(text, self._current_index, text_style)
        self.requests.append(request)
        self._current_index += len(text)
        return self
    
    def insert_paragraph(self, text: str = "", text_style: Optional[TextStyle] = None,
                        paragraph_style: Optional[ParagraphStyle] = None) -> 'BatchRequestBuilder':
        """Add a paragraph with optional text and styling."""
        # Insert text first if provided
        if text:
            self.insert_text(text, text_style)
        
        # Insert paragraph break
        self.insert_text("\n")
        
        # Apply paragraph style if provided
        if paragraph_style:
            end_index = self._current_index
            start_index = end_index - len(text) - 1
            style_request = RequestBuilder.update_paragraph_style(
                start_index, end_index, paragraph_style
            )
            self.requests.append(style_request)
        
        return self
    
    def insert_table(self, rows: int, columns: int) -> 'BatchRequestBuilder':
        """Add a table insertion request."""
        request = RequestBuilder.insert_table(self._current_index, rows, columns)
        self.requests.append(request)
        # Tables take up space - estimate based on size
        self._current_index += (rows * columns * 2) + 1
        return self
    
    def insert_page_break(self) -> 'BatchRequestBuilder':
        """Add a page break."""
        request = RequestBuilder.insert_page_break(self._current_index)
        self.requests.append(request)
        self._current_index += 1
        return self
    
    def get_requests(self) -> List[Dict[str, Any]]:
        """Get all accumulated requests."""
        return self.requests.copy()
    
    def clear(self) -> 'BatchRequestBuilder':
        """Clear all requests and reset index."""
        self.requests.clear()
        self._current_index = 1
        return self


def generate_unique_id() -> str:
    """Generate a unique identifier for elements."""
    return str(uuid.uuid4())


def validate_index_range(start_index: int, end_index: int) -> None:
    """
    Validate that index range is valid.
    
    Args:
        start_index (int): Start index
        end_index (int): End index
        
    Raises:
        ValueError: If range is invalid
    """
    if start_index < 1:
        raise ValueError("Start index must be >= 1")
    if end_index <= start_index:
        raise ValueError("End index must be > start index")


def calculate_text_length(elements: List[DocumentElement]) -> int:
    """
    Calculate the total text length of document elements.
    
    Args:
        elements (List[DocumentElement]): List of document elements
        
    Returns:
        int: Total text length
    """
    total_length = 0
    
    for element in elements:
        if isinstance(element, TextRun):
            total_length += len(element.content)
        elif isinstance(element, Paragraph):
            for para_element in element.elements:
                if isinstance(para_element, TextRun):
                    total_length += len(para_element.content)
            total_length += 1  # For paragraph break
    
    return total_length


def clean_text(text: str) -> str:
    """
    Clean text for Google Docs insertion.
    
    Args:
        text (str): Text to clean
        
    Returns:
        str: Cleaned text
    """
    # Replace problematic characters
    text = text.replace('\r\n', '\n')  # Normalize line endings
    text = text.replace('\r', '\n')    # Mac line endings
    
    # Remove null characters and other problematic characters
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
    
    return text


def format_timestamp(timestamp: Optional[str] = None) -> str:
    """
    Format a timestamp for document metadata.
    
    Args:
        timestamp (str, optional): ISO timestamp string. If None, uses current time.
        
    Returns:
        str: Formatted timestamp
    """
    if timestamp is None:
        return datetime.now().isoformat()
    return timestamp


def merge_text_styles(*styles: Optional[TextStyle]) -> TextStyle:
    """
    Merge multiple text styles, with later styles overriding earlier ones.
    
    Args:
        *styles: Variable number of TextStyle objects
        
    Returns:
        TextStyle: Merged style
    """
    merged = TextStyle()
    
    for style in styles:
        if style is None:
            continue
        
        for field in style.__dataclass_fields__:
            value = getattr(style, field)
            if value is not None:
                setattr(merged, field, value)
    
    return merged


def create_hyperlink_style(url: str, color: Optional[Color] = None) -> TextStyle:
    """
    Create a text style for hyperlinks.
    
    Args:
        url (str): URL for the hyperlink
        color (Color, optional): Link color. Defaults to blue.
        
    Returns:
        TextStyle: Hyperlink text style
    """
    if color is None:
        color = Color(0.11, 0.46, 0.98)  # Google Docs default link blue
    
    return TextStyle(
        link_url=url,
        foreground_color=color,
        underline=True
    )


def estimate_document_size(elements: List[DocumentElement]) -> Dict[str, int]:
    """
    Estimate document size metrics.
    
    Args:
        elements (List[DocumentElement]): Document elements
        
    Returns:
        Dict[str, int]: Size metrics (characters, words, paragraphs, etc.)
    """
    metrics = {
        "characters": 0,
        "words": 0,
        "paragraphs": 0,
        "tables": 0,
        "lists": 0
    }
    
    for element in elements:
        if isinstance(element, TextRun):
            text = element.content
            metrics["characters"] += len(text)
            metrics["words"] += len(text.split())
        elif isinstance(element, Paragraph):
            metrics["paragraphs"] += 1
            for para_element in element.elements:
                if isinstance(para_element, TextRun):
                    text = para_element.content
                    metrics["characters"] += len(text)
                    metrics["words"] += len(text.split())
        elif isinstance(element, Table):
            metrics["tables"] += 1
        elif hasattr(element, 'items'):  # List-like element
            metrics["lists"] += 1
    
    return metrics


class ErrorHandler:
    """Helper class for handling Google Docs API errors."""
    
    @staticmethod
    def is_quota_exceeded_error(error: Exception) -> bool:
        """Check if error is due to quota exceeded."""
        error_str = str(error).lower()
        return "quota" in error_str or "rate limit" in error_str
    
    @staticmethod
    def is_permission_error(error: Exception) -> bool:
        """Check if error is due to insufficient permissions."""
        error_str = str(error).lower()
        return "permission" in error_str or "forbidden" in error_str
    
    @staticmethod
    def is_not_found_error(error: Exception) -> bool:
        """Check if error is due to document not found."""
        error_str = str(error).lower()
        return "not found" in error_str or "404" in error_str
    
    @staticmethod
    def get_retry_delay(attempt: int, base_delay: float = 1.0) -> float:
        """
        Calculate exponential backoff delay for retries.
        
        Args:
            attempt (int): Attempt number (1-based)
            base_delay (float): Base delay in seconds
            
        Returns:
            float: Delay in seconds
        """
        return min(base_delay * (2 ** (attempt - 1)), 60.0)  # Cap at 60 seconds
