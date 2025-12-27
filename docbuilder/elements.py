"""
Google Docs API Elements Module

This module contains classes representing various Google Docs document elements
such as paragraphs, text runs, tables, lists, and other document components.
"""

from __future__ import annotations
from typing import Dict, Any, List, Optional, Union
from dataclasses import dataclass, field
from abc import ABC, abstractmethod

from .styles import (
    TextStyle, ParagraphStyle, TableCellStyle, Color, Dimension,
    ParagraphStyleType, TextAlignment, ListType, Colors
)


class DocumentElement(ABC):
    """Abstract base class for all document elements."""
    
    @abstractmethod
    def to_request(self) -> Dict[str, Any]:
        """Convert the element to a Google Docs API request format."""
        pass


@dataclass
class TextRun(DocumentElement):
    """
    Represents a run of text with consistent formatting.
    
    A text run is a sequence of characters that share the same text style.
    """
    content: str
    style: Optional[TextStyle] = None
    
    def to_request(self) -> Dict[str, Any]:
        """Convert to Google Docs API insertText request format."""
        request = {
            "insertText": {
                "text": self.content
            }
        }
        
        if self.style:
            request["insertText"]["textStyle"] = self.style.to_dict()
        
        return request
    
    def to_content_dict(self) -> Dict[str, Any]:
        """Convert to document content format."""
        content = {
            "textRun": {
                "content": self.content
            }
        }
        
        if self.style:
            content["textRun"]["textStyle"] = self.style.to_dict()
        
        return content


@dataclass
class Paragraph(DocumentElement):
    """
    Represents a paragraph containing text runs and other inline elements.
    """
    elements: List[Union[TextRun, 'InlineObject']] = field(default_factory=list)
    style: Optional[ParagraphStyle] = None
    bullet: Optional[Dict[str, Any]] = None
    
    def add_text(self, text: str, style: Optional[TextStyle] = None) -> 'Paragraph':
        """
        Add a text run to the paragraph.
        
        Args:
            text (str): Text content
            style (TextStyle, optional): Text formatting
            
        Returns:
            Paragraph: Self for method chaining
        """
        self.elements.append(TextRun(content=text, style=style))
        return self
    
    def add_text_run(self, text_run: TextRun) -> 'Paragraph':
        """
        Add a text run object to the paragraph.
        
        Args:
            text_run (TextRun): Text run to add
            
        Returns:
            Paragraph: Self for method chaining
        """
        self.elements.append(text_run)
        return self
    
    def to_request(self) -> List[Dict[str, Any]]:
        """Convert to a list of Google Docs API requests."""
        requests = []
        
        # Insert paragraph break first
        requests.append({
            "insertText": {
                "text": "\n"
            }
        })
        
        # Apply paragraph style if specified
        if self.style or self.bullet:
            style_request = {"updateParagraphStyle": {"paragraphStyle": {}}}
            
            if self.style:
                style_request["updateParagraphStyle"]["paragraphStyle"].update(
                    self.style.to_dict()
                )
            
            if self.bullet:
                style_request["updateParagraphStyle"]["paragraphStyle"]["bullet"] = self.bullet
            
            requests.append(style_request)
        
        # Add text runs
        for element in self.elements:
            if isinstance(element, TextRun):
                requests.append(element.to_request())
        
        return requests
    
    def to_content_dict(self) -> Dict[str, Any]:
        """Convert to document content format."""
        content = {
            "paragraph": {
                "elements": [elem.to_content_dict() for elem in self.elements]
            }
        }
        
        if self.style:
            content["paragraph"]["paragraphStyle"] = self.style.to_dict()
        
        if self.bullet:
            content["paragraph"]["bullet"] = self.bullet
        
        return content


@dataclass
class TableCell:
    """Represents a single cell in a table."""
    content: List[Paragraph] = field(default_factory=list)
    style: Optional[TableCellStyle] = None
    
    def add_paragraph(self, paragraph: Paragraph) -> 'TableCell':
        """Add a paragraph to the cell."""
        self.content.append(paragraph)
        return self
    
    def add_text(self, text: str, text_style: Optional[TextStyle] = None) -> 'TableCell':
        """Add text as a new paragraph in the cell."""
        paragraph = Paragraph().add_text(text, text_style)
        self.content.append(paragraph)
        return self
    
    def to_content_dict(self) -> Dict[str, Any]:
        """Convert to document content format."""
        cell_content = {
            "content": [p.to_content_dict() for p in self.content]
        }
        
        if self.style:
            cell_content["tableCellStyle"] = self.style.to_dict()
        
        return cell_content


@dataclass
class TableRow:
    """Represents a row in a table."""
    cells: List[TableCell] = field(default_factory=list)
    
    def add_cell(self, cell: TableCell) -> 'TableRow':
        """Add a cell to the row."""
        self.cells.append(cell)
        return self
    
    def add_text_cell(self, text: str, text_style: Optional[TextStyle] = None) -> 'TableRow':
        """Add a cell with text content."""
        cell = TableCell().add_text(text, text_style)
        self.cells.append(cell)
        return self
    
    def to_content_dict(self) -> Dict[str, Any]:
        """Convert to document content format."""
        return {
            "tableCells": [cell.to_content_dict() for cell in self.cells]
        }


@dataclass
class Table(DocumentElement):
    """Represents a table with rows and columns."""
    rows: List[TableRow] = field(default_factory=list)
    columns: int = 0
    
    def __post_init__(self):
        """Calculate number of columns based on first row."""
        if self.rows and self.columns == 0:
            self.columns = len(self.rows[0].cells)
    
    def add_row(self, row: TableRow) -> 'Table':
        """Add a row to the table."""
        self.rows.append(row)
        if self.columns == 0:
            self.columns = len(row.cells)
        return self
    
    def create_row(self, cell_texts: List[str]) -> 'Table':
        """Create and add a row with text cells."""
        row = TableRow()
        for text in cell_texts:
            row.add_text_cell(text)
        self.add_row(row)
        return self
    
    def to_request(self) -> Dict[str, Any]:
        """Convert to Google Docs API insertTable request."""
        return {
            "insertTable": {
                "rows": len(self.rows),
                "columns": self.columns
            }
        }
    
    def to_content_dict(self) -> Dict[str, Any]:
        """Convert to document content format."""
        return {
            "table": {
                "columns": self.columns,
                "rows": len(self.rows),
                "tableRows": [row.to_content_dict() for row in self.rows]
            }
        }


@dataclass
class List(DocumentElement):
    """Represents a bulleted or numbered list."""
    items: List[Paragraph] = field(default_factory=list)
    list_type: ListType = ListType.BULLETED
    list_id: Optional[str] = None
    
    def add_item(self, text: str, style: Optional[TextStyle] = None, level: int = 0) -> 'List':
        """
        Add an item to the list.
        
        Args:
            text (str): Item text
            style (TextStyle, optional): Text formatting
            level (int): Nesting level (0-based)
            
        Returns:
            List: Self for method chaining
        """
        bullet_config = {
            "listId": self.list_id or f"{self.list_type.value.lower()}_list",
            "nestingLevel": level
        }
        
        paragraph = Paragraph(
            elements=[TextRun(content=text, style=style)],
            bullet=bullet_config
        )
        
        self.items.append(paragraph)
        return self
    
    def add_paragraph_item(self, paragraph: Paragraph, level: int = 0) -> 'List':
        """Add a paragraph as a list item."""
        bullet_config = {
            "listId": self.list_id or f"{self.list_type.value.lower()}_list",
            "nestingLevel": level
        }
        paragraph.bullet = bullet_config
        self.items.append(paragraph)
        return self
    
    def to_request(self) -> List[Dict[str, Any]]:
        """Convert to list of Google Docs API requests."""
        requests = []
        
        # Create list if needed
        if self.list_id is None:
            self.list_id = f"{self.list_type.value.lower()}_list"
        
        # Add each list item
        for item in self.items:
            requests.extend(item.to_request())
        
        return requests


@dataclass
class InlineObject(DocumentElement):
    """Represents an inline object like an image."""
    object_id: str
    inline_object_properties: Dict[str, Any] = field(default_factory=dict)
    
    def to_request(self) -> Dict[str, Any]:
        """Convert to Google Docs API insertInlineImage request."""
        return {
            "insertInlineImage": {
                "objectId": self.object_id,
                "inlineObjectElement": {
                    "inlineObjectProperties": self.inline_object_properties
                }
            }
        }
    
    def to_content_dict(self) -> Dict[str, Any]:
        """Convert to document content format."""
        return {
            "inlineObjectElement": {
                "objectId": self.object_id,
                "inlineObjectProperties": self.inline_object_properties
            }
        }


@dataclass
class Bookmark(DocumentElement):
    """Represents a bookmark in the document."""
    bookmark_id: str
    name: str
    
    def to_request(self) -> Dict[str, Any]:
        """Convert to Google Docs API insertBookmark request."""
        return {
            "insertBookmark": {
                "bookmarkId": self.bookmark_id,
                "name": self.name
            }
        }


@dataclass
class PageBreak(DocumentElement):
    """Represents a page break."""
    
    def to_request(self) -> Dict[str, Any]:
        """Convert to Google Docs API insertPageBreak request."""
        return {
            "insertPageBreak": {}
        }
    
    def to_content_dict(self) -> Dict[str, Any]:
        """Convert to document content format."""
        return {
            "pageBreak": {}
        }


@dataclass
class SectionBreak(DocumentElement):
    """Represents a section break."""
    section_type: str = "NEXT_PAGE"  # NEXT_PAGE, CONTINUOUS, EVEN_PAGE, ODD_PAGE
    
    def to_request(self) -> Dict[str, Any]:
        """Convert to Google Docs API insertSectionBreak request."""
        return {
            "insertSectionBreak": {
                "sectionType": self.section_type
            }
        }


class DocumentBuilder:
    """
    Helper class for building document content with a fluent interface.
    """
    
    def __init__(self):
        self.elements: List[DocumentElement] = []
    
    def add_title(self, text: str, style: Optional[TextStyle] = None) -> 'DocumentBuilder':
        """Add a title paragraph."""
        title_style = ParagraphStyle(namedStyleType=ParagraphStyleType.TITLE)
        paragraph = Paragraph(style=title_style).add_text(text, style)
        self.elements.append(paragraph)
        return self
    
    def add_subtitle(self, text: str, style: Optional[TextStyle] = None) -> 'DocumentBuilder':
        """Add a subtitle paragraph."""
        subtitle_style = ParagraphStyle(namedStyleType=ParagraphStyleType.SUBTITLE)
        paragraph = Paragraph(style=subtitle_style).add_text(text, style)
        self.elements.append(paragraph)
        return self
    
    def add_heading(self, text: str, level: int = 1, style: Optional[TextStyle] = None) -> 'DocumentBuilder':
        """Add a heading paragraph."""
        heading_types = {
            1: ParagraphStyleType.HEADING_1,
            2: ParagraphStyleType.HEADING_2,
            3: ParagraphStyleType.HEADING_3,
            4: ParagraphStyleType.HEADING_4,
            5: ParagraphStyleType.HEADING_5,
            6: ParagraphStyleType.HEADING_6
        }
        
        if level not in heading_types:
            raise ValueError("Heading level must be between 1 and 6")
        
        heading_style = ParagraphStyle(namedStyleType=heading_types[level])
        paragraph = Paragraph(style=heading_style).add_text(text, style)
        self.elements.append(paragraph)
        return self
    
    def add_paragraph(self, text: str = "", text_style: Optional[TextStyle] = None, 
                     paragraph_style: Optional[ParagraphStyle] = None) -> 'DocumentBuilder':
        """Add a regular paragraph."""
        paragraph = Paragraph(style=paragraph_style)
        if text:
            paragraph.add_text(text, text_style)
        self.elements.append(paragraph)
        return self
    
    def add_custom_paragraph(self, paragraph: Paragraph) -> 'DocumentBuilder':
        """Add a custom paragraph object."""
        self.elements.append(paragraph)
        return self
    
    def add_table(self, table: Table) -> 'DocumentBuilder':
        """Add a table."""
        self.elements.append(table)
        return self
    
    def add_list(self, list_obj: List) -> 'DocumentBuilder':
        """Add a list."""
        self.elements.append(list_obj)
        return self
    
    def add_page_break(self) -> 'DocumentBuilder':
        """Add a page break."""
        self.elements.append(PageBreak())
        return self
    
    def add_section_break(self, section_type: str = "NEXT_PAGE") -> 'DocumentBuilder':
        """Add a section break."""
        self.elements.append(SectionBreak(section_type=section_type))
        return self
    
    def build_requests(self) -> List[Dict[str, Any]]:
        """Build a list of Google Docs API requests from all elements."""
        requests = []
        
        for element in self.elements:
            if hasattr(element, 'to_request'):
                element_requests = element.to_request()
                if isinstance(element_requests, list):
                    requests.extend(element_requests)
                else:
                    requests.append(element_requests)
        
        return requests
    
    def build_content(self) -> List[Dict[str, Any]]:
        """Build document content structure."""
        return [element.to_content_dict() for element in self.elements 
                if hasattr(element, 'to_content_dict')]


def create_simple_table(data: List[List[str]], 
                       header_style: Optional[TextStyle] = None,
                       cell_style: Optional[TextStyle] = None) -> Table:
    """
    Create a simple table from a 2D list of strings.
    
    Args:
        data (List[List[str]]): 2D list where each inner list is a row
        header_style (TextStyle, optional): Style for header row
        cell_style (TextStyle, optional): Style for data cells
        
    Returns:
        Table: Configured table object
    """
    if not data:
        raise ValueError("Data cannot be empty")
    
    table = Table()
    
    for i, row_data in enumerate(data):
        row = TableRow()
        style = header_style if i == 0 and header_style else cell_style
        
        for cell_text in row_data:
            cell = TableCell().add_text(cell_text, style)
            row.add_cell(cell)
        
        table.add_row(row)
    
    return table


def create_bulleted_list(items: List[str], 
                        style: Optional[TextStyle] = None) -> List:
    """
    Create a bulleted list from a list of strings.
    
    Args:
        items (List[str]): List item texts
        style (TextStyle, optional): Text style for items
        
    Returns:
        List: Configured list object
    """
    bullet_list = List(list_type=ListType.BULLETED)
    
    for item in items:
        bullet_list.add_item(item, style)
    
    return bullet_list


def create_numbered_list(items: List[str], 
                        style: Optional[TextStyle] = None) -> List:
    """
    Create a numbered list from a list of strings.
    
    Args:
        items (List[str]): List item texts
        style (TextStyle, optional): Text style for items
        
    Returns:
        List: Configured list object
    """
    numbered_list = List(list_type=ListType.NUMBERED)
    
    for item in items:
        numbered_list.add_item(item, style)
    
    return numbered_list
