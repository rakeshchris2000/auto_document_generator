"""
Google Docs API Client Library

A comprehensive Python client library for the Google Docs API that provides
easy-to-use interfaces for creating, editing, and managing Google Docs documents
using service account authentication.

This library supports:
- Document creation and management
- Text formatting and styling
- Paragraph and heading styles
- Tables with full styling support
- Lists (bulleted and numbered)
- Images and inline objects
- Page and section breaks
- Document sharing and export
- Batch operations for efficiency

Example usage:
    from google_docs_client import GoogleDocsClient, TextStyle, ParagraphStyle
    from google_docs_client.styles import Colors, FontSizes
    
    # Initialize client
    client = GoogleDocsClient('path/to/service_account.json')
    
    # Create a new document
    doc_id = client.create_document("My Document")
    
    # Add content
    client.add_title(doc_id, "Document Title")
    client.add_heading(doc_id, "Chapter 1", level=1)
    client.append_text(doc_id, "This is some content.")
    
    # Create a table
    data = [["Header 1", "Header 2"], ["Cell 1", "Cell 2"]]
    client.insert_table_from_data(doc_id, data)
"""

__version__ = "1.0.0"
__author__ = "Google Docs Client Library"
__email__ = "support@example.com"

# Main client class
from .docs_client import GoogleDocsClient

# Authentication
from .auth import GoogleDocsAuth, create_authenticated_client

# Document elements
from .elements import (
    DocumentElement,
    TextRun,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    List as DocList,
    InlineObject,
    Bookmark,
    PageBreak,
    SectionBreak,
    DocumentBuilder,
    create_simple_table,
    create_bulleted_list,
    create_numbered_list
)

# Styles and formatting
from .styles import (
    # Style classes
    TextStyle,
    ParagraphStyle,
    TableCellStyle,
    Color,
    Dimension,
    
    # Enums
    ParagraphStyleType,
    TextAlignment,
    ListType,
    FontWeight,
    TextDecoration,
    VerticalAlignment,
    PageOrientation,
    BorderStyle,
    TableCellVerticalAlignment,
    
    # Predefined values
    Colors,
    FontFamilies,
    FontSizes,
    
    # Helper functions
    create_border,
    create_heading_style,
    create_list_style,
    merge_styles
)

# Utilities
from .utils import (
    RequestBuilder,
    DocumentParser,
    BatchRequestBuilder,
    ErrorHandler,
    generate_unique_id,
    clean_text,
    create_hyperlink_style,
    merge_text_styles
)

# Convenience imports for common operations
__all__ = [
    # Main client
    'GoogleDocsClient',
    
    # Authentication
    'GoogleDocsAuth',
    'create_authenticated_client',
    
    # Core elements
    'DocumentElement',
    'TextRun',
    'Paragraph',
    'Table',
    'TableCell',
    'TableRow',
    'DocList',
    'InlineObject',
    'Bookmark',
    'PageBreak',
    'SectionBreak',
    'DocumentBuilder',
    
    # Style classes
    'TextStyle',
    'ParagraphStyle',
    'TableCellStyle',
    'Color',
    'Dimension',
    
    # Enums
    'ParagraphStyleType',
    'TextAlignment',
    'ListType',
    'FontWeight',
    'TextDecoration',
    'VerticalAlignment',
    'PageOrientation',
    'BorderStyle',
    'TableCellVerticalAlignment',
    
    # Predefined values
    'Colors',
    'FontFamilies',
    'FontSizes',
    
    # Utility classes
    'RequestBuilder',
    'DocumentParser',
    'BatchRequestBuilder',
    'ErrorHandler',
    
    # Helper functions
    'create_simple_table',
    'create_bulleted_list',
    'create_numbered_list',
    'create_border',
    'create_heading_style',
    'create_list_style',
    'merge_styles',
    'generate_unique_id',
    'clean_text',
    'create_hyperlink_style',
    'merge_text_styles'
]


def get_version() -> str:
    """Get the library version."""
    return __version__


def get_supported_features() -> dict:
    """
    Get a dictionary of supported Google Docs features.
    
    Returns:
        dict: Dictionary of feature categories and their supported operations
    """
    return {
        "authentication": [
            "Service account authentication",
            "OAuth2 flow",
            "Credential refresh"
        ],
        "document_operations": [
            "Create documents",
            "Get document content",
            "Update documents",
            "Delete content",
            "Share documents",
            "Export documents"
        ],
        "text_formatting": [
            "Bold, italic, underline",
            "Font family and size",
            "Text color and highlighting",
            "Superscript and subscript",
            "Hyperlinks",
            "Small caps"
        ],
        "paragraph_formatting": [
            "Heading styles (1-6)",
            "Title and subtitle styles",
            "Text alignment",
            "Line spacing",
            "Indentation",
            "Borders and spacing"
        ],
        "lists": [
            "Bulleted lists",
            "Numbered lists",
            "Multi-level nesting",
            "Custom list styles"
        ],
        "tables": [
            "Table creation",
            "Cell content editing",
            "Cell styling and borders",
            "Background colors",
            "Cell padding",
            "Table from data arrays"
        ],
        "page_layout": [
            "Page breaks",
            "Section breaks",
            "Headers and footers",
            "Page margins",
            "Page orientation"
        ],
        "media": [
            "Inline images",
            "Image sizing",
            "Drawing objects"
        ],
        "advanced": [
            "Bookmarks",
            "Named ranges",
            "Footnotes and endnotes",
            "Document metadata",
            "Batch operations",
            "Error handling and retries"
        ]
    }


# Quick start helper function
def quick_start(service_account_file: str, document_title: str = "New Document") -> tuple:
    """
    Quick start function to create a client and document.
    
    Args:
        service_account_file (str): Path to service account JSON file
        document_title (str): Title for the new document
        
    Returns:
        tuple: (GoogleDocsClient instance, document_id)
        
    Example:
        client, doc_id = quick_start('service_account.json', 'My Document')
        client.add_title(doc_id, 'Welcome!')
    """
    client = GoogleDocsClient(service_account_file)
    doc_id = client.create_document(document_title)
    return client, doc_id
