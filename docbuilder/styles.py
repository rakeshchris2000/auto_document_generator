"""
Google Docs API Styles Module

This module contains constants, enums, and helper functions for working with
Google Docs styles, formatting, and layout options.
"""

from __future__ import annotations
from enum import Enum
from typing import Dict, Any, Optional, Union, List
from dataclasses import dataclass


class ParagraphStyleType(Enum):
    """Enumeration of paragraph style types."""
    NORMAL_TEXT = "NORMAL_TEXT"
    TITLE = "TITLE"
    SUBTITLE = "SUBTITLE"
    HEADING_1 = "HEADING_1"
    HEADING_2 = "HEADING_2"
    HEADING_3 = "HEADING_3"
    HEADING_4 = "HEADING_4"
    HEADING_5 = "HEADING_5"
    HEADING_6 = "HEADING_6"


class TextAlignment(Enum):
    """Text alignment options."""
    START = "START"
    CENTER = "CENTER"
    END = "END"
    JUSTIFY = "JUSTIFY"


class ListType(Enum):
    """List types for bulleted and numbered lists."""
    BULLETED = "BULLETED"
    NUMBERED = "NUMBERED"


class FontWeight(Enum):
    """Font weight options."""
    NORMAL = "NORMAL"
    BOLD = "BOLD"


class TextDecoration(Enum):
    """Text decoration options."""
    NONE = "NONE"
    UNDERLINE = "UNDERLINE"
    STRIKETHROUGH = "STRIKETHROUGH"


class VerticalAlignment(Enum):
    """Vertical alignment for text."""
    BASELINE = "BASELINE"
    SUBSCRIPT = "SUBSCRIPT"
    SUPERSCRIPT = "SUPERSCRIPT"


class PageOrientation(Enum):
    """Page orientation options."""
    PORTRAIT = "PORTRAIT"
    LANDSCAPE = "LANDSCAPE"


class BorderStyle(Enum):
    """Border style options for tables."""
    SOLID = "SOLID"
    DOTTED = "DOTTED"
    DASHED = "DASHED"


class TableCellVerticalAlignment(Enum):
    """Vertical alignment options for table cells."""
    TOP = "TOP"
    MIDDLE = "MIDDLE"
    BOTTOM = "BOTTOM"


@dataclass
class Color:
    """Represents a color in RGB format."""
    red: float = 0.0
    green: float = 0.0
    blue: float = 0.0
    
    @classmethod
    def from_hex(cls, hex_color: str) -> 'Color':
        """
        Create a Color from a hex string.
        
        Args:
            hex_color (str): Hex color string (e.g., "#FF0000" or "FF0000")
            
        Returns:
            Color: Color instance
        """
        hex_color = hex_color.lstrip('#')
        if len(hex_color) != 6:
            raise ValueError("Hex color must be 6 characters long")
        
        return cls(
            red=int(hex_color[0:2], 16) / 255.0,
            green=int(hex_color[2:4], 16) / 255.0,
            blue=int(hex_color[4:6], 16) / 255.0
        )
    
    @classmethod
    def from_rgb(cls, r: int, g: int, b: int) -> 'Color':
        """
        Create a Color from RGB values (0-255).
        
        Args:
            r (int): Red component (0-255)
            g (int): Green component (0-255)
            b (int): Blue component (0-255)
            
        Returns:
            Color: Color instance
        """
        return cls(red=r/255.0, green=g/255.0, blue=b/255.0)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to Google Docs API color format."""
        return {
            "color": {
                "rgbColor": {
                    "red": self.red,
                    "green": self.green,
                    "blue": self.blue
                }
            }
        }


# Predefined colors
class Colors:
    """Common predefined colors."""
    BLACK = Color(0.0, 0.0, 0.0)
    WHITE = Color(1.0, 1.0, 1.0)
    RED = Color(1.0, 0.0, 0.0)
    GREEN = Color(0.0, 1.0, 0.0)
    BLUE = Color(0.0, 0.0, 1.0)
    YELLOW = Color(1.0, 1.0, 0.0)
    CYAN = Color(0.0, 1.0, 1.0)
    MAGENTA = Color(1.0, 0.0, 1.0)
    GRAY = Color(0.5, 0.5, 0.5)
    LIGHT_GRAY = Color(0.8, 0.8, 0.8)
    DARK_GRAY = Color(0.2, 0.2, 0.2)


@dataclass
class Dimension:
    """Represents a dimension with magnitude and unit."""
    magnitude: float
    unit: str = "PT"  # Points by default
    
    @classmethod
    def points(cls, value: float) -> 'Dimension':
        """Create a dimension in points."""
        return cls(magnitude=value, unit="PT")
    
    @classmethod
    def inches(cls, value: float) -> 'Dimension':
        """Create a dimension in inches."""
        return cls(magnitude=value * 72, unit="PT")  # Convert to points
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to Google Docs API dimension format."""
        return {
            "magnitude": self.magnitude,
            "unit": self.unit
        }


@dataclass
class TextStyle:
    """Comprehensive text style configuration."""
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    strikethrough: Optional[bool] = None
    small_caps: Optional[bool] = None
    font_family: Optional[str] = None
    font_size: Optional[Dimension] = None
    foreground_color: Optional[Color] = None
    background_color: Optional[Color] = None
    baseline_offset: Optional[VerticalAlignment] = None
    link_url: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to Google Docs API text style format."""
        style = {}
        
        if self.bold is not None:
            style["bold"] = self.bold
        if self.italic is not None:
            style["italic"] = self.italic
        if self.underline is not None:
            style["underline"] = self.underline
        if self.strikethrough is not None:
            style["strikethrough"] = self.strikethrough
        if self.small_caps is not None:
            style["smallCaps"] = self.small_caps
        if self.font_family is not None:
            style["weightedFontFamily"] = {
                "fontFamily": self.font_family,
                "weight": 400
            }
        if self.font_size is not None:
            style["fontSize"] = self.font_size.to_dict()
        if self.foreground_color is not None:
            style["foregroundColor"] = self.foreground_color.to_dict()
        if self.background_color is not None:
            style["backgroundColor"] = self.background_color.to_dict()
        if self.baseline_offset is not None:
            style["baselineOffset"] = self.baseline_offset.value
        if self.link_url is not None:
            style["link"] = {"url": self.link_url}
        
        return style


@dataclass
class ParagraphStyle:
    """Paragraph style configuration."""

    namedStyleType: Optional[ParagraphStyleType] = None
    alignment: Optional[TextAlignment] = None
    line_spacing: Optional[float] = None
    direction: Optional[str] = None
    spacing_mode: Optional[str] = None
    space_above: Optional[Dimension] = None
    space_below: Optional[Dimension] = None
    border_between: Optional[Dict] = None
    border_top: Optional[Dict] = None
    border_bottom: Optional[Dict] = None
    border_left: Optional[Dict] = None
    border_right: Optional[Dict] = None
    indent_first_line: Optional[Dimension] = None
    indent_start: Optional[Dimension] = None
    indent_end: Optional[Dimension] = None
    
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to Google Docs API paragraph style format."""
        style = {}
        
        if self.namedStyleType is not None and self.namedStyleType.value != "":
            style["namedStyleType"] = self.namedStyleType.value
        if self.alignment is not None:
            style["alignment"] = self.alignment.value
        if self.line_spacing is not None:
            style["lineSpacing"] = self.line_spacing
        if self.direction is not None:
            style["direction"] = self.direction
        if self.spacing_mode is not None:
            style["spacingMode"] = self.spacing_mode
        if self.space_above is not None:
            style["spaceAbove"] = self.space_above.to_dict()
        if self.space_below is not None:
            style["spaceBelow"] = self.space_below.to_dict()
        if self.indent_first_line is not None:
            style["indentFirstLine"] = self.indent_first_line.to_dict()
        if self.indent_start is not None:
            style["indentStart"] = self.indent_start.to_dict()
        if self.indent_end is not None:
            style["indentEnd"] = self.indent_end.to_dict()
        
        
        # Add borders if specified
        for border_name, border_value in [
            ("borderBetween", self.border_between),
            ("borderTop", self.border_top),
            ("borderBottom", self.border_bottom),
            ("borderLeft", self.border_left),
            ("borderRight", self.border_right)
        ]:
            if border_value is not None:
                style[border_name] = border_value
        
        return style


@dataclass
class TableCellStyle:
    """Table cell style configuration."""
    background_color: Optional[Color] = None
    border_left: Optional[Dict] = None
    border_right: Optional[Dict] = None
    border_top: Optional[Dict] = None
    border_bottom: Optional[Dict] = None
    padding_left: Optional[Dimension] = None
    padding_right: Optional[Dimension] = None
    padding_top: Optional[Dimension] = None
    padding_bottom: Optional[Dimension] = None
    content_alignment: Optional[TableCellVerticalAlignment] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to Google Docs API table cell style format."""
        style = {}
        
        if self.background_color is not None:
            style["backgroundColor"] = self.background_color.to_dict()
        
        # Add borders
        for border_name, border_value in [
            ("borderLeft", self.border_left),
            ("borderRight", self.border_right),
            ("borderTop", self.border_top),
            ("borderBottom", self.border_bottom)
        ]:
            if border_value is not None:
                style[border_name] = border_value
        
        # Add padding
        if any([self.padding_left, self.padding_right, self.padding_top, self.padding_bottom]):
            padding = {}
            if self.padding_left is not None:
                padding["paddingLeft"] = self.padding_left.to_dict()
            if self.padding_right is not None:
                padding["paddingRight"] = self.padding_right.to_dict()
            if self.padding_top is not None:
                padding["paddingTop"] = self.padding_top.to_dict()
            if self.padding_bottom is not None:
                padding["paddingBottom"] = self.padding_bottom.to_dict()
            style.update(padding)
        
        if self.content_alignment is not None:
            style["contentAlignment"] = self.content_alignment.value
        
        return style


def create_border(
    color: Color = Colors.BLACK,
    width: Dimension = Dimension.points(1),
    style: BorderStyle = BorderStyle.SOLID,
    dash_style: Optional[str] = None
) -> Dict[str, Any]:
    """
    Create a border specification.
    
    Args:
        color (Color): Border color
        width (Dimension): Border width
        style (BorderStyle): Border style
        dash_style (str, optional): Dash pattern for dashed borders
        
    Returns:
        Dict[str, Any]: Border specification for Google Docs API
    """
    border = {
        "color": color.to_dict()["color"],
        "width": width.to_dict(),
        "dashStyle": style.value
    }
    
    if dash_style is not None:
        border["dashStyle"] = dash_style
    
    return border


# Common font families
class FontFamilies:
    """Common font family names."""
    ARIAL = "Arial"
    CALIBRI = "Calibri"
    CAMBRIA = "Cambria"
    COMIC_SANS = "Comic Sans MS"
    COURIER_NEW = "Courier New"
    GEORGIA = "Georgia"
    HELVETICA = "Helvetica"
    IMPACT = "Impact"
    ROBOTO = "Roboto"
    TIMES_NEW_ROMAN = "Times New Roman"
    TREBUCHET = "Trebuchet MS"
    VERDANA = "Verdana"


# Common font sizes
class FontSizes:
    """Common font sizes in points."""
    SMALL = Dimension.points(8)
    NORMAL = Dimension.points(11)
    MEDIUM = Dimension.points(12)
    LARGE = Dimension.points(14)
    XLARGE = Dimension.points(18)
    XXLARGE = Dimension.points(24)
    HUGE = Dimension.points(36)


def create_heading_style(level: int, color: Optional[Color] = None) -> ParagraphStyle:
    """
    Create a heading style for the specified level.
    
    Args:
        level (int): Heading level (1-6)
        color (Color, optional): Text color for the heading
        
    Returns:
        ParagraphStyle: Configured heading style
        
    Raises:
        ValueError: If level is not between 1 and 6
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
    
    return ParagraphStyle(
        namedStyleType=heading_types[level],
        space_below=Dimension.points(6)
    )


def create_list_style(list_type: ListType, level: int = 0) -> Dict[str, Any]:
    """
    Create a list style specification.
    
    Args:
        list_type (ListType): Type of list (bulleted or numbered)
        level (int): Nesting level (0-based)
        
    Returns:
        Dict[str, Any]: List style specification
    """
    return {
        "namedStyleType": "NORMAL_TEXT",
        "direction": "LEFT_TO_RIGHT",
        "bullet": {
            "listId": f"{list_type.value.lower()}_list",
            "nestingLevel": level
        }
    }


def merge_styles(*styles: Dict[str, Any]) -> Dict[str, Any]:
    """
    Merge multiple style dictionaries, with later styles overriding earlier ones.
    
    Args:
        *styles: Variable number of style dictionaries
        
    Returns:
        Dict[str, Any]: Merged style dictionary
    """
    merged = {}
    for style in styles:
        if style:
            merged.update(style)
    return merged
