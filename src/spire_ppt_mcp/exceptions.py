class PptMCPError(Exception):
    """Base exception for Ppt MCP errors."""
    pass

class PresentationError(PptMCPError):
    """Raised when presentation operations fail."""
    pass

class SlideError(PptMCPError):
    """Raised when slide operations fail."""
    pass

class ShapeError(PptMCPError):
    """Raised when shape operations fail."""
    pass

class ChartError(PptMCPError):
    """Raised when chart operations fail."""
    pass

class SmartArtError(PptMCPError):
    """Raised when smartart operations fail."""
    pass

class TableError(PptMCPError):
    """Raised when table operations fail."""
    pass

class ConversionError(Exception):
    """Exception raised for errors during file conversion."""
    pass