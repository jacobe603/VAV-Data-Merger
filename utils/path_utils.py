"""Utility functions for file path handling."""


def sanitize_path(p: str) -> str:
    """Strip surrounding quotes and whitespace from a filesystem path.

    Handles paths copied via Windows "Copy as path" which include quotes.

    Args:
        p: File path string, potentially with quotes

    Returns:
        Cleaned path string without surrounding quotes
    """
    if not p:
        return p
    p = p.strip()
    if (p.startswith('"') and p.endswith('"')) or (p.startswith("'") and p.endswith("'")):
        return p[1:-1].strip()
    return p


def allowed_file(filename: str, allowed_extensions: set = None) -> bool:
    """Check if a filename has an allowed extension.

    Args:
        filename: Name of the file to check
        allowed_extensions: Set of allowed extensions (default: {'xlsx', 'xls', 'tw2', 'mdb'})

    Returns:
        True if file extension is allowed, False otherwise
    """
    if allowed_extensions is None:
        allowed_extensions = {'xlsx', 'xls', 'tw2', 'mdb'}
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions
