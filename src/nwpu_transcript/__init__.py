"""
NWPU Transcript Tools

Convert NWPU transcript PDFs (Chinese/English) into the official Excel
upload template.
"""

__all__ = [
    "parse_chinese",
    "parse_english",
    "write_to_template",
]

__version__ = "0.1.0"

from .parser import parse_chinese, parse_english  # noqa: E402
from .excel import write_to_template  # noqa: E402

