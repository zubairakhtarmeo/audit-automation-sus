"""Compatibility shim for Python 3.13+.

Python removed the stdlib `imghdr` module in 3.13. Some older versions of
Streamlit still import it.

This file provides a minimal subset of the original API that Streamlit uses.
"""

from __future__ import annotations

from typing import BinaryIO, Optional, Union


def what(file: Union[str, BinaryIO, None] = None, h: Optional[bytes] = None) -> Optional[str]:
    """Return image type based on header bytes.

    This implements only the subset Streamlit needs.

    Args:
        file: A filename or binary file object.
        h: Optional header bytes.

    Returns:
        A lowercase image type string like 'png', 'jpeg', 'gif', 'webp', etc.,
        or None if unknown.
    """

    header = h

    if header is None:
        if file is None:
            return None

        if isinstance(file, (str, bytes)):
            try:
                with open(file, "rb") as f:
                    header = f.read(64)
            except OSError:
                return None
        else:
            try:
                pos = file.tell()
            except Exception:
                pos = None
            try:
                header = file.read(64)
            except Exception:
                return None
            finally:
                if pos is not None:
                    try:
                        file.seek(pos)
                    except Exception:
                        pass

    if not header:
        return None

    b = header

    # PNG
    if b.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"

    # JPEG
    if b.startswith(b"\xff\xd8\xff"):
        return "jpeg"

    # GIF
    if b.startswith(b"GIF87a") or b.startswith(b"GIF89a"):
        return "gif"

    # WEBP (RIFF....WEBP)
    if len(b) >= 12 and b[0:4] == b"RIFF" and b[8:12] == b"WEBP":
        return "webp"

    # BMP
    if b.startswith(b"BM"):
        return "bmp"

    # TIFF
    if b.startswith(b"II*\x00") or b.startswith(b"MM\x00*"):
        return "tiff"

    return None
