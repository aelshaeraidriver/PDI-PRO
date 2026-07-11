import hashlib
import os


def format_mb(size_bytes):
    """
    Convert bytes to MB.
    """
    return round(size_bytes / (1024 * 1024), 2)


def sha256_file(filepath):
    """
    Calculate SHA256 hash of a file.

    Returns:
        str: SHA256 hash
        "" : if file cannot be read
    """

    try:

        filepath = os.path.normpath(filepath)

        sha = hashlib.sha256()

        with open(filepath, "rb") as file:

            while True:

                chunk = file.read(
                    1024 * 1024
                )  # 1 MB chunks

                if not chunk:
                    break

                sha.update(chunk)

        return sha.hexdigest()

    except Exception:

        return ""


def auto_width(ws):
    """
    Automatically adjust Excel column widths.
    """

    for column in ws.columns:

        max_length = 0

        try:
            column_letter = (
                column[0].column_letter
            )
        except Exception:
            continue

        for cell in column:

            try:

                value = ""

                if cell.value is not None:
                    value = str(cell.value)

                max_length = max(
                    max_length,
                    len(value)
                )

            except Exception:
                pass

        adjusted_width = min(
            max_length + 4,
            80
        )

        ws.column_dimensions[
            column_letter
        ].width = adjusted_width