from collections import defaultdict

from utils import sha256_file


def detect_duplicates(records):
    """
    Detect duplicate files using a two-step process:

    1. Group files by size.
    2. Compute SHA256 only for files sharing the same size.

    Updates:
        record["DuplicateGroup"]

    Returns:
        records
    """

    size_map = defaultdict(list)

    # First pass: group by file size
    for record in records:

        try:
            size_map[
                record["SizeBytes"]
            ].append(record)

        except KeyError:
            continue

    duplicate_counter = 1

    # Second pass: hash only same-size groups
    for size_group in size_map.values():

        if len(size_group) < 2:
            continue

        hash_map = defaultdict(list)

        for record in size_group:

            try:

                sha = sha256_file(
                    record["Path"]
                )

                if sha:
                    hash_map[sha].append(
                        record
                    )

            except Exception:
                continue

        # Assign duplicate groups
        for duplicate_group in hash_map.values():

            if len(duplicate_group) < 2:
                continue

            duplicate_id = (
                f"DUP-{duplicate_counter:04d}"
            )

            for record in duplicate_group:

                record["DuplicateGroup"] = (
                    duplicate_id
                )

            duplicate_counter += 1

    return records


def count_duplicate_files(records):
    """
    Returns count of files classified
    as duplicates.
    """

    return sum(
        1
        for record in records
        if record.get(
            "DuplicateGroup",
            ""
        )
    )


def count_duplicate_groups(records):
    """
    Returns count of duplicate groups.
    """

    groups = {
        record.get("DuplicateGroup")
        for record in records
        if record.get("DuplicateGroup")
    }

    return len(groups)