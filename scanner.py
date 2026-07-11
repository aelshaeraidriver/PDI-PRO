import os
from datetime import datetime

from utils import format_mb


class Scanner:

    def __init__(self):
        self.folder_cache = {}
        self.errors = []

    def folder_size_mb(self, folder):
        """
        Returns folder size in MB using a cache
        to avoid recalculating repeatedly.
        """

        folder = os.path.normpath(folder)

        if folder in self.folder_cache:
            return self.folder_cache[folder]

        total = 0

        try:
            for root, _, files in os.walk(folder):

                for file_name in files:

                    file_path = os.path.join(
                        root,
                        file_name
                    )

                    try:
                        total += os.path.getsize(
                            file_path
                        )

                    except Exception:
                        pass

        except Exception:
            pass

        size_mb = format_mb(total)

        self.folder_cache[folder] = size_mb

        return size_mb

    def scan(
        self,
        root_folder,
        ext_filter=None,
        progress_callback=None
    ):
        """
        Scan directory and return file metadata.
        """

        root_folder = os.path.normpath(root_folder)

        # Reset caches/errors for new scan
        self.folder_cache.clear()
        self.errors.clear()

        records = []
        files_list = []

        try:

            for root, _, files in os.walk(root_folder):

                for file_name in files:

                    files_list.append(
                        os.path.join(root, file_name)
                    )

        except Exception as ex:

            self.errors.append({
                "Path": root_folder,
                "Error": str(ex)
            })

            return records

        total = len(files_list)

        project_name = os.path.basename(
            os.path.normpath(root_folder)
        )

        for index, path in enumerate(files_list):

            try:

                path = os.path.normpath(path)

                ext = os.path.splitext(
                    path
                )[1].lower()

                if ext_filter:

                    if ext not in ext_filter:
                        continue

                stat = os.stat(path)

                rel_path = os.path.relpath(
                    path,
                    root_folder
                )

                parts = rel_path.split(os.sep)

                record = {

                    "Project":
                        project_name,

                    "Levels":
                        [project_name] + parts[:-1],

                    "File":
                        parts[-1],

                    "Extension":
                        ext,

                    "SizeMB":
                        format_mb(
                            stat.st_size
                        ),

                    "SizeBytes":
                        stat.st_size,

                    "FolderSizeMB":
                        self.folder_size_mb(
                            os.path.dirname(path)
                        ),

                    "Created":
                        datetime.fromtimestamp(
                            stat.st_ctime
                        ).strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),

                    "Modified":
                        datetime.fromtimestamp(
                            stat.st_mtime
                        ).strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),

                    "Accessed":
                        datetime.fromtimestamp(
                            stat.st_atime
                        ).strftime(
                            "%Y-%m-%d %H:%M:%S"
                        ),

                    "DuplicateGroup":
                        "",

                    "Path":
                        path
                }

                records.append(record)

            except Exception as ex:

                self.errors.append({
                    "Path": path,
                    "Error": str(ex)
                })

            if progress_callback:

                try:
                    progress_callback(
                        index + 1,
                        total
                    )

                except Exception:
                    pass

        return records