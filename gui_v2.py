import os
import sys
import json
import threading

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

import customtkinter as ctk
import pandas as pd

from PIL import Image
from PIL import ImageTk

from docx import Document
import fitz

from scanner import Scanner
from duplicates import (
    detect_duplicates,
    count_duplicate_files,
    count_duplicate_groups
)

from exporter import Exporter


class Application:

    def __init__(self, root):

        self.root = root

        self.records = []
        self.filtered = []

        self.recent_folders = []

        self.scanner = Scanner()

        self.tk_img = None

        self.dark_mode = True

        self.scan_cancelled = False

        self.current_file = ""

        self.settings_file = "config.json"

        self.folder_var = ctk.StringVar()
        self.search = ctk.StringVar()
        self.extension_filter = ctk.StringVar()

        self.duplicates_only = tk.BooleanVar()

        self.summary = {
            "files": 0,
            "folders": 0,
            "duplicates": 0,
            "size_gb": 0
        }

        self.build_ui()

        self.load_settings()

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_closing
        )


    def resource_path(
        self,
        relative_path
    ):

        try:
            base_path = sys._MEIPASS

        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(
            base_path,
            relative_path
        )


    def build_ui(self):

        self.root.title(
            "PDI Pro V3"
        )

        self.root.geometry(
            "1900x1000"
        )

        ctk.set_appearance_mode(
            "dark"
        )

        self.build_top_bar()

        self.build_main_area()

        self.build_bottom_bar()

        try:

            self.root.iconbitmap(
                self.resource_path(
                    "assets/folder_indexer.ico"
                )
            )

        except Exception:

            pass
    def build_bottom_bar(self):

        bottom = ctk.CTkFrame(
            self.root
        )

        bottom.pack(
            fill="x",
            padx=10,
            pady=10
        )


        ctk.CTkButton(
            bottom,
            text="Open",
            command=self.open_selected
        ).pack(
            side="left",
            padx=5
        )


        ctk.CTkButton(
            bottom,
            text="Open Folder",
            command=self.open_folder
        ).pack(
            side="left",
            padx=5
        )


        ctk.CTkButton(
            bottom,
            text="Copy Path",
            command=self.copy_path
        ).pack(
            side="left",
            padx=5
        )


        ctk.CTkButton(
            bottom,
            text="Export Excel",
            command=self.export_excel
        ).pack(
            side="left",
            padx=5
        )


        ctk.CTkButton(
            bottom,
            text="Export CSV",
            command=self.export_csv
        ).pack(
            side="left",
            padx=5
        )


        self.progress = ctk.CTkProgressBar(
            bottom,
            width=250
        )

        self.progress.pack(
            side="left",
            padx=20
        )

        self.progress.set(0)


        self.status = ctk.CTkLabel(
            bottom,
            text="Ready"
        )

        self.status.pack(
            side="left",
            padx=15
        )


        self.counter_label = ctk.CTkLabel(
            bottom,
            text="Showing 0 / 0"
        )

        self.counter_label.pack(
            side="right",
            padx=15
        )

    def build_top_bar(self):

        top = ctk.CTkFrame(
            self.root
        )

        top.pack(
            fill="x",
            padx=10,
            pady=10
        )


        row1 = ctk.CTkFrame(top)

        row1.pack(
            fill="x",
            pady=5
        )


        row2 = ctk.CTkFrame(top)

        row2.pack(
            fill="x",
            pady=5
        )


        ctk.CTkLabel(
            row1,
            text="Directory"
        ).pack(
            side="left",
            padx=5
        )


        self.dir_entry = ctk.CTkEntry(
            row1,
            width=500,
            textvariable=self.folder_var
        )

        self.dir_entry.pack(
            side="left",
            padx=5
        )


        ctk.CTkButton(
            row1,
            text="Browse",
            command=self.browse
        ).pack(
            side="left",
            padx=5
        )


        ctk.CTkButton(
            row1,
            text="Scan",
            command=self.scan
        ).pack(
            side="left",
            padx=5
        )


        ctk.CTkButton(
            row1,
            text="Stop",
            fg_color="darkred",
            command=self.stop_scan
        ).pack(
            side="left",
            padx=5
        )


        self.theme_switch = ctk.CTkSwitch(
            row1,
            text="Dark Mode",
            command=self.toggle_theme
        )

        self.theme_switch.select()

        self.theme_switch.pack(
            side="left",
            padx=15
        )


        ctk.CTkLabel(
            row2,
            text="Search"
        ).pack(
            side="left",
            padx=5
        )


        self.search_entry = ctk.CTkEntry(
            row2,
            width=300,
            textvariable=self.search
        )

        self.search_entry.pack(
            side="left",
            padx=5
        )


        self.search_entry.bind(
            "<KeyRelease>",
            self.filter_data
        )


        ctk.CTkLabel(
            row2,
            text="Extensions"
        ).pack(
            side="left",
            padx=5
        )


        self.extension_entry = ctk.CTkEntry(
            row2,
            width=250,
            textvariable=self.extension_filter
        )

        self.extension_entry.pack(
            side="left",
            padx=5
        )


        ctk.CTkCheckBox(
            row2,
            text="Duplicates Only",
            variable=self.duplicates_only,
            command=self.filter_data
        ).pack(
            side="left",
            padx=15
        )


        ctk.CTkButton(
            row2,
            text="Clear",
            command=self.clear_search
        ).pack(
            side="left",
            padx=5
        )

    def stop_scan(self):

        self.scan_cancelled = True

        self.status.configure(
            text="Stopping scan..."
        )
    def scan(self):

        self.scan_cancelled = False

        self.progress.set(0)

        self.status.configure(
            text="Starting scan..."
        )

        self.records = []
        self.filtered = []

        threading.Thread(
            target=self.scan_worker,
            daemon=True
        ).start()

    def scan_worker(self):

        folder = self.folder_var.get()

        if not folder:

            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "Error",
                    "Please select a folder."
                )
            )

            return

        self.records = self.scanner.scan(
            folder
        )

        detect_duplicates(
            self.records
        )

        self.filtered = self.records.copy()

        self.root.after(
            0,
            self.populate
        )

    def browse(self):

        folder = filedialog.askdirectory()

        if not folder:
            return

        self.folder_var.set(
            folder
        )

        if folder not in self.recent_folders:

            self.recent_folders.insert(
                0,
                folder
            )

        self.recent_folders = (
            self.recent_folders[:10]
        )
    def toggle_theme(self):

        if self.theme_switch.get():

            ctk.set_appearance_mode(
                "dark"
            )

        else:

            ctk.set_appearance_mode(
                "light"
            )

        self.save_settings()



    def build_main_area(self):

        container = ctk.CTkFrame(
            self.root
        )

        container.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=5
        )


        self.dashboard_frame = ctk.CTkFrame(
            container,
            width=250
        )

        self.dashboard_frame.pack(
            side="left",
            fill="y"
        )

        self.dashboard_frame.pack_propagate(
            False
        )


        self.center_frame = ctk.CTkFrame(
            container
        )

        self.center_frame.pack(
            side="left",
            fill="both",
            expand=True,
            padx=5
        )


        self.preview_frame = ctk.CTkFrame(
            container,
            width=450
        )

        self.preview_frame.pack(
            side="right",
            fill="both"
        )

        self.preview_frame.pack_propagate(
            False
        )


        self.build_dashboard()

        self.build_tree()

        self.build_preview()



    def build_dashboard(self):

        title = ctk.CTkLabel(
            self.dashboard_frame,
            text="Statistics",
            font=("Segoe UI", 18, "bold")
        )

        title.pack(
            pady=10
        )


        self.files_card = ctk.CTkLabel(
            self.dashboard_frame,
            text="Files: 0"
        )

        self.files_card.pack(
            anchor="w",
            padx=10,
            pady=5
        )


        self.folders_card = ctk.CTkLabel(
            self.dashboard_frame,
            text="Folders: 0"
        )

        self.folders_card.pack(
            anchor="w",
            padx=10,
            pady=5
        )


        self.duplicates_card = ctk.CTkLabel(
            self.dashboard_frame,
            text="Duplicates: 0"
        )

        self.duplicates_card.pack(
            anchor="w",
            padx=10,
            pady=5
        )


        self.size_card = ctk.CTkLabel(
            self.dashboard_frame,
            text="Size: 0 GB"
        )

        self.size_card.pack(
            anchor="w",
            padx=10,
            pady=5
        )


        self.largest_file_card = ctk.CTkLabel(
            self.dashboard_frame,
            text="Largest File:"
        )

        self.largest_file_card.pack(
            anchor="w",
            padx=10,
            pady=5
        )



    def build_tree(self):

        columns = (
            "Project",
            "Level1",
            "Level2",
            "Level3",
            "File",
            "Extension",
            "SizeMB",
            "Created",
            "Modified",
            "DuplicateGroup",
            "Path"
        )


        tree_frame = ctk.CTkFrame(
            self.center_frame
        )

        tree_frame.pack(
            fill="both",
            expand=True
        )


        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings"
        )


        widths = {

            "Project": 120,
            "Level1": 120,
            "Level2": 120,
            "Level3": 120,
            "File": 220,
            "Extension": 90,
            "SizeMB": 90,
            "Created": 150,
            "Modified": 150,
            "DuplicateGroup": 120,
            "Path": 500
        }


        for col in columns:

            self.tree.heading(
                col,
                text=col,
                command=lambda c=col:
                self.sort_tree(c)
            )


            self.tree.column(
                col,
                width=widths.get(
                    col,
                    100
                ),
                anchor="w"
            )


        vertical_scroll = ttk.Scrollbar(
            tree_frame,
            orient="vertical",
            command=self.tree.yview
        )


        horizontal_scroll = ttk.Scrollbar(
            tree_frame,
            orient="horizontal",
            command=self.tree.xview
        )


        self.tree.configure(
            yscrollcommand=vertical_scroll.set,
            xscrollcommand=horizontal_scroll.set
        )


        self.tree.grid(
            row=0,
            column=0,
            sticky="nsew"
        )


        vertical_scroll.grid(
            row=0,
            column=1,
            sticky="ns"
        )


        horizontal_scroll.grid(
            row=1,
            column=0,
            sticky="ew"
        )


        tree_frame.grid_rowconfigure(
            0,
            weight=1
        )

        tree_frame.grid_columnconfigure(
            0,
            weight=1
        )


        self.tree.bind(
            "<<TreeviewSelect>>",
            self.preview_file
        )



    def build_preview(self):

        title = ctk.CTkLabel(
            self.preview_frame,
            text="Preview",
            font=("Segoe UI", 18, "bold")
        )

        title.pack(
            pady=10
        )


        self.preview_image = ctk.CTkLabel(
            self.preview_frame,
            text=""
        )

        self.preview_image.pack(
            pady=10
        )


        self.preview_text = ctk.CTkTextbox(
            self.preview_frame
        )

        self.preview_text.pack(
            fill="both",
            expand=True,
            padx=5,
            pady=5
        )

    def scan_worker(self):

        folder = self.folder_var.get()

        if not folder:

            self.root.after(
                0,
                lambda:
                messagebox.showerror(
                    "Error",
                    "Please select a folder."
                )
            )

            return


        extensions = []

        if self.extension_filter.get().strip():

            extensions = [

                ext.strip().lower()

                for ext in
                self.extension_filter.get()
                .split(";")

                if ext.strip()

            ]


        self.root.after(
            0,
            lambda:
            self.status.configure(
                text="Scanning..."
            )
        )


        def progress_callback(
            current,
            total
        ):

            if self.scan_cancelled:
                return


            percent = (
                current /
                max(total, 1)
            )


            self.root.after(
                0,
                lambda p=percent:
                self.progress.set(p)
            )


            self.root.after(
                0,
                lambda:
                self.status.configure(
                    text=
                    f"Scanning "
                    f"{current:,} / "
                    f"{total:,}"
                )
            )


        self.records = self.scanner.scan(
            folder,
            ext_filter=extensions,
            progress_callback=progress_callback
        )


        if self.scan_cancelled:

            self.root.after(
                0,
                lambda:
                self.status.configure(
                    text="Scan cancelled"
                )
            )

            return


        detect_duplicates(
            self.records
        )


        self.filtered = (
            self.records.copy()
        )


        self.root.after(
            0,
            self.populate
        )


        self.root.after(
            0,
            lambda:
            self.progress.set(1)
        )


        self.root.after(
            0,
            lambda:
            self.status.configure(
                text=
                f"Finished - "
                f"{len(self.records):,} files"
            )
        )


    def populate(self):

        self.tree.delete(
            *self.tree.get_children()
        )


        for row in self.filtered:

            levels = row.get(
                "Levels",
                []
            )


            self.tree.insert(
                "",
                "end",
                values=(

                    row.get(
                        "Project",
                        ""
                    ),

                    levels[0]
                    if len(levels) > 0
                    else "",

                    levels[1]
                    if len(levels) > 1
                    else "",

                    levels[2]
                    if len(levels) > 2
                    else "",

                    row.get(
                        "File",
                        ""
                    ),

                    row.get(
                        "Extension",
                        ""
                    ),

                    row.get(
                        "SizeMB",
                        ""
                    ),

                    row.get(
                        "Created",
                        ""
                    ),

                    row.get(
                        "Modified",
                        ""
                    ),

                    row.get(
                        "DuplicateGroup",
                        ""
                    ),

                    row.get(
                        "Path",
                        ""
                    )
                )
            )


        self.update_counter()

        self.update_dashboard()



    def filter_data(
        self,
        event=None
    ):

        keyword = (
            self.search.get()
            .strip()
            .lower()
        )


        duplicates_only = (
            self.duplicates_only.get()
        )


        self.filtered = []


        for row in self.records:


            if duplicates_only:

                if not row.get(
                    "DuplicateGroup"
                ):

                    continue


            if keyword:


                searchable = " ".join([

                    str(
                        row.get(
                            "Project",
                            ""
                        )
                    ),

                    str(
                        row.get(
                            "File",
                            ""
                        )
                    ),

                    str(
                        row.get(
                            "Extension",
                            ""
                        )
                    ),

                    str(
                        row.get(
                            "Path",
                            ""
                        )
                    )

                ]).lower()


                if keyword not in searchable:

                    continue


            self.filtered.append(
                row
            )


        self.populate()


        self.status.configure(
            text=
            f"{len(self.filtered):,} "
            f"matching files"
        )



    def clear_search(self):

        self.search.set("")

        self.filtered = (
            self.records.copy()
        )

        self.populate()

        self.status.configure(
            text="Search cleared"
        )


    def update_counter(self):

        self.counter_label.configure(
            text=
            f"Showing "
            f"{len(self.filtered):,}"
            f" / "
            f"{len(self.records):,}"
        )
    def update_dashboard(self):

        total_files = len(
            self.records
        )


        total_size = sum(
            row.get(
                "SizeBytes",
                0
            )
            for row in self.records
        )


        size_gb = round(
            total_size / (1024 ** 3),
            2
        )


        folders = set()

        for row in self.records:

            folders.add(
                os.path.dirname(
                    row.get(
                        "Path",
                        ""
                    )
                )
            )


        duplicates = count_duplicate_files(
            self.records
        )


        largest_file = ""

        largest_size = 0


        for row in self.records:

            if row.get(
                "SizeBytes",
                0
            ) > largest_size:

                largest_size = row["SizeBytes"]

                largest_file = row.get(
                    "File",
                    ""
                )


        self.files_card.configure(
            text=f"Files: {total_files:,}"
        )


        self.folders_card.configure(
            text=f"Folders: {len(folders):,}"
        )


        self.duplicates_card.configure(
            text=f"Duplicates: {duplicates:,}"
        )


        self.size_card.configure(
            text=f"Size: {size_gb:.2f} GB"
        )


        self.largest_file_card.configure(
            text=f"Largest: {largest_file}"
        )



    def sort_tree(
        self,
        column,
        reverse=False
    ):

        items = []


        for child in self.tree.get_children():

            value = self.tree.set(
                child,
                column
            )

            items.append(
                (
                    value,
                    child
                )
            )


        try:

            items.sort(
                key=lambda x:
                float(x[0]),
                reverse=reverse
            )

        except Exception:

            items.sort(
                key=lambda x:
                str(x[0]).lower(),
                reverse=reverse
            )


        for index, (_, child) in enumerate(items):

            self.tree.move(
                child,
                "",
                index
            )


        self.tree.heading(
            column,
            command=lambda:
            self.sort_tree(
                column,
                not reverse
            )
        )



    def open_selected(
        self,
        event=None
    ):

        selected = (
            self.tree.selection()
        )


        if not selected:

            return


        path = self.tree.item(
            selected[0],
            "values"
        )[-1]


        if os.path.exists(path):

            os.startfile(path)



    def open_folder(self):

        selected = (
            self.tree.selection()
        )


        if not selected:

            return


        path = self.tree.item(
            selected[0],
            "values"
        )[-1]


        folder = os.path.dirname(
            path
        )


        if os.path.exists(folder):

            os.startfile(folder)



    def copy_path(self):

        selected = (
            self.tree.selection()
        )


        if not selected:

            return


        paths = []


        for item in selected:

            values = self.tree.item(
                item,
                "values"
            )

            paths.append(
                values[-1]
            )


        self.root.clipboard_clear()

        self.root.clipboard_append(
            "\n".join(paths)
        )


        self.status.configure(
            text=
            f"{len(paths)} path(s) copied"
        )



    def preview_file(
        self,
        event=None
    ):

        selected = self.tree.selection()


        if not selected:

            return


        path = self.tree.item(
            selected[0],
            "values"
        )[-1]


        self.preview_text.delete(
            "1.0",
            "end"
        )


        self.preview_image.configure(
            image=None
        )


        self.tk_img = None


        if not os.path.exists(path):

            return


        ext = os.path.splitext(
            path
        )[1].lower()


        try:

            if ext in (
                ".txt",
                ".log",
                ".json",
                ".xml",
                ".cfg",
                ".ini",
                ".py",
                ".ps1",
                ".bat",
                ".yaml",
                ".yml"
            ):

                self.preview_text_file(path)


            elif ext == ".csv":

                self.preview_csv(path)


            elif ext in (
                ".xlsx",
                ".xls"
            ):

                self.preview_excel(path)


            elif ext == ".docx":

                self.preview_docx(path)


            elif ext == ".pdf":

                self.preview_pdf(path)


            elif ext in (
                ".jpg",
                ".jpeg",
                ".png",
                ".bmp",
                ".gif",
                ".webp"
            ):

                self.preview_image_file(path)


            else:

                self.preview_generic(path)


        except Exception as ex:

            self.preview_text.insert(
                "1.0",
                str(ex)
            )



    def preview_text_file(
        self,
        path
    ):

        with open(
            path,
            "r",
            encoding="utf-8",
            errors="ignore"
        ) as file:

            content = file.read(
                15000
            )


        self.preview_text.insert(
            "1.0",
            content
        )
    def preview_csv(
        self,
        path
    ):

        try:

            df = pd.read_csv(
                path,
                nrows=100
            )

            self.preview_text.insert(
                "1.0",
                df.to_string(
                    index=False
                )
            )

        except Exception as ex:

            self.preview_text.insert(
                "1.0",
                str(ex)
            )



    def preview_excel(
        self,
        path
    ):

        output = []

        excel = pd.ExcelFile(
            path
        )


        output.append(
            f"Workbook: {os.path.basename(path)}\n"
        )

        output.append(
            f"Sheets: {excel.sheet_names}\n"
        )


        for sheet in excel.sheet_names[:3]:

            output.append(
                f"\n===== {sheet} =====\n"
            )

            try:

                df = pd.read_excel(
                    path,
                    sheet_name=sheet,
                    nrows=25
                )

                output.append(
                    df.to_string(
                        index=False
                    )
                )

            except Exception as ex:

                output.append(
                    str(ex)
                )


        self.preview_text.insert(
            "1.0",
            "\n".join(output)
        )



    def preview_docx(
        self,
        path
    ):

        document = Document(
            path
        )

        text = []


        for paragraph in document.paragraphs:

            value = paragraph.text.strip()

            if value:

                text.append(
                    value
                )


        self.preview_text.insert(
            "1.0",
            "\n".join(text)[:15000]
        )



    def preview_pdf(
        self,
        path
    ):

        pdf = fitz.open(
            path
        )

        output = []

        output.append(
            f"Pages: {pdf.page_count}\n"
        )


        for index in range(
            min(pdf.page_count, 5)
        ):

            page = pdf[index]

            output.append(
                f"\n===== PAGE {index + 1} =====\n"
            )

            output.append(
                page.get_text()[:3000]
            )


        self.preview_text.insert(
            "1.0",
            "\n".join(output)
        )



    def preview_image_file(
        self,
        path
    ):

        image = Image.open(
            path
        )

        image.thumbnail(
            (400, 400)
        )


        self.tk_img = ImageTk.PhotoImage(
            image
        )


        self.preview_image.configure(
            image=self.tk_img
        )


        self.preview_text.insert(
            "1.0",
            f"""
File:
{os.path.basename(path)}

Resolution:
{image.width} x {image.height}

Size:
{os.path.getsize(path):,} bytes
"""
        )



    def preview_generic(
        self,
        path
    ):

        stat = os.stat(
            path
        )

        self.preview_text.insert(
            "1.0",
            f"""
Filename:
{os.path.basename(path)}

Path:
{path}

Extension:
{os.path.splitext(path)[1]}

Size:
{stat.st_size:,} bytes

Preview not available.
"""
        )



    def export_excel(self):

        if not self.records:

            return


        filename = filedialog.asksaveasfilename(

            defaultextension=".xlsx",

            filetypes=[
                (
                    "Excel Files",
                    "*.xlsx"
                )
            ]

        )


        if not filename:

            return


        stats_df = pd.DataFrame(
            self.get_extension_statistics()
        )


        summary_df = (
            self.get_summary_dataframe()
        )


        errors_df = pd.DataFrame(
            self.scanner.errors
        )


        Exporter.export_excel(
            self.records,
            stats_df,
            summary_df,
            errors_df,
            filename
        )


        messagebox.showinfo(
            "Success",
            "Excel export completed."
        )



    def export_csv(self):

        if not self.records:

            return


        filename = filedialog.asksaveasfilename(

            defaultextension=".csv",

            filetypes=[
                (
                    "CSV Files",
                    "*.csv"
                )
            ]

        )


        if not filename:

            return


        Exporter.export_csv(
            self.records,
            filename
        )


        messagebox.showinfo(
            "Success",
            "CSV export completed."
        )



    def get_extension_statistics(self):

        if not self.records:

            return pd.DataFrame()


        df = pd.DataFrame(
            self.records
        )


        return (
            df.groupby(
                "Extension"
            )
            .size()
            .reset_index(
                name="Count"
            )
            .sort_values(
                "Count",
                ascending=False
            )
        )



    def get_summary_dataframe(self):

        if not self.records:

            return pd.DataFrame()


        df = pd.DataFrame(
            self.records
        )


        return pd.DataFrame([{

            "Files Scanned":
                len(df),

            "Duplicate Files":
                count_duplicate_files(
                    self.records
                ),

            "Duplicate Groups":
                count_duplicate_groups(
                    self.records
                ),

            "Total Size GB":
                round(
                    df["SizeBytes"].sum()
                    /
                    (1024 ** 3),
                    2
                )

        }])



    def load_settings(self):

        if not os.path.exists(
            self.settings_file
        ):

            return


        try:

            with open(
                self.settings_file,
                "r",
                encoding="utf-8"
            ) as file:

                data = json.load(
                    file
                )


            self.folder_var.set(
                data.get(
                    "last_folder",
                    ""
                )
            )


        except Exception:

            pass



    def save_settings(self):

        try:

            data = {

                "last_folder":
                    self.folder_var.get(),

                "window_size":
                    self.root.geometry()

            }


            with open(
                self.settings_file,
                "w",
                encoding="utf-8"
            ) as file:

                json.dump(
                    data,
                    file,
                    indent=4
                )


        except Exception:

            pass



    def on_closing(self):

        self.save_settings()

        self.root.destroy()



def main():

    root = ctk.CTk()

    app = Application(
        root
    )

    root.mainloop()



if __name__ == "__main__":

    main()