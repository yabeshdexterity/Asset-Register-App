import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
import glob
import json
import subprocess
from datetime import timedelta, datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

APP_TITLE = "Attendance Management Software"

APP_FOLDER = os.path.dirname(os.path.abspath(__file__))

LOGO_FILE = os.path.join(
    APP_FOLDER,
    "Attendance_logo.png"
)

APP_DATA_FOLDER = os.path.join(
    os.environ.get("LOCALAPPDATA", APP_FOLDER),
    "AttendanceManagementSoftware"
)

os.makedirs(
    APP_DATA_FOLDER,
    exist_ok=True
)

BACKUP_DATA_FILE = os.path.join(
    APP_DATA_FOLDER,
    "backup_paths.json"
)

DEFAULT_CATEGORIES = [
    "S3D",
    "PDS",
    "SPI",
    "SPPID",
    "Other"
]


# ============================================================
# COLORS
# ============================================================

SIDEBAR_BG = "#1F2937"
SIDEBAR_BUTTON = "#273449"
SIDEBAR_ACTIVE = "#2563EB"

MAIN_BG = "#F5F7FA"
CARD_BG = "#FFFFFF"

TEXT_DARK = "#1F2937"
TEXT_LIGHT = "#FFFFFF"
TEXT_GRAY = "#6B7280"

SUCCESS = "#16A34A"
DANGER = "#DC2626"
WARNING = "#D97706"


# ============================================================
# GLOBAL VARIABLES
# ============================================================

current_page = None

logo_image = None

attendance_file_var = None
attendance_status_var = None

category_filter_var = None
search_var = None

backup_tree = None
backup_status_label = None

backup_paths = []


# ============================================================
# PATH NORMALIZATION
# ============================================================

def normalize_path(path):

    if not path:
        return ""

    path = str(path).strip()

    path = path.replace("/", "\\")

    if path.startswith("\\") and not path.startswith("\\\\"):
        path = "\\" + path

    return path


# ============================================================
# BACKUP DATA
# ============================================================

def load_backup_paths():

    if not os.path.exists(BACKUP_DATA_FILE):
        return []

    try:

        with open(
            BACKUP_DATA_FILE,
            "r",
            encoding="utf-8"
        ) as file:

            data = json.load(file)

        if not isinstance(data, list):
            return []

        for item in data:

            if "category" not in item:
                item["category"] = "Other"

            if "name" not in item:
                item["name"] = ""

            if "path" not in item:
                item["path"] = ""

            item["path"] = normalize_path(
                item["path"]
            )

        return data

    except Exception as error:

        messagebox.showerror(
            "Error",
            f"Could not load backup paths.\n\n{error}"
        )

        return []


def save_backup_paths():

    try:

        with open(
            BACKUP_DATA_FILE,
            "w",
            encoding="utf-8"
        ) as file:

            json.dump(
                backup_paths,
                file,
                indent=4,
                ensure_ascii=False
            )

    except Exception as error:

        messagebox.showerror(
            "Error",
            f"Could not save backup paths.\n\n{error}"
        )


# ============================================================
# GET CATEGORIES
# ============================================================

def get_categories():

    categories = DEFAULT_CATEGORIES.copy()

    for item in backup_paths:

        category = item.get(
            "category",
            "Other"
        )

        if category not in categories:
            categories.append(category)

    return categories


# ============================================================
# OPEN FOLDER
# ============================================================

def open_folder(path):

    path = normalize_path(path)

    if not path:

        messagebox.showwarning(
            "Invalid Path",
            "The folder path is empty."
        )

        return

    if not os.path.exists(path):

        messagebox.showerror(
            "Folder Not Available",
            "The following folder could not be accessed:\n\n"
            f"{path}\n\n"
            "Please check the network connection or "
            "server availability."
        )

        return

    try:

        subprocess.Popen(
            ["explorer.exe", path],
            shell=False
        )

    except Exception as error:

        messagebox.showerror(
            "Error",
            f"Could not open the folder.\n\n{error}"
        )


# ============================================================
# TEST BACKUP PATH
# ============================================================

def test_path(path):

    path = normalize_path(path)

    if not path:

        messagebox.showwarning(
            "Path Test",
            "The path is empty."
        )

        return

    if os.path.exists(path):

        messagebox.showinfo(
            "Path Test",
            "Path is accessible.\n\n"
            f"{path}"
        )

    else:

        messagebox.showerror(
            "Path Test",
            "Path is NOT accessible.\n\n"
            f"{path}"
        )


# ============================================================
# TIME CONVERSION
# ============================================================

def convert_to_time(value):

    try:

        if pd.isna(value) or value == "":
            return None

        if isinstance(value, timedelta):
            return value

        if isinstance(value, datetime):
            return timedelta(
                hours=value.hour,
                minutes=value.minute,
                seconds=value.second
            )

        value_string = str(value).strip()

        if ":" not in value_string:
            return None

        parts = value_string.split(":")

        hours = int(parts[0])
        minutes = int(parts[1])

        seconds = 0

        if len(parts) >= 3:

            try:
                seconds = int(float(parts[2]))
            except:
                seconds = 0

        return timedelta(
            hours=hours,
            minutes=minutes,
            seconds=seconds
        )

    except:

        return None


# ============================================================
# FIND EMPLOYEE SECTIONS
# ============================================================

def split_employee_sections(df):

    employees = {}

    current_employee = None
    data = []

    for _, row in df.iterrows():

        values = [
            str(x).strip()
            for x in row
            if str(x).lower() != "nan"
        ]

        employee_header = None

        for value in values:

            if " - " in value:

                employee_header = value
                break

        if employee_header:

            if current_employee and data:

                employees[current_employee] = pd.DataFrame(
                    data
                )

            current_employee = employee_header
            data = []

            continue

        if current_employee:

            data.append(
                row.tolist()
            )

    if current_employee and data:

        employees[current_employee] = pd.DataFrame(
            data
        )

    return employees


# ============================================================
# PROCESS ATTENDANCE
# ============================================================

def process_attendance(input_file):

    try:

        # ----------------------------------------------------
        # Read raw biometric file
        # ----------------------------------------------------

        df = pd.read_excel(
            input_file,
            header=None
        )

        employees = split_employee_sections(
            df
        )

        if not employees:

            raise ValueError(
                "No employee attendance data was found "
                "in the selected Excel file."
            )

        # ----------------------------------------------------
        # Sort employees
        # ----------------------------------------------------

        employees = dict(
            sorted(
                employees.items(),
                key=lambda x: x[0]
            )
        )

        # ----------------------------------------------------
        # Output name
        # ----------------------------------------------------

        output_folder = os.path.dirname(
            input_file
        )

        output_file = os.path.join(
            output_folder,
            "Month_Timesheet.xlsx"
        )

        if os.path.exists(output_file):

            try:
                os.remove(output_file)

            except PermissionError:

                raise PermissionError(
                    "Month_Timesheet.xlsx is currently open.\n\n"
                    "Please close the Excel file and try again."
                )

        # ----------------------------------------------------
        # Create Excel writer
        # ----------------------------------------------------

        writer = pd.ExcelWriter(
            output_file,
            engine="openpyxl"
        )

        employee_count = 0

        # ----------------------------------------------------
        # Process every employee
        # ----------------------------------------------------

        for emp, emp_df in employees.items():

            try:

                # ------------------------------------------------
                # Take first 12 biometric columns
                # ------------------------------------------------

                if emp_df.shape[1] < 12:

                    continue

                emp_df = emp_df.iloc[
                    :,
                    :12
                ].copy()

                emp_df.columns = [

                    "Sl.No",
                    "Date",
                    "Shift",
                    "First IN",
                    "Last OUT",
                    "1st Half",
                    "2nd Half",
                    "Gross Work Hrs",
                    "OUT Time",
                    "N-Punch Work Hrs",
                    "O/P Code For Status",
                    "Man Entry"

                ]

                # ------------------------------------------------
                # Employee name
                # ------------------------------------------------

                if " - " in emp:

                    employee_name = emp.split(
                        " - ",
                        1
                    )[1].strip()

                else:

                    employee_name = emp.strip()

                # ------------------------------------------------
                # Add Day
                # ------------------------------------------------

                emp_df.insert(
                    2,
                    "Day",
                    ""
                )

                # ------------------------------------------------
                # Add Name
                # ------------------------------------------------

                emp_df.insert(
                    3,
                    "Name",
                    employee_name
                )

                # ------------------------------------------------
                # Remove unwanted columns
                # ------------------------------------------------

                emp_df.drop(
                    columns=[
                        "1st Half",
                        "2nd Half",
                        "O/P Code For Status",
                        "Man Entry"
                    ],
                    inplace=True,
                    errors="ignore"
                )

                # ------------------------------------------------
                # Additional columns
                # ------------------------------------------------

                emp_df["Required Hrs"] = ""

                emp_df["Lag Hrs"] = ""

                emp_df["Hours Deducted"] = ""

                emp_df["Leave/Permission Remarks"] = ""

                emp_df["Approval"] = ""

                # ------------------------------------------------
                # Date conversion
                # ------------------------------------------------

                emp_df["Date"] = pd.to_datetime(
                    emp_df["Date"],
                    dayfirst=True,
                    errors="coerce"
                )

                # ------------------------------------------------
                # Day calculation
                # ------------------------------------------------

                emp_df["Day"] = (
                    emp_df["Date"]
                    .dt.strftime("%A")
                )

                # ------------------------------------------------
                # Convert time columns
                # ------------------------------------------------

                for col in [

                    "First IN",
                    "Last OUT",
                    "Gross Work Hrs",
                    "OUT Time",
                    "N-Punch Work Hrs"

                ]:

                    emp_df[col] = (
                        emp_df[col]
                        .apply(convert_to_time)
                    )

                # ------------------------------------------------
                # Required hours
                # ------------------------------------------------

                required_time = timedelta(
                    hours=8,
                    minutes=15
                )

                emp_df["Required Hrs"] = (

                    emp_df["N-Punch Work Hrs"]
                    .apply(
                        lambda x:
                        required_time
                        if pd.notna(x)
                        else None
                    )

                )

                # ------------------------------------------------
                # Lag hours
                # ------------------------------------------------

                emp_df["Lag Hrs"] = (

                    emp_df["N-Punch Work Hrs"]
                    .apply(

                        lambda x:

                        timedelta(0)

                        if pd.notna(x)
                        and x == timedelta(0)

                        else

                        (
                            max(
                                required_time - x,
                                timedelta(0)
                            )

                            if pd.notna(x)

                            else None
                        )

                    )

                )

                # ------------------------------------------------
                # Sheet name
                # ------------------------------------------------

                sheet_name = emp.replace(
                    " - ",
                    "-"
                )

                # Excel sheet limit
                sheet_name = sheet_name[:31]

                # Remove invalid characters
                for char in [
                    "\\",
                    "/",
                    "*",
                    "?",
                    ":",
                    "[",
                    "]"
                ]:

                    sheet_name = sheet_name.replace(
                        char,
                        ""
                    )

                if not sheet_name:
                    sheet_name = f"Employee{employee_count + 1}"

                # ------------------------------------------------
                # Make unique sheet name
                # ------------------------------------------------

                original_sheet_name = sheet_name

                counter = 1

                while sheet_name in writer.book.sheetnames:

                    suffix = f"_{counter}"

                    sheet_name = (
                        original_sheet_name[
                            :31 - len(suffix)
                        ]
                        + suffix
                    )

                    counter += 1

                # ------------------------------------------------
                # Export sheet
                # ------------------------------------------------

                emp_df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

                ws = writer.sheets[
                    sheet_name
                ]

                employee_count += 1

                # =================================================
                # EXCEL FORMATTING
                # =================================================

                bold = Font(
                    bold=True
                )

                yellow_fill = PatternFill(
                    start_color="FFFF00",
                    end_color="FFFF00",
                    fill_type="solid"
                )

                white_fill = PatternFill(
                    start_color="FFFFFF",
                    end_color="FFFFFF",
                    fill_type="solid"
                )

                red_font = Font(
                    bold=True,
                    color="FF0000"
                )

                center = Alignment(
                    horizontal="center",
                    vertical="center"
                )

                border = Border(
                    left=Side(style="thin"),
                    right=Side(style="thin"),
                    top=Side(style="thin"),
                    bottom=Side(style="thin")
                )

                # ------------------------------------------------
                # Header
                # ------------------------------------------------

                for cell in ws[1]:

                    cell.font = bold

                    cell.fill = yellow_fill

                    cell.alignment = center

                    cell.border = border

                # ------------------------------------------------
                # Data formatting
                # ------------------------------------------------

                for row in ws.iter_rows(
                    min_row=2
                ):

                    for cell in row:

                        cell.alignment = center

                        cell.border = border

                    # Date
                    row[1].number_format = (
                        "dd/mm/yyyy"
                    )

                    # First IN
                    row[4].number_format = (
                        "hh:mm AM/PM"
                    )

                    # Last OUT
                    row[5].number_format = (
                        "hh:mm AM/PM"
                    )

                    # Sunday
                    date_value = row[1].value

                    if date_value:

                        try:

                            if (
                                pd.to_datetime(
                                    date_value
                                ).day_name()
                                == "Sunday"
                            ):

                                for cell in row:

                                    cell.fill = white_fill

                                    cell.font = red_font

                        except:

                            pass

                # ------------------------------------------------
                # Hour format
                # ------------------------------------------------

                for r in range(
                    2,
                    ws.max_row + 1
                ):

                    for cell in ws[r]:

                        header = ws.cell(
                            row=1,
                            column=cell.column
                        ).value

                        if header in [

                            "Gross Work Hrs",
                            "OUT Time",
                            "N-Punch Work Hrs",
                            "Required Hrs",
                            "Lag Hrs",
                            "Hours Deducted"

                        ]:

                            cell.number_format = "[h]:mm"

                # ------------------------------------------------
                # Total row
                # ------------------------------------------------

                total_row = ws.max_row + 1

                total_columns = []

                for col_index in range(
                    1,
                    ws.max_column + 1
                ):

                    header = ws.cell(
                        row=1,
                        column=col_index
                    ).value

                    if header in [

                        "Gross Work Hrs",
                        "OUT Time",
                        "N-Punch Work Hrs",
                        "Required Hrs",
                        "Lag Hrs",
                        "Hours Deducted"

                    ]:

                        total_columns.append(
                            col_index
                        )

                for col_index in total_columns:

                    letter = ws.cell(
                        row=1,
                        column=col_index
                    ).column_letter

                    cell = ws.cell(
                        row=total_row,
                        column=col_index
                    )

                    cell.value = (
                        f"=SUM("
                        f"{letter}2:"
                        f"{letter}{total_row - 1}"
                        f")"
                    )

                    cell.font = bold

                    cell.fill = yellow_fill

                    cell.alignment = center

                    cell.border = border

                    cell.number_format = "[h]:mm"

                # ------------------------------------------------
                # Auto column width
                # ------------------------------------------------

                for column in ws.columns:

                    max_length = 0

                    column_letter = (
                        column[0].column_letter
                    )

                    for cell in column:

                        if cell.value is not None:

                            value = str(
                                cell.value
                            )

                            max_length = max(
                                max_length,
                                len(value)
                            )

                    ws.column_dimensions[
                        column_letter
                    ].width = min(
                        max_length + 2,
                        45
                    )

                # ------------------------------------------------
                # Freeze header
                # ------------------------------------------------

                ws.freeze_panes = "A2"

                # ------------------------------------------------
                # Filter
                # ------------------------------------------------

                ws.auto_filter.ref = ws.dimensions

            except Exception as employee_error:

                print(
                    f"Skipping employee {emp}: "
                    f"{employee_error}"
                )

        # ----------------------------------------------------
        # Save
        # ----------------------------------------------------

        writer.close()

        if employee_count == 0:

            raise ValueError(
                "No valid employee sheets could be created."
            )

        return output_file, employee_count

    except Exception:

        raise


# ============================================================
# ATTENDANCE FILE BROWSE
# ============================================================

def browse_attendance_file():

    file_path = filedialog.askopenfilename(

        title="Select Attendance Excel File",

        filetypes=[

            (
                "Excel Files",
                "*.xls *.xlsx"
            ),

            (
                "Excel 97-2003",
                "*.xls"
            ),

            (
                "Excel Workbook",
                "*.xlsx"
            ),

            (
                "All Files",
                "*.*"
            )

        ]

    )

    if file_path:

        attendance_file_var.set(
            file_path
        )

        attendance_status_var.set(
            "Attendance file selected."
        )


# ============================================================
# GENERATE ATTENDANCE
# ============================================================

def generate_attendance():

    input_file = (
        attendance_file_var.get()
        .strip()
    )

    if not input_file:

        messagebox.showwarning(
            "Select File",
            "Please select the biometric attendance Excel file."
        )

        return

    if not os.path.exists(input_file):

        messagebox.showerror(
            "File Not Found",
            "The selected Excel file does not exist."
        )

        return

    try:

        attendance_status_var.set(
            "Processing attendance..."
        )

        root.update_idletasks()

        output_file, employee_count = (
            process_attendance(
                input_file
            )
        )

        attendance_status_var.set(
            f"Completed - {employee_count} employee sheet(s) created."
        )

        result = messagebox.askyesno(
            "Attendance Completed",
            "Month_Timesheet.xlsx created successfully.\n\n"
            f"Location:\n{output_file}\n\n"
            "Would you like to open the file?"
        )

        if result:

            try:

                os.startfile(
                    output_file
                )

            except Exception:

                pass

    except Exception as error:

        attendance_status_var.set(
            "Processing failed."
        )

        messagebox.showerror(
            "Attendance Error",
            f"Could not generate the attendance file.\n\n"
            f"{error}"
        )


# ============================================================
# CLEAR ATTENDANCE FILE
# ============================================================

def clear_attendance_file():

    attendance_file_var.set("")

    attendance_status_var.set(
        "Ready"
    )


# ============================================================
# BACKUP CATEGORY FILTER
# ============================================================

def refresh_category_filter():

    if category_filter_var is None:
        return

    categories = get_categories()

    category_filter["values"] = [
        "All Categories"
    ] + categories

    if (
        category_filter_var.get()
        not in category_filter["values"]
    ):

        category_filter_var.set(
            "All Categories"
        )


# ============================================================
# REFRESH BACKUP TABLE
# ============================================================

def refresh_backup_table():

    if backup_tree is None:
        return

    for item in backup_tree.get_children():

        backup_tree.delete(
            item
        )

    selected_category = (
        category_filter_var.get()
    )

    search_text = (
        search_var.get()
        .lower()
        .strip()
    )

    count = 0

    for index, item in enumerate(
        backup_paths
    ):

        category = item.get(
            "category",
            "Other"
        )

        name = item.get(
            "name",
            ""
        )

        path = normalize_path(
            item.get(
                "path",
                ""
            )
        )

        if (
            selected_category
            != "All Categories"
            and category
            != selected_category
        ):

            continue

        if search_text:

            search_content = (
                category
                + " "
                + name
                + " "
                + path
            ).lower()

            if search_text not in search_content:

                continue

        backup_tree.insert(
            "",
            "end",
            iid=str(index),
            values=(
                category,
                name,
                path
            )
        )

        count += 1

    if backup_status_label:

        backup_status_label.config(
            text=(
                f"{count} path(s) shown  |  "
                f"{len(backup_paths)} total"
            )
        )


# ============================================================
# ADD BACKUP PATH
# ============================================================

def add_backup_path():

    add_window = tk.Toplevel(
        root
    )

    add_window.title(
        "Add Backup Path"
    )

    add_window.geometry(
        "680x360"
    )

    add_window.resizable(
        False,
        False
    )

    add_window.transient(
        root
    )

    add_window.grab_set()

    # --------------------------------------------------------
    # Category
    # --------------------------------------------------------

    tk.Label(
        add_window,
        text="Category:",
        font=("Segoe UI", 10)
    ).pack(
        anchor="w",
        padx=20,
        pady=(20, 5)
    )

    category_var = tk.StringVar()

    category_combo = ttk.Combobox(
        add_window,
        textvariable=category_var,
        values=get_categories(),
        state="normal",
        font=("Segoe UI", 10)
    )

    category_combo.pack(
        fill="x",
        padx=20
    )

    category_combo.set(
        "S3D"
    )

    # --------------------------------------------------------
    # Name
    # --------------------------------------------------------

    tk.Label(
        add_window,
        text="Backup Name:",
        font=("Segoe UI", 10)
    ).pack(
        anchor="w",
        padx=20,
        pady=(15, 5)
    )

    name_entry = tk.Entry(
        add_window,
        font=("Segoe UI", 10)
    )

    name_entry.pack(
        fill="x",
        padx=20
    )

    # --------------------------------------------------------
    # Path
    # --------------------------------------------------------

    tk.Label(
        add_window,
        text="Folder Path:",
        font=("Segoe UI", 10)
    ).pack(
        anchor="w",
        padx=20,
        pady=(15, 5)
    )

    path_frame = tk.Frame(
        add_window
    )

    path_frame.pack(
        fill="x",
        padx=20
    )

    path_entry = tk.Entry(
        path_frame,
        font=("Segoe UI", 10)
    )

    path_entry.pack(
        side="left",
        fill="x",
        expand=True
    )

    def browse_folder():

        folder = filedialog.askdirectory(
            title="Select Backup Folder"
        )

        if folder:

            path_entry.delete(
                0,
                tk.END
            )

            path_entry.insert(
                0,
                normalize_path(folder)
            )

    tk.Button(
        path_frame,
        text="Browse",
        width=10,
        command=browse_folder
    ).pack(
        side="left",
        padx=(10, 0)
    )

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    def save_new_path():

        category = (
            category_var.get()
            .strip()
        )

        name = (
            name_entry.get()
            .strip()
        )

        path = normalize_path(
            path_entry.get()
        )

        if not category:

            messagebox.showwarning(
                "Missing Category",
                "Please enter a category.",
                parent=add_window
            )

            return

        if not name:

            messagebox.showwarning(
                "Missing Name",
                "Please enter a backup name.",
                parent=add_window
            )

            return

        if not path:

            messagebox.showwarning(
                "Missing Path",
                "Please enter a folder path.",
                parent=add_window
            )

            return

        for item in backup_paths:

            if (
                item.get(
                    "category",
                    "Other"
                ).lower()
                == category.lower()
                and
                item.get(
                    "name",
                    ""
                ).lower()
                == name.lower()
            ):

                messagebox.showwarning(
                    "Duplicate Name",
                    "A backup with this name already exists "
                    "in this category.",
                    parent=add_window
                )

                return

        backup_paths.append(
            {
                "category": category,
                "name": name,
                "path": path
            }
        )

        save_backup_paths()

        refresh_category_filter()

        refresh_backup_table()

        add_window.destroy()

    # --------------------------------------------------------
    # Buttons
    # --------------------------------------------------------

    button_frame = tk.Frame(
        add_window
    )

    button_frame.pack(
        pady=25
    )

    tk.Button(
        button_frame,
        text="Save",
        width=12,
        command=save_new_path
    ).pack(
        side="left",
        padx=5
    )

    tk.Button(
        button_frame,
        text="Cancel",
        width=12,
        command=add_window.destroy
    ).pack(
        side="left",
        padx=5
    )

    name_entry.focus()


# ============================================================
# EDIT BACKUP PATH
# ============================================================

def edit_backup_path():

    selected = backup_tree.selection()

    if not selected:

        messagebox.showwarning(
            "Edit Path",
            "Please select a backup path."
        )

        return

    index = int(
        selected[0]
    )

    current = backup_paths[
        index
    ]

    edit_window = tk.Toplevel(
        root
    )

    edit_window.title(
        "Edit Backup Path"
    )

    edit_window.geometry(
        "680x360"
    )

    edit_window.resizable(
        False,
        False
    )

    edit_window.transient(
        root
    )

    edit_window.grab_set()

    # --------------------------------------------------------
    # Category
    # --------------------------------------------------------

    tk.Label(
        edit_window,
        text="Category:",
        font=("Segoe UI", 10)
    ).pack(
        anchor="w",
        padx=20,
        pady=(20, 5)
    )

    category_var = tk.StringVar()

    category_combo = ttk.Combobox(
        edit_window,
        textvariable=category_var,
        values=get_categories(),
        state="normal",
        font=("Segoe UI", 10)
    )

    category_combo.pack(
        fill="x",
        padx=20
    )

    category_combo.set(
        current.get(
            "category",
            "Other"
        )
    )

    # --------------------------------------------------------
    # Name
    # --------------------------------------------------------

    tk.Label(
        edit_window,
        text="Backup Name:",
        font=("Segoe UI", 10)
    ).pack(
        anchor="w",
        padx=20,
        pady=(15, 5)
    )

    name_entry = tk.Entry(
        edit_window,
        font=("Segoe UI", 10)
    )

    name_entry.pack(
        fill="x",
        padx=20
    )

    name_entry.insert(
        0,
        current.get(
            "name",
            ""
        )
    )

    # --------------------------------------------------------
    # Path
    # --------------------------------------------------------

    tk.Label(
        edit_window,
        text="Folder Path:",
        font=("Segoe UI", 10)
    ).pack(
        anchor="w",
        padx=20,
        pady=(15, 5)
    )

    path_frame = tk.Frame(
        edit_window
    )

    path_frame.pack(
        fill="x",
        padx=20
    )

    path_entry = tk.Entry(
        path_frame,
        font=("Segoe UI", 10)
    )

    path_entry.pack(
        side="left",
        fill="x",
        expand=True
    )

    path_entry.insert(
        0,
        normalize_path(
            current.get(
                "path",
                ""
            )
        )
    )

    def browse_folder():

        folder = filedialog.askdirectory(
            title="Select Backup Folder"
        )

        if folder:

            path_entry.delete(
                0,
                tk.END
            )

            path_entry.insert(
                0,
                normalize_path(folder)
            )

    tk.Button(
        path_frame,
        text="Browse",
        width=10,
        command=browse_folder
    ).pack(
        side="left",
        padx=(10, 0)
    )

    # --------------------------------------------------------
    # Update
    # --------------------------------------------------------

    def update_path():

        category = (
            category_var.get()
            .strip()
        )

        name = (
            name_entry.get()
            .strip()
        )

        path = normalize_path(
            path_entry.get()
        )

        if not category:

            messagebox.showwarning(
                "Missing Category",
                "Please enter a category.",
                parent=edit_window
            )

            return

        if not name:

            messagebox.showwarning(
                "Missing Name",
                "Please enter a backup name.",
                parent=edit_window
            )

            return

        if not path:

            messagebox.showwarning(
                "Missing Path",
                "Please enter a folder path.",
                parent=edit_window
            )

            return

        for i, item in enumerate(
            backup_paths
        ):

            if i == index:
                continue

            if (
                item.get(
                    "category",
                    "Other"
                ).lower()
                == category.lower()
                and
                item.get(
                    "name",
                    ""
                ).lower()
                == name.lower()
            ):

                messagebox.showwarning(
                    "Duplicate Name",
                    "A backup with this name already exists "
                    "in this category.",
                    parent=edit_window
                )

                return

        backup_paths[index] = {
            "category": category,
            "name": name,
            "path": path
        }

        save_backup_paths()

        refresh_category_filter()

        refresh_backup_table()

        edit_window.destroy()

    button_frame = tk.Frame(
        edit_window
    )

    button_frame.pack(
        pady=25
    )

    tk.Button(
        button_frame,
        text="Update",
        width=12,
        command=update_path
    ).pack(
        side="left",
        padx=5
    )

    tk.Button(
        button_frame,
        text="Cancel",
        width=12,
        command=edit_window.destroy
    ).pack(
        side="left",
        padx=5
    )


# ============================================================
# DELETE BACKUP PATH
# ============================================================

def delete_backup_path():

    selected = backup_tree.selection()

    if not selected:

        messagebox.showwarning(
            "Delete Path",
            "Please select a backup path."
        )

        return

    index = int(
        selected[0]
    )

    item = backup_paths[
        index
    ]

    category = item.get(
        "category",
        "Other"
    )

    name = item.get(
        "name",
        ""
    )

    path = normalize_path(
        item.get(
            "path",
            ""
        )
    )

    result = messagebox.askyesno(
        "Confirm Delete",
        "Are you sure you want to delete:\n\n"
        f"Category: {category}\n"
        f"Name: {name}\n\n"
        f"{path}"
    )

    if result:

        backup_paths.pop(
            index
        )

        save_backup_paths()

        refresh_category_filter()

        refresh_backup_table()


# ============================================================
# TEST SELECTED BACKUP
# ============================================================

def test_selected_backup():

    selected = backup_tree.selection()

    if not selected:

        messagebox.showwarning(
            "Test Path",
            "Please select a backup path."
        )

        return

    index = int(
        selected[0]
    )

    path = backup_paths[
        index
    ].get(
        "path",
        ""
    )

    test_path(path)


# ============================================================
# OPEN SELECTED BACKUP
# ============================================================

def open_selected_backup():

    selected = backup_tree.selection()

    if not selected:

        messagebox.showwarning(
            "Open Folder",
            "Please select a backup path."
        )

        return

    index = int(
        selected[0]
    )

    path = backup_paths[
        index
    ].get(
        "path",
        ""
    )

    open_folder(path)


# ============================================================
# BACKUP DOUBLE CLICK
# ============================================================

def backup_double_click(event):

    selected = backup_tree.selection()

    if not selected:
        return

    open_selected_backup()


# ============================================================
# CONTEXT MENU
# ============================================================

def show_backup_context_menu(event):

    row = backup_tree.identify_row(
        event.y
    )

    if not row:
        return

    backup_tree.selection_set(
        row
    )

    backup_context_menu.post(
        event.x_root,
        event.y_root
    )


# ============================================================
# SIDEBAR BUTTON STYLE
# ============================================================

def create_sidebar_button(
    parent,
    text,
    command
):

    button = tk.Button(

        parent,

        text=text,

        font=(
            "Segoe UI",
            10,
            "bold"
        ),

        bg=SIDEBAR_BUTTON,

        fg=TEXT_LIGHT,

        activebackground=SIDEBAR_ACTIVE,

        activeforeground=TEXT_LIGHT,

        relief="flat",

        bd=0,

        anchor="w",

        padx=20,

        cursor="hand2",

        command=command

    )

    button.pack(
        fill="x",
        padx=10,
        pady=3,
        ipady=10
    )

    return button


# ============================================================
# CLEAR MAIN CONTENT
# ============================================================

def clear_content():

    global current_page

    if current_page:

        current_page.destroy()

        current_page = None


# ============================================================
# PAGE TITLE
# ============================================================

def create_page_header(
    parent,
    title,
    subtitle
):

    frame = tk.Frame(
        parent,
        bg=MAIN_BG
    )

    frame.pack(
        fill="x",
        padx=30,
        pady=(25, 15)
    )

    tk.Label(
        frame,
        text=title,
        font=(
            "Segoe UI",
            22,
            "bold"
        ),
        fg=TEXT_DARK,
        bg=MAIN_BG
    ).pack(
        anchor="w"
    )

    tk.Label(
        frame,
        text=subtitle,
        font=(
            "Segoe UI",
            10
        ),
        fg=TEXT_GRAY,
        bg=MAIN_BG
    ).pack(
        anchor="w",
        pady=(5, 0)
    )

    return frame


# ============================================================
# DASHBOARD PAGE
# ============================================================

def show_dashboard():

    global current_page

    clear_content()

    current_page = tk.Frame(
        content_frame,
        bg=MAIN_BG
    )

    current_page.pack(
        fill="both",
        expand=True
    )

    create_page_header(
        current_page,
        "Dashboard",
        "Attendance Management Software"
    )

    # --------------------------------------------------------
    # Cards
    # --------------------------------------------------------

    cards_frame = tk.Frame(
        current_page,
        bg=MAIN_BG
    )

    cards_frame.pack(
        fill="x",
        padx=30,
        pady=15
    )

    # Attendance card
    attendance_card = tk.Frame(
        cards_frame,
        bg=CARD_BG,
        bd=1,
        relief="solid"
    )

    attendance_card.pack(
        side="left",
        fill="both",
        expand=True,
        padx=(0, 10),
        ipady=20
    )

    tk.Label(
        attendance_card,
        text="ATTENDANCE",
        font=(
            "Segoe UI",
            12,
            "bold"
        ),
        bg=CARD_BG,
        fg=TEXT_DARK
    ).pack(
        pady=(20, 5)
    )

    tk.Label(
        attendance_card,
        text="Process biometric\nattendance Excel files",
        font=(
            "Segoe UI",
            10
        ),
        bg=CARD_BG,
        fg=TEXT_GRAY,
        justify="center"
    ).pack(
        pady=10
    )

    tk.Button(
        attendance_card,
        text="Open Attendance",
        bg=SIDEBAR_ACTIVE,
        fg="white",
        relief="flat",
        cursor="hand2",
        command=show_attendance
    ).pack(
        pady=15,
        ipadx=15,
        ipady=6
    )

    # Backup card
    backup_card = tk.Frame(
        cards_frame,
        bg=CARD_BG,
        bd=1,
        relief="solid"
    )

    backup_card.pack(
        side="left",
        fill="both",
        expand=True,
        padx=(10, 0),
        ipady=20
    )

    tk.Label(
        backup_card,
        text="BACKUP PATHS",
        font=(
            "Segoe UI",
            12,
            "bold"
        ),
        bg=CARD_BG,
        fg=TEXT_DARK
    ).pack(
        pady=(20, 5)
    )

    tk.Label(
        backup_card,
        text="Manage S3D, PDS and\nother backup folders",
        font=(
            "Segoe UI",
            10
        ),
        bg=CARD_BG,
        fg=TEXT_GRAY,
        justify="center"
    ).pack(
        pady=10
    )

    tk.Button(
        backup_card,
        text="Open Backup Manager",
        bg=SIDEBAR_ACTIVE,
        fg="white",
        relief="flat",
        cursor="hand2",
        command=show_backup_manager
    ).pack(
        pady=15,
        ipadx=15,
        ipady=6
    )

    # --------------------------------------------------------
    # Information
    # --------------------------------------------------------

    info = tk.Frame(
        current_page,
        bg=CARD_BG,
        bd=1,
        relief="solid"
    )

    info.pack(
        fill="x",
        padx=30,
        pady=20
    )

    tk.Label(
        info,
        text="Quick Information",
        font=(
            "Segoe UI",
            13,
            "bold"
        ),
        bg=CARD_BG,
        fg=TEXT_DARK
    ).pack(
        anchor="w",
        padx=20,
        pady=(20, 10)
    )

    tk.Label(
        info,
        text=(
            "• Select Attendance from the sidebar to process your biometric Excel file.\n"
            "• Generated file name: Month_Timesheet.xlsx\n"
            "• Use Backup Manager to store and quickly access network backup folders.\n"
            "• Backup paths are stored locally for the current Windows user."
        ),
        font=(
            "Segoe UI",
            10
        ),
        bg=CARD_BG,
        fg=TEXT_GRAY,
        justify="left"
    ).pack(
        anchor="w",
        padx=20,
        pady=(0, 20)
    )


# ============================================================
# ATTENDANCE PAGE
# ============================================================

def show_attendance():

    global current_page

    clear_content()

    current_page = tk.Frame(
        content_frame,
        bg=MAIN_BG
    )

    current_page.pack(
        fill="both",
        expand=True
    )

    create_page_header(
        current_page,
        "Attendance Processing",
        "Generate the monthly attendance timesheet"
    )

    # --------------------------------------------------------
    # Main card
    # --------------------------------------------------------

    card = tk.Frame(
        current_page,
        bg=CARD_BG,
        bd=1,
        relief="solid"
    )

    card.pack(
        fill="x",
        padx=30,
        pady=10
    )

    tk.Label(
        card,
        text="Biometric Excel File",
        font=(
            "Segoe UI",
            11,
            "bold"
        ),
        bg=CARD_BG,
        fg=TEXT_DARK
    ).pack(
        anchor="w",
        padx=25,
        pady=(25, 8)
    )

    file_frame = tk.Frame(
        card,
        bg=CARD_BG
    )

    file_frame.pack(
        fill="x",
        padx=25
    )

    file_entry = tk.Entry(
        file_frame,
        textvariable=attendance_file_var,
        font=(
            "Segoe UI",
            10
        ),
        relief="solid",
        bd=1
    )

    file_entry.pack(
        side="left",
        fill="x",
        expand=True,
        ipady=7
    )

    tk.Button(
        file_frame,
        text="Browse",
        font=(
            "Segoe UI",
            10,
            "bold"
        ),
        bg=SIDEBAR_ACTIVE,
        fg="white",
        relief="flat",
        cursor="hand2",
        command=browse_attendance_file
    ).pack(
        side="left",
        padx=(10, 0),
        ipadx=15,
        ipady=5
    )

    # --------------------------------------------------------
    # Buttons
    # --------------------------------------------------------

    button_frame = tk.Frame(
        card,
        bg=CARD_BG
    )

    button_frame.pack(
        pady=25
    )

    tk.Button(
        button_frame,
        text="Generate Month Timesheet",
        font=(
            "Segoe UI",
            11,
            "bold"
        ),
        bg=SUCCESS,
        fg="white",
        relief="flat",
        cursor="hand2",
        command=generate_attendance
    ).pack(
        side="left",
        padx=5,
        ipadx=15,
        ipady=8
    )

    tk.Button(
        button_frame,
        text="Clear",
        font=(
            "Segoe UI",
            10
        ),
        bg="#E5E7EB",
        fg=TEXT_DARK,
        relief="flat",
        cursor="hand2",
        command=clear_attendance_file
    ).pack(
        side="left",
        padx=5,
        ipadx=15,
        ipady=8
    )

    # --------------------------------------------------------
    # Status
    # --------------------------------------------------------

    status_frame = tk.Frame(
        current_page,
        bg=CARD_BG,
        bd=1,
        relief="solid"
    )

    status_frame.pack(
        fill="x",
        padx=30,
        pady=10
    )

    tk.Label(
        status_frame,
        text="Status",
        font=(
            "Segoe UI",
            11,
            "bold"
        ),
        bg=CARD_BG,
        fg=TEXT_DARK
    ).pack(
        anchor="w",
        padx=25,
        pady=(20, 5)
    )

    tk.Label(
        status_frame,
        textvariable=attendance_status_var,
        font=(
            "Segoe UI",
            10
        ),
        bg=CARD_BG,
        fg=TEXT_GRAY
    ).pack(
        anchor="w",
        padx=25,
        pady=(0, 20)
    )

    # --------------------------------------------------------
    # Information
    # --------------------------------------------------------

    info_frame = tk.Frame(
        current_page,
        bg=CARD_BG,
        bd=1,
        relief="solid"
    )

    info_frame.pack(
        fill="x",
        padx=30,
        pady=10
    )

    tk.Label(
        info_frame,
        text="Output Information",
        font=(
            "Segoe UI",
            11,
            "bold"
        ),
        bg=CARD_BG,
        fg=TEXT_DARK
    ).pack(
        anchor="w",
        padx=25,
        pady=(20, 8)
    )

    tk.Label(
        info_frame,
        text=(
            "The generated workbook will be saved in the same folder "
            "as the selected biometric Excel file.\n\n"
            "Output file:\n"
            "Month_Timesheet.xlsx"
        ),
        font=(
            "Segoe UI",
            10
        ),
        bg=CARD_BG,
        fg=TEXT_GRAY,
        justify="left"
    ).pack(
        anchor="w",
        padx=25,
        pady=(0, 20)
    )


# ============================================================
# BACKUP MANAGER PAGE
# ============================================================

def show_backup_manager():

    global current_page
    global backup_tree
    global category_filter
    global backup_status_label

    clear_content()

    current_page = tk.Frame(
        content_frame,
        bg=MAIN_BG
    )

    current_page.pack(
        fill="both",
        expand=True
    )

    create_page_header(
        current_page,
        "Backup Path Manager",
        "Manage local and network backup folders"
    )

    # --------------------------------------------------------
    # Filter
    # --------------------------------------------------------

    filter_card = tk.Frame(
        current_page,
        bg=CARD_BG
    )

    filter_card.pack(
        fill="x",
        padx=30,
        pady=(0, 10)
    )

    tk.Label(
        filter_card,
        text="Category:",
        font=(
            "Segoe UI",
            10
        ),
        bg=CARD_BG,
        fg=TEXT_DARK
    ).pack(
        side="left",
        padx=(15, 5),
        pady=12
    )

    category_filter = ttk.Combobox(
        filter_card,
        textvariable=category_filter_var,
        state="readonly",
        width=20,
        font=(
            "Segoe UI",
            10
        )
    )

    category_filter.pack(
        side="left",
        padx=5,
        pady=10
    )

    category_filter.bind(
        "<<ComboboxSelected>>",
        lambda event: refresh_backup_table()
    )

    tk.Label(
        filter_card,
        text="Search:",
        font=(
            "Segoe UI",
            10
        ),
        bg=CARD_BG,
        fg=TEXT_DARK
    ).pack(
        side="left",
        padx=(20, 5)
    )

    search_entry = tk.Entry(
        filter_card,
        textvariable=search_var,
        font=(
            "Segoe UI",
            10
        )
    )

    search_entry.pack(
        side="left",
        fill="x",
        expand=True,
        padx=(5, 15),
        ipady=5
    )

    search_var.trace_add(
        "write",
        lambda *args: refresh_backup_table()
    )

    # --------------------------------------------------------
    # Table
    # --------------------------------------------------------

    table_card = tk.Frame(
        current_page,
        bg=CARD_BG
    )

    table_card.pack(
        fill="both",
        expand=True,
        padx=30,
        pady=5
    )

    columns = (
        "category",
        "name",
        "path"
    )

    backup_tree = ttk.Treeview(
        table_card,
        columns=columns,
        show="headings",
        selectmode="browse"
    )

    backup_tree.heading(
        "category",
        text="Category"
    )

    backup_tree.heading(
        "name",
        text="Backup Name"
    )

    backup_tree.heading(
        "path",
        text="Folder Path"
    )

    backup_tree.column(
        "category",
        width=130,
        anchor="w"
    )

    backup_tree.column(
        "name",
        width=220,
        anchor="w"
    )

    backup_tree.column(
        "path",
        width=650,
        anchor="w"
    )

    vertical_scrollbar = ttk.Scrollbar(
        table_card,
        orient="vertical",
        command=backup_tree.yview
    )

    horizontal_scrollbar = ttk.Scrollbar(
        table_card,
        orient="horizontal",
        command=backup_tree.xview
    )

    backup_tree.configure(
        yscrollcommand=vertical_scrollbar.set,
        xscrollcommand=horizontal_scrollbar.set
    )

    backup_tree.pack(
        side="left",
        fill="both",
        expand=True
    )

    vertical_scrollbar.pack(
        side="right",
        fill="y"
    )

    horizontal_scrollbar.pack(
        side="bottom",
        fill="x"
    )

    backup_tree.bind(
        "<Double-1>",
        backup_double_click
    )

    # --------------------------------------------------------
    # Context Menu
    # --------------------------------------------------------

    backup_context_menu.delete(
        0,
        tk.END
    )

    backup_context_menu.add_command(
        label="Open Folder",
        command=open_selected_backup
    )

    backup_context_menu.add_command(
        label="Test Path",
        command=test_selected_backup
    )

    backup_context_menu.add_separator()

    backup_context_menu.add_command(
        label="Edit",
        command=edit_backup_path
    )

    backup_context_menu.add_command(
        label="Delete",
        command=delete_backup_path
    )

    backup_tree.bind(
        "<Button-3>",
        show_backup_context_menu
    )

    # --------------------------------------------------------
    # Buttons
    # --------------------------------------------------------

    button_frame = tk.Frame(
        current_page,
        bg=MAIN_BG
    )

    button_frame.pack(
        fill="x",
        padx=30,
        pady=10
    )

    tk.Button(
        button_frame,
        text="+ Add",
        width=14,
        height=2,
        bg=SUCCESS,
        fg="white",
        relief="flat",
        cursor="hand2",
        command=add_backup_path
    ).pack(
        side="left",
        padx=4
    )

    tk.Button(
        button_frame,
        text="Edit",
        width=14,
        height=2,
        bg=SIDEBAR_ACTIVE,
        fg="white",
        relief="flat",
        cursor="hand2",
        command=edit_backup_path
    ).pack(
        side="left",
        padx=4
    )

    tk.Button(
        button_frame,
        text="Delete",
        width=14,
        height=2,
        bg=DANGER,
        fg="white",
        relief="flat",
        cursor="hand2",
        command=delete_backup_path
    ).pack(
        side="left",
        padx=4
    )

    tk.Button(
        button_frame,
        text="Test Path",
        width=14,
        height=2,
        bg=WARNING,
        fg="white",
        relief="flat",
        cursor="hand2",
        command=test_selected_backup
    ).pack(
        side="left",
        padx=4
    )

    tk.Button(
        button_frame,
        text="Open Folder",
        width=14,
        height=2,
        bg="#374151",
        fg="white",
        relief="flat",
        cursor="hand2",
        command=open_selected_backup
    ).pack(
        side="left",
        padx=4
    )

    # --------------------------------------------------------
    # Status
    # --------------------------------------------------------

    backup_status_label = tk.Label(
        current_page,
        text="Ready",
        anchor="w",
        bg=MAIN_BG,
        fg=TEXT_GRAY,
        font=(
            "Segoe UI",
            9
        )
    )

    backup_status_label.pack(
        fill="x",
        padx=30,
        pady=(0, 10)
    )

    refresh_category_filter()

    refresh_backup_table()


# ============================================================
# SETTINGS PAGE
# ============================================================

def show_settings():

    global current_page

    clear_content()

    current_page = tk.Frame(
        content_frame,
        bg=MAIN_BG
    )

    current_page.pack(
        fill="both",
        expand=True
    )

    create_page_header(
        current_page,
        "Settings",
        "Application information and configuration"
    )

    card = tk.Frame(
        current_page,
        bg=CARD_BG,
        bd=1,
        relief="solid"
    )

    card.pack(
        fill="x",
        padx=30,
        pady=10
    )

    tk.Label(
        card,
        text="Application",
        font=(
            "Segoe UI",
            12,
            "bold"
        ),
        bg=CARD_BG,
        fg=TEXT_DARK
    ).pack(
        anchor="w",
        padx=25,
        pady=(25, 10)
    )

    tk.Label(
        card,
        text=(
            "Attendance Management Software\n\n"
            "Version: 1.0\n\n"
            "Backup configuration is stored at:\n"
            f"{BACKUP_DATA_FILE}"
        ),
        font=(
            "Segoe UI",
            10
        ),
        bg=CARD_BG,
        fg=TEXT_GRAY,
        justify="left"
    ).pack(
        anchor="w",
        padx=25,
        pady=(0, 25)
    )


# ============================================================
# EXIT APPLICATION
# ============================================================

def exit_application():

    result = messagebox.askyesno(
        "Exit",
        "Are you sure you want to exit?"
    )

    if result:

        root.destroy()


# ============================================================
# SIDEBAR ACTIVE BUTTON
# ============================================================

sidebar_buttons = {}


def set_active_button(name):

    for button_name, button in sidebar_buttons.items():

        if button_name == name:

            button.config(
                bg=SIDEBAR_ACTIVE
            )

        else:

            button.config(
                bg=SIDEBAR_BUTTON
            )


def open_dashboard():

    set_active_button(
        "Dashboard"
    )

    show_dashboard()


def open_attendance():

    set_active_button(
        "Attendance"
    )

    show_attendance()


def open_backup():

    set_active_button(
        "Backup Manager"
    )

    show_backup_manager()


def open_settings():

    set_active_button(
        "Settings"
    )

    show_settings()


# ============================================================
# MAIN WINDOW
# ============================================================

root = tk.Tk()

root.title(
    APP_TITLE
)

root.geometry(
    "1250x750"
)

root.minsize(
    1000,
    650
)

root.configure(
    bg=MAIN_BG
)


# ============================================================
# APPLICATION ICON / LOGO
# ============================================================

if os.path.exists(LOGO_FILE):

    try:

        logo_image = tk.PhotoImage(
            file=LOGO_FILE
        )

        root.iconphoto(
            True,
            logo_image
        )

    except Exception as error:

        print(
            "Could not load logo:",
            error
        )


# ============================================================
# LOAD BACKUP DATA
# ============================================================

backup_paths = load_backup_paths()


# ============================================================
# VARIABLES
# ============================================================

attendance_file_var = tk.StringVar()

attendance_status_var = tk.StringVar(
    value="Ready"
)

category_filter_var = tk.StringVar(
    value="All Categories"
)

search_var = tk.StringVar()


# ============================================================
# SIDEBAR
# ============================================================

sidebar = tk.Frame(
    root,
    bg=SIDEBAR_BG,
    width=230
)

sidebar.pack(
    side="left",
    fill="y"
)

sidebar.pack_propagate(
    False
)


# ============================================================
# LOGO AREA
# ============================================================

logo_frame = tk.Frame(
    sidebar,
    bg=SIDEBAR_BG
)

logo_frame.pack(
    fill="x",
    pady=(25, 10)
)

if logo_image:

    # Resize is not supported directly by PhotoImage
    # so display original image if available.

    logo_label = tk.Label(
        logo_frame,
        image=logo_image,
        bg=SIDEBAR_BG
    )

    logo_label.pack(
        pady=5
    )

else:

    logo_label = tk.Label(
        logo_frame,
        text="D",
        font=(
            "Segoe UI",
            35,
            "bold"
        ),
        fg="white",
        bg=SIDEBAR_BG
    )

    logo_label.pack(
        pady=5
    )


tk.Label(
    logo_frame,
    text="ATTENDANCE\nMANAGEMENT",
    font=(
        "Segoe UI",
        11,
        "bold"
    ),
    fg="white",
    bg=SIDEBAR_BG,
    justify="center"
).pack(
    pady=(5, 20)
)


# ============================================================
# SIDEBAR MENU TITLE
# ============================================================

tk.Label(
    sidebar,
    text="MENU",
    font=(
        "Segoe UI",
        8,
        "bold"
    ),
    fg="#9CA3AF",
    bg=SIDEBAR_BG,
    anchor="w"
).pack(
    fill="x",
    padx=20,
    pady=(5, 5)
)


# ============================================================
# SIDEBAR BUTTONS
# ============================================================

sidebar_buttons["Dashboard"] = create_sidebar_button(
    sidebar,
    "  Dashboard",
    open_dashboard
)

sidebar_buttons["Attendance"] = create_sidebar_button(
    sidebar,
    "  Attendance",
    open_attendance
)

sidebar_buttons["Backup Manager"] = create_sidebar_button(
    sidebar,
    "  Backup Manager",
    open_backup
)

sidebar_buttons["Settings"] = create_sidebar_button(
    sidebar,
    "  Settings",
    open_settings
)


# ============================================================
# SIDEBAR BOTTOM
# ============================================================

sidebar_bottom = tk.Frame(
    sidebar,
    bg=SIDEBAR_BG
)

sidebar_bottom.pack(
    side="bottom",
    fill="x",
    pady=15
)

tk.Frame(
    sidebar_bottom,
    bg="#374151",
    height=1
).pack(
    fill="x",
    padx=15,
    pady=(0, 10)
)

exit_button = tk.Button(
    sidebar_bottom,
    text="  Exit",
    font=(
        "Segoe UI",
        10,
        "bold"
    ),
    bg=SIDEBAR_BUTTON,
    fg="white",
    activebackground=DANGER,
    activeforeground="white",
    relief="flat",
    bd=0,
    anchor="w",
    padx=20,
    cursor="hand2",
    command=exit_application
)

exit_button.pack(
    fill="x",
    padx=10,
    ipady=10
)


# ============================================================
# MAIN CONTENT
# ============================================================

content_frame = tk.Frame(
    root,
    bg=MAIN_BG
)

content_frame.pack(
    side="right",
    fill="both",
    expand=True
)


# ============================================================
# BACKUP CONTEXT MENU
# ============================================================

backup_context_menu = tk.Menu(
    root,
    tearoff=0
)


# ============================================================
# SHOW DASHBOARD FIRST
# ============================================================

set_active_button(
    "Dashboard"
)

show_dashboard()


# ============================================================
# START APPLICATION
# ============================================================

root.mainloop()