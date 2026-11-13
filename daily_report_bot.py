
import os
import re
import time
import subprocess
from datetime import datetime

import pyautogui
import pyperclip
import pygetwindow as gw

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ============================================================
# CONFIGURATION
# ============================================================

OUTPUT_DIR = (
    r"C:\Users\vishpra2\OneDrive - Magna\Vishnu\Gen AI Assignments"
    r"\Week 1\PyAutoGUI_Automation_Assignment\PyAutoGUI Automation"
)

CRICBUZZ_URL = "https://www.cricbuzz.com/cricket-match/live-scores"

TODAY = datetime.now().strftime("%Y-%m-%d")

REPORT_PATH = os.path.join(
    OUTPUT_DIR, f"daily_report_{TODAY}.xlsx"
)

CRICBUZZ_SCREENSHOT = os.path.join(
    OUTPUT_DIR, f"cricbuzz_screenshot_{TODAY}.png"
)

EXCEL_SCREENSHOT = os.path.join(
    OUTPUT_DIR, f"excel_screenshot_{TODAY}.png"
)

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5


# ============================================================
# WINDOW MAXIMIZATION
# ============================================================

def maximize_window(title_keyword, timeout=20):

    deadline = time.time() + timeout

    while time.time() < deadline:

        windows = [
            window for window in gw.getAllWindows()
            if title_keyword.lower() in
            (window.title or "").lower()
        ]

        for window in windows:

            try:

                if window.isMinimized:
                    window.restore()
                    time.sleep(1)

                window.activate()
                time.sleep(1)

                window.maximize()
                time.sleep(2)

                return True

            except Exception:
                continue

        time.sleep(1)

    return False


# ============================================================
# STEP 1: OPEN CRICBUZZ
# ============================================================

def open_cricbuzz():

    print("\nSTEP 1: Opening Cricbuzz...")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    subprocess.Popen(
        f'cmd /c start "" chrome --new-window "{CRICBUZZ_URL}"',
        shell=True
    )

    print("Waiting for Chrome...")
    time.sleep(10)

    if not maximize_window("Cricbuzz", 15):
        maximize_window("Chrome", 10)

    pyautogui.hotkey("ctrl", "l")

    pyperclip.copy(CRICBUZZ_URL)

    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")

    print("Loading Cricbuzz Live Scores...")

    time.sleep(22)

    maximize_window("Cricbuzz", 15)

    pyautogui.press("home")

    time.sleep(3)

    print("Cricbuzz opened successfully!")


# ============================================================
# STEP 2: CAPTURE CRICBUZZ SCREENSHOT
# ============================================================

def capture_cricbuzz():

    print("\nSTEP 2: Capturing Cricbuzz screenshot...")

    maximize_window("Cricbuzz", 15)

    time.sleep(3)

    screenshot = pyautogui.screenshot()

    screenshot.save(CRICBUZZ_SCREENSHOT)

    print("Cricbuzz screenshot saved!")

    print(CRICBUZZ_SCREENSHOT)


# ============================================================
# STEP 3: EXTRACT CRICKET INFORMATION
# ============================================================

def extract_cricket_data():

    print("\nSTEP 3: Extracting cricket information...")

    if not maximize_window("Cricbuzz", 15):

        raise RuntimeError(
            "Unable to activate the Cricbuzz browser window."
        )

    time.sleep(2)

    pyautogui.press("esc")

    pyautogui.hotkey("ctrl", "a")

    time.sleep(1)

    pyautogui.hotkey("ctrl", "c")

    time.sleep(2)

    copied_text = pyperclip.paste()

    print(
        "Clipboard characters copied:",
        len(copied_text)
    )

    if not copied_text or len(copied_text.strip()) < 20:

        raise RuntimeError(
            "Unable to copy Cricbuzz page information."
        )

    raw_lines = copied_text.replace("\r", "\n").split("\n")

    lines = []
    seen = set()

    for raw_line in raw_lines:

        line = re.sub(r"\s+", " ", raw_line).strip()

        if not line:
            continue

        key = line.casefold()

        if key not in seen:

            seen.add(key)
            lines.append(line)

    keywords = (
        " vs ",
        " v ",
        "live",
        "won by",
        "need ",
        "required",
        "target",
        "innings",
        "overs",
        "wicket",
        "all out",
        "stumps",
        "drawn",
        "tied",
        "abandoned",
        "no result",
        "scheduled",
        "run rate",
        "crr",
        "rr"
    )

    score_pattern = re.compile(
        r"\b\d{1,3}\s*/\s*\d{1,2}\b"
        r"|\b\d{1,3}\s*-\s*\d{1,2}\b"
    )

    useful_lines = []

    for line in lines:

        lower_line = line.casefold()

        if (
            score_pattern.search(line)
            or any(word in lower_line for word in keywords)
        ):

            if line not in useful_lines:
                useful_lines.append(line)

    # Fallback if no score or match keywords are detected.

    if not useful_lines:

        useful_lines = [
            line for line in lines
            if len(line) >= 8
            and not any(
                term in line.casefold()
                for term in (
                    "go premium",
                    "privacy policy",
                    "terms of use"
                )
            )
        ][:30]

    cricket_information = " | ".join(
        useful_lines[:30]
    )

    cricket_information = re.sub(
        r"[\t\r\n]+", " ", cricket_information
    )

    cricket_information = re.sub(
        r"\s+", " ", cricket_information
    ).strip()

    cricket_information = cricket_information[:5000]

    if not cricket_information:

        cricket_information = (
            "Cricbuzz page opened, but no match information "
            "was available to copy."
        )

    print("\nExtracted cricket information:")

    print(cricket_information[:2000])

    return cricket_information


# ============================================================
# STEP 4: CREATE EXCEL WORKBOOK AND INSERT DATA
# ============================================================

def create_excel_report(cricket_information):

    print("\nSTEP 4: Creating Excel workbook...")

    current_time = datetime.now()

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Cricbuzz Daily Report"

    # Excel headers

    headers = [
        "Date and Time",
        "Fetched Cricket Information",
        "Comment"
    ]

    sheet.append(headers)

    # Excel data

    data = [
        current_time.strftime("%d-%m-%Y %H:%M:%S"),
        cricket_information,
        "Review the latest match updates."
    ]

    sheet.append(data)

    # Header formatting

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="008F72"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True,
        size=11
    )

    for cell in sheet[1]:

        cell.fill = header_fill
        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    # Data formatting

    for cell in sheet[2]:

        cell.alignment = Alignment(
            vertical="top",
            wrap_text=True
        )

    # Column widths

    sheet.column_dimensions["A"].width = 23
    sheet.column_dimensions["B"].width = 100
    sheet.column_dimensions["C"].width = 38

    sheet.row_dimensions[1].height = 25
    sheet.row_dimensions[2].height = 110

    # Freeze header row

    sheet.freeze_panes = "A2"

    # Enable filters

    sheet.auto_filter.ref = "A1:C2"

    # Cell borders

    thin_border = Side(
        style="thin",
        color="D9E1F2"
    )

    for row in sheet.iter_rows(
        min_row=1,
        max_row=2,
        min_col=1,
        max_col=3
    ):

        for cell in row:

            cell.border = Border(
                left=thin_border,
                right=thin_border,
                top=thin_border,
                bottom=thin_border
            )

    # Save workbook

    workbook.save(REPORT_PATH)

    # Verify saved Excel workbook

    if (
        not os.path.isfile(REPORT_PATH)
        or os.path.getsize(REPORT_PATH) == 0
    ):

        raise RuntimeError(
            "Excel workbook was not created properly."
        )

    check_workbook = load_workbook(
        REPORT_PATH,
        read_only=True,
        data_only=True
    )

    check_sheet = check_workbook["Cricbuzz Daily Report"]

    if (
        check_sheet["A2"].value is None
        or check_sheet["B2"].value is None
    ):

        check_workbook.close()

        raise RuntimeError(
            "Excel workbook was saved, but data is missing."
        )

    check_workbook.close()

    print("Excel workbook created successfully!")

    print("Data inserted successfully!")

    print("Excel file:", REPORT_PATH)


# ============================================================
# STEP 5: OPEN COMPLETED EXCEL WORKBOOK
# ============================================================

def open_excel_report():

    print("\nSTEP 5: Opening Excel report...")

    if not os.path.isfile(REPORT_PATH):

        raise FileNotFoundError(
            "Excel report does not exist."
        )

    # Open the completed Excel workbook.

    os.startfile(REPORT_PATH)

    print("Waiting for Microsoft Excel...")

    time.sleep(12)

    if maximize_window("Excel", 15):

        print("Excel workbook opened successfully!")

    else:

        print(
            "Workbook was sent to the default application. "
            "Could not detect an Excel window to maximize."
        )


# ============================================================
# STEP 6: CAPTURE EXCEL SCREENSHOT
# ============================================================

def capture_excel():

    print("\nSTEP 6: Capturing Excel screenshot...")

    maximize_window("Excel", 10)

    time.sleep(3)

    screenshot = pyautogui.screenshot()

    screenshot.save(EXCEL_SCREENSHOT)

    print("Excel screenshot saved!")

    print(EXCEL_SCREENSHOT)


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():

    print("=" * 60)

    print("CRICBUZZ DAILY REPORT AUTOMATION")

    print("=" * 60)

    try:

        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # Open Cricbuzz

        open_cricbuzz()

        # Capture Cricbuzz screenshot

        capture_cricbuzz()

        # Extract cricket information

        cricket_information = extract_cricket_data()

        # Create Excel and insert data

        create_excel_report(cricket_information)

        # Open completed workbook

        open_excel_report()

        # Capture Excel screenshot

        capture_excel()

        print("\n" + "=" * 60)

        print("AUTOMATION COMPLETED SUCCESSFULLY!")

        print("=" * 60)

        print("\nGenerated files:")

        print("1. Excel report:", REPORT_PATH)

        print("2. Cricbuzz screenshot:", CRICBUZZ_SCREENSHOT)

        print("3. Excel screenshot:", EXCEL_SCREENSHOT)

    except Exception as error:

        print("\nAUTOMATION FAILED!")

        print("Error:", repr(error))

        print("\nPlease check the terminal output.")


if __name__ == "__main__":

    main()
