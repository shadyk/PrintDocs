import platform
import subprocess
import os

arabic_months = {
        "January": "كانون الثاني",
        "February": "شباط",
        "March": "آذار",
        "April": "نيسان",
        "May": "أيار",
        "June": "حزيران",
        "July": "تموز",
        "August": "آب",
        "September": "أيلول",
        "October": "تشرين الأول",
        "November": "تشرين الثاني",
        "December": "كانون الأول"
    }

def convert_to_eastern_arabic(number):
    eastern_arabic_digits = {
        "0": "٠", "1": "١", "2": "٢", "3": "٣", "4": "٤",
        "5": "٥", "6": "٦", "7": "٧", "8": "٨", "9": "٩"
    }
    return "".join(eastern_arabic_digits[digit] for digit in str(number))

def replace_text_in_paragraph(paragraph, old_text, new_text):
    """Replace text even if it is split across multiple runs."""
    full_text = "".join(run.text for run in paragraph.runs)  # Combine all runs
    if old_text in full_text:
        full_text = full_text.replace(old_text, new_text)  # Replace text
        for run in paragraph.runs:
            run.text = ""  # Clear all runs
        paragraph.add_run(full_text)  # Add updated text as a single run

def open_directory(path):
    """Open the directory containing the generated file."""
    if platform.system() == "Windows":
        os.startfile(path)  # Open directory on Windows
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", path])  # Open directory on macOS
    else:
        subprocess.Popen(["xdg-open", path])  # Open directory on Linux