import os
import time
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime  # Import datetime module for timestamp
import random
import yagooglesearch


def perform_google_search(keyword):
    try:

        client = yagooglesearch.SearchClient(
            keyword,
            tbs="li:1",
            max_search_result_urls_to_return=10,
            http_429_cool_off_time_in_minutes=45,
            http_429_cool_off_factor=1.5,
            verbosity=5,


            verbose_output=True,
        )
        client.assign_random_user_agent()

        results = client.search()
        # print(results)
        sleep_duration = random.uniform(12, 20)
        print(sleep_duration)
        time.sleep(sleep_duration)
        print(sleep_duration)
        return results
    except Exception as e:
        print("Error", f"Error performing Google search for '{keyword}': {e}")
        messagebox.showerror(
            "Error", f"Error performing Google search for '{keyword}': {e}"
        )
        return []


def extract_info(search_result):
    return search_result['url'], search_result['title'], search_result['description']


def write_to_excel(keyword, results, sheet, header_written=False):
    # Write headers to the sheet if not written before
    if header_written:
        sheet.append(["Keyword", "Rank", "Title", "URL", "Description"])

    # Write the keyword to each row
    for index, result in enumerate(results, start=1):
        url, title, description = extract_info(result)
        sheet.append([keyword, index, title, url, description])


def run_search(keywords, output_folder, wb):

    # Create a new sheet for each search
    sheet = wb.create_sheet(title="Results")
    header_written = True
    for i, keyword in enumerate(keywords):
        if i % 40 == 0 and i != 0:
            print("120")
            time.sleep(120)
            print("120")
        elif i % 20 == 0 and i != 0:
            print("90")
            time.sleep(90)
            print("90")
        elif i % 10 == 0 and i != 0:
            print("60")
            time.sleep(60)
            print("60")
        elif i % 5 == 0 and i != 0:
            print("30")
            time.sleep(30)
            print("30")

        print(f"Searching for: {keyword.strip()}")
        results = perform_google_search(keyword.strip())
        # print(results)
        # Pass header_written as an argument to write_to_excel
        write_to_excel(keyword, results, sheet, header_written)
        header_written = False

    # Use datetime to get current date and time for the file name
    current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")
    # output_path = filedialog.asksaveasfilename(initialdir=output_folder, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile=f"Google_{current_datetime}")
    output_filename = f"Google_{current_datetime}.xlsx"
    output_path = os.path.join(output_folder, output_filename)
    if output_path:
        try:
            wb.save(output_path)
            messagebox.showinfo(
                "Success", "Google search results saved successfully.")
        except Exception as e:
            print("Error", f"Error saving Excel file: {e}")
            messagebox.showerror("Error", f"Error saving Excel file: {e}")


def browse_and_set_variable(var, dialog_method):
    file_path = dialog_method()
    if file_path:
        var.set(file_path)


def create_ui():
    root = tk.Tk()
    root.title("Google Search Results")

    # Create a new Excel workbook
    wb = openpyxl.Workbook()

    file_path_var = tk.StringVar()
    output_folder_var = tk.StringVar()

    label_file = tk.Label(root, text="Select Keywords File:")
    label_file.pack()

    entry_file = tk.Entry(root, textvariable=file_path_var)
    entry_file.pack()

    label_output = tk.Label(root, text="Select Output Folder:")
    label_output.pack()

    entry_output = tk.Entry(root, textvariable=output_folder_var)
    entry_output.pack()

    button_browse = tk.Button(
        root,
        text="Browse Keywords File",
        command=lambda: browse_and_set_variable(
            file_path_var, filedialog.askopenfilename
        ),
    )
    button_browse.pack()

    button_browse_output = tk.Button(
        root,
        text="Browse Output Folder",
        command=lambda: browse_and_set_variable(
            output_folder_var, filedialog.askdirectory
        ),
    )
    button_browse_output.pack()

    button_run = tk.Button(
        root,
        text="Run",
        command=lambda: run_search(
            get_keywords(file_path_var.get()), output_folder_var.get(), wb
        ),
    )
    button_run.pack()

    root.mainloop()


def get_keywords(file_path):
    try:
        with open(file_path, "r") as file:
            return [line.strip() for line in file]
    except Exception as e:
        print("Error", f"Error reading keywords file: {e}")
        messagebox.showerror("Error", f"Error reading keywords file: {e}")
        return []


if __name__ == "__main__":
    create_ui()
