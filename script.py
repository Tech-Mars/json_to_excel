import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pyperclip
import io


def convert_json_to_excel():
    # Get JSON data from clipboard
    try:
        json_data = pyperclip.paste()
    except pyperclip.PyperclipException:
        messagebox.showerror("Error", "Failed to access clipboard.")
        return

    # Check if JSON data is empty
    if not json_data:
        messagebox.showerror("Error", "Clipboard does not contain JSON data.")
        return

    try:
        # Load JSON data into a DataFrame
        with io.StringIO(json_data) as json_buffer:
            df = pd.read_json(json_buffer)

        # Ask user to choose file name and location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not file_path:
            return  # User canceled saving

        # Export DataFrame to Excel file
        df.to_excel(file_path, index=False)

        messagebox.showinfo("Success", f"Excel file saved successfully: {file_path}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


# GUI setup
root = tk.Tk()
root.title("JSON to Excel Converter")

# Main frame
main_frame = tk.Frame(root, padx=10, pady=10)
main_frame.pack()

# Instruction label
label = tk.Label(
    main_frame,
    text="Copy JSON data to clipboard and click 'Convert' to save as Excel file.",
)
label.pack()

# Convert button
convert_button = tk.Button(main_frame, text="Convert", command=convert_json_to_excel)
convert_button.pack()

root.mainloop()
