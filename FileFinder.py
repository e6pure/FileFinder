import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import os
import shutil
import threading
from openpyxl import Workbook

# --- Configuration Constants ---
BG_COLOR_SCAN = "#4CAF50"
BG_COLOR_ACTION = "#f44336"
TEXT_COLOR_LIGHT = "white"
FONT_CONFIG = ('Helvetica', 10, 'bold')
WARN_COLOR_DELETE = "red"

# Global Variables
file_list = []

# --- UTILITY FUNCTIONS ---

def center_toplevel_window(dialog, width, height):
    """Centers a Toplevel window relative to the root window."""
    dialog.update_idletasks()
    root_x, root_y = root.winfo_x(), root.winfo_y()
    root_width, root_height = root.winfo_width(), root.winfo_height()
    position_x = root_x + (root_width // 2) - (width // 2)
    position_y = root_y + (root_height // 2) - (height // 2)
    dialog.geometry(f"{width}x{height}+{int(position_x)}+{int(position_y)}")

def show_custom_message(title, message):
    """Displays a custom centered message dialog."""
    dialog = tk.Toplevel(root)
    dialog.title(title)
    dialog.grab_set()
    dialog.resizable(False, False)

    dialog_width = 350
    num_lines = message.count('\n') + 1
    dialog_height = 80 + (num_lines * 15)
    center_toplevel_window(dialog, dialog_width, dialog_height)
    
    content_frame = tk.Frame(dialog)
    content_frame.pack(expand=True, fill='both', padx=20, pady=15)
    
    button_frame = tk.Frame(dialog)
    button_frame.pack(side='bottom', pady=(0, 10))
    
    msg_label = tk.Label(content_frame, text=message, justify="left", font=('Helvetica', 10), wraplength=dialog_width - 40)
    msg_label.pack()
    
    ok_button = tk.Button(button_frame, text="OK", width=10, command=dialog.destroy)
    ok_button.pack()
    ok_button.focus_set()
    
    root.wait_window(dialog)

def get_unique_filename(directory, filename):
    """Checks if a file exists and returns a unique name like name_1.ext."""
    base_name, extension = os.path.splitext(filename)
    counter = 1
    new_filename = filename
    while os.path.exists(os.path.join(directory, new_filename)):
        new_filename = f"{base_name}_{counter}{extension}"
        counter += 1
    return new_filename

# --- UI HANDLERS ---

def select_directory(label_widget):
    directory = filedialog.askdirectory(parent=root)
    if directory:
        label_widget.config(text=directory)

def toggle_custom_entry_state():
    custom_ext_entry.config(state='normal' if custom_ext_var.get() else 'disabled')

def run_scan():
    global file_list
    selected_dir = directory_label.cget("text")
    if not os.path.isdir(selected_dir):
        show_custom_message("Error", "Please select a valid directory to scan.")
        return

    file_list.clear()
    found_counts = {}
    selected_extensions = []

    # Map variables to extensions
    if png_var.get(): selected_extensions.append(".png")
    if jpg_var.get(): selected_extensions.append(".jpg")
    if jpeg_var.get(): selected_extensions.append(".jpeg")
    if webp_var.get(): selected_extensions.append(".webp")
    if rar_var.get(): selected_extensions.append(".rar")
    if zip_var.get(): selected_extensions.append(".zip")
    if seven_zip_var.get(): selected_extensions.append(".7z")
    if tar_var.get(): selected_extensions.append(".tar")
    if mp4_var.get(): selected_extensions.append(".mp4")
    if mov_var.get(): selected_extensions.append(".mov")
    if avi_var.get(): selected_extensions.append(".avi")
    if mkv_var.get(): selected_extensions.append(".mkv")
    if mp3_var.get(): selected_extensions.append(".mp3")
    if wav_var.get(): selected_extensions.append(".wav")
    if flac_var.get(): selected_extensions.append(".flac")
    if aac_var.get(): selected_extensions.append(".aac")
    if word_var.get(): selected_extensions.extend([".doc", ".docx"])
    if excel_var.get(): selected_extensions.extend([".xls", ".xlsx"])
    if pdf_var.get(): selected_extensions.append(".pdf")
    if psd_var.get(): selected_extensions.append(".psd")

    if custom_ext_var.get():
        custom_input = custom_ext_entry_var.get().strip()
        if custom_input:
            custom_list = [ext.strip() for ext in custom_input.split(',')]
            for ext in custom_list:
                if ext:
                    if not ext.startswith('.'): ext = '.' + ext
                    if ext not in selected_extensions: selected_extensions.append(ext)

    if not selected_extensions:
        show_custom_message("Error", "Please select at least one file type.")
        return

    display_map = {
        ".png": "PNG", ".jpg": "JPG", ".jpeg": "JPEG", ".webp": "WEBP",
        ".rar": "RAR", ".zip": "ZIP", ".7z": "7Z", ".tar": "TAR",
        ".mp4": "MP4", ".mov": "MOV", ".avi": "AVI", ".mkv": "MKV",
        ".mp3": "MP3", ".wav": "WAV", ".flac": "FLAC", ".aac": "AAC",
        ".doc": "Word", ".docx": "Word", ".xls": "Excel", ".xlsx": "Excel",
        ".pdf": "PDF", ".psd": "PSD"
    }

    for root_dir, dirs, files in os.walk(selected_dir):
        for file in files:
            file_lower = file.lower()
            for ext in selected_extensions:
                if file_lower.endswith(ext):
                    file_list.append(os.path.join(root_dir, file))
                    display_name = display_map.get(ext, ext.upper())
                    found_counts[display_name] = found_counts.get(display_name, 0) + 1
                    break

    if not file_list:
        show_custom_message("Notification", "No files found matching your criteria.")
        return

    result_lines = [f"Found {len(file_list)} files:", ""]
    for name, count in sorted(found_counts.items()):
        result_lines.append(f"{name}: {count} files")
    
    final_message = "\n".join(result_lines)
    
    result_dialog = tk.Toplevel(root)
    result_dialog.title("Scan Complete")
    result_dialog.grab_set()
    result_dialog.resizable(False, False)
    
    dialog_width = 300
    dialog_height = 80 + (len(found_counts) * 22) + 40
    center_toplevel_window(result_dialog, dialog_width, dialog_height)
    
    content_frame = tk.Frame(result_dialog)
    content_frame.pack(expand=True, fill='both', padx=20, pady=15)
    
    msg_label = tk.Label(content_frame, text=final_message, justify="left", font=('Helvetica', 10))
    msg_label.pack()
    
    ok_button = tk.Button(result_dialog, text="OK", width=10, command=result_dialog.destroy)
    ok_button.pack(side='bottom', pady=10)
    ok_button.focus_set()

def perform_actions():
    if not file_list:
        show_custom_message("Error", "Please 'Run (Scan)' to find files first.")
        return

    # Check which actions are selected
    do_copy = copy_var.get()
    do_delete = delete_var.get()
    do_excel = create_excel_var.get() 

    if not (do_copy or do_delete or do_excel):
        show_custom_message("Error", "Please select at least one action.")
        return

    dest_dir_for_copy = ""
    if do_copy:
        dest_dir_for_copy = filedialog.askdirectory(title="Select Destination Directory", parent=root)
        if not dest_dir_for_copy:
            show_custom_message("Notification", "Copy action cancelled.")
            return

    delete_confirmed = False
    if do_delete:
        delete_confirmed = messagebox.askyesno(
            "DELETE CONFIRMATION",
            "Are you sure you want to PERMANENTLY delete selected files?\nThis action cannot be undone.",
            icon='warning', parent=root
        )
        if not delete_confirmed:
            show_custom_message("Notification", "Delete action cancelled.")
    
    excel_save_path = ""
    if do_excel:
        excel_save_path = filedialog.asksaveasfilename(
            parent=root, defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")], title="Save Excel File"
        )
        if not excel_save_path:
             show_custom_message("Notification", "Excel creation cancelled.")
             # If Excel was the only action and cancelled, stop.
             if not (do_copy or (do_delete and delete_confirmed)):
                 return
             do_excel = False # Disable excel action for the thread

    # If we have valid actions, proceed to setup thread
    if (do_copy and dest_dir_for_copy) or (do_delete and delete_confirmed) or (do_excel and excel_save_path):
        start_processing_thread(do_copy, dest_dir_for_copy, do_delete, delete_confirmed, do_excel, excel_save_path)

def start_processing_thread(do_copy, dest_dir, do_delete, delete_confirmed, do_excel, excel_path):
    # Create Progress Dialog
    progress_dialog = tk.Toplevel(root)
    progress_dialog.title("Processing...")
    progress_dialog.grab_set()
    progress_dialog.resizable(False, False)
    
    center_toplevel_window(progress_dialog, 450, 180)
    
    status_label = tk.Label(progress_dialog, text="Preparing...", font=('Helvetica', 10), wraplength=430)
    status_label.pack(pady=(20, 5))
    
    progress_bar = ttk.Progressbar(progress_dialog, orient='horizontal', length=400, mode='determinate')
    progress_bar.pack(pady=10)
    
    count_label = tk.Label(progress_dialog, text="", font=('Helvetica', 9))
    count_label.pack()

    ok_button = tk.Button(progress_dialog, text="OK", command=progress_dialog.destroy, width=10, state='disabled')
    ok_button.pack(pady=20)

    # Define the worker function
    def worker():
        results_summary = []
        
        # --- COPY ACTION ---
        if do_copy and dest_dir:
            try:
                total_files = len(file_list)
                copied_count = 0
                errors = 0
                
                # Update UI thread-safe
                root.after(0, lambda: status_label.config(text="Copying files..."))
                root.after(0, lambda: progress_bar.config(maximum=total_files, value=0))

                for i, file_path in enumerate(file_list):
                    try:
                        file_extension = os.path.splitext(file_path)[1]
                        subfolder_name = file_extension[1:].upper() if file_extension else 'NO_EXTENSION'
                        target_subfolder = os.path.join(dest_dir, subfolder_name)
                        os.makedirs(target_subfolder, exist_ok=True)
                        
                        file_name = os.path.basename(file_path)
                        # Handle duplicate filenames
                        unique_name = get_unique_filename(target_subfolder, file_name)
                        shutil.copy(file_path, os.path.join(target_subfolder, unique_name))
                        copied_count += 1
                    except Exception as e:
                        print(f"Copy Error: {e}")
                        errors += 1
                    
                    # Update progress
                    root.after(0, lambda v=i+1: progress_bar.config(value=v))
                    root.after(0, lambda v=i+1, t=total_files: count_label.config(text=f"{v} / {t}"))

                msg = f"Successfully copied {copied_count} files."
                if errors > 0: msg += f" (Failed: {errors})"
                results_summary.append(msg)
            except Exception as e:
                results_summary.append(f"Copy Process Error: {e}")

        # --- DELETE ACTION ---
        if do_delete and delete_confirmed:
            try:
                # Create a copy of list to avoid index issues while removing
                list_to_delete = list(file_list) 
                total_files = len(list_to_delete)
                deleted_count = 0
                errors = 0

                root.after(0, lambda: status_label.config(text="Deleting files..."))
                root.after(0, lambda: progress_bar.config(maximum=total_files, value=0))

                for i, file_path in enumerate(list_to_delete):
                    try:
                        os.remove(file_path)
                        if file_path in file_list:
                            file_list.remove(file_path)
                        deleted_count += 1
                    except Exception as e:
                        print(f"Delete Error: {e}")
                        errors += 1
                    
                    root.after(0, lambda v=i+1: progress_bar.config(value=v))
                    root.after(0, lambda v=i+1, t=total_files: count_label.config(text=f"{v} / {t}"))

                msg = f"Successfully deleted {deleted_count} files."
                if errors > 0: msg += f" (Failed: {errors})"
                results_summary.append(msg)
            except Exception as e:
                results_summary.append(f"Delete Process Error: {e}")

        # --- EXCEL ACTION ---
        if do_excel and excel_path:
            root.after(0, lambda: status_label.config(text="Generating Excel file..."))
            root.after(0, lambda: progress_bar.config(value=0, maximum=100)) # Indeterminate state visually
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "File List"
                ws.append(["No.", "File Name", "Full Path"])
                for i, file_path in enumerate(file_list, 1):
                    ws.append([i, os.path.basename(file_path), file_path])
                wb.save(excel_path)
                results_summary.append("Successfully saved Excel file.")
            except Exception as e:
                results_summary.append(f"Excel Creation Error: {e}")
            root.after(0, lambda: progress_bar.config(value=100))

        # --- FINISH ---
        final_message = "\n".join(results_summary) if results_summary else "No valid actions performed."
        
        def finish_ui():
            progress_dialog.title("Finished!")
            status_label.config(text=final_message)
            progress_bar.pack_forget()
            count_label.pack_forget()
            ok_button.config(state='normal')
            ok_button.focus_set()

        root.after(0, finish_ui)

    # Start the thread
    threading.Thread(target=worker, daemon=True).start()

# --- GUI SETUP ---
root = tk.Tk()
root.title("FileFinder1.1")
root.geometry("700x560")

main_frame = tk.Frame(root, padx=10, pady=10)
main_frame.pack(fill="both", expand=True)

# Section 1: Source
source_frame = tk.LabelFrame(main_frame, text="1. Select Source Directory")
source_frame.pack(pady=5, padx=5, fill="x")

directory_label = tk.Label(source_frame, text="No directory selected", wraplength=680, justify="left")
directory_label.pack(pady=5, padx=5)

select_button = tk.Button(source_frame, text="Browse...", command=lambda: select_directory(directory_label))
select_button.pack(pady=5)

# Section 2: File Types
file_type_frame = tk.LabelFrame(main_frame, text="2. Select File Types to Scan")
file_type_frame.pack(pady=10, padx=5, fill="x")

# Variables for File Types
png_var, jpg_var, jpeg_var, webp_var = tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar()
rar_var, zip_var, seven_zip_var, tar_var = tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar()
mp4_var, mov_var, avi_var, mkv_var = tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar()
mp3_var, wav_var, flac_var, aac_var = tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar()
word_var, excel_var, pdf_var, psd_var = tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar()

columns_container = tk.Frame(file_type_frame)
columns_container.pack(pady=5)

image_col = tk.Frame(columns_container)
video_col = tk.Frame(columns_container)
archive_col = tk.Frame(columns_container)
audio_col = tk.Frame(columns_container)
doc_col = tk.Frame(columns_container)

for col in (image_col, video_col, archive_col, audio_col, doc_col):
    col.pack(side="left", padx=12, anchor="n")

# Layout Checkboxes
tk.Label(image_col, text="Images", font=FONT_CONFIG).pack(anchor='center')
tk.Checkbutton(image_col, text="PNG (.png)", variable=png_var).pack(anchor='w')
tk.Checkbutton(image_col, text="JPG (.jpg)", variable=jpg_var).pack(anchor='w')
tk.Checkbutton(image_col, text="JPEG (.jpeg)", variable=jpeg_var).pack(anchor='w')
tk.Checkbutton(image_col, text="WebP (.webp)", variable=webp_var).pack(anchor='w')

tk.Label(video_col, text="Video", font=FONT_CONFIG).pack(anchor='center')
tk.Checkbutton(video_col, text="MP4 (.mp4)", variable=mp4_var).pack(anchor='w')
tk.Checkbutton(video_col, text="MOV (.mov)", variable=mov_var).pack(anchor='w')
tk.Checkbutton(video_col, text="AVI (.avi)", variable=avi_var).pack(anchor='w')
tk.Checkbutton(video_col, text="MKV (.mkv)", variable=mkv_var).pack(anchor='w')

tk.Label(archive_col, text="Archive", font=FONT_CONFIG).pack(anchor='center')
tk.Checkbutton(archive_col, text="RAR (.rar)", variable=rar_var).pack(anchor='w')
tk.Checkbutton(archive_col, text="ZIP (.zip)", variable=zip_var).pack(anchor='w')
tk.Checkbutton(archive_col, text="7z (.7z)", variable=seven_zip_var).pack(anchor='w')
tk.Checkbutton(archive_col, text="TAR (.tar)", variable=tar_var).pack(anchor='w')

tk.Label(audio_col, text="Audio", font=FONT_CONFIG).pack(anchor='center')
tk.Checkbutton(audio_col, text="MP3 (.mp3)", variable=mp3_var).pack(anchor='w')
tk.Checkbutton(audio_col, text="WAV (.wav)", variable=wav_var).pack(anchor='w')
tk.Checkbutton(audio_col, text="FLAC (.flac)", variable=flac_var).pack(anchor='w')
tk.Checkbutton(audio_col, text="AAC (.aac)", variable=aac_var).pack(anchor='w')

tk.Label(doc_col, text="Docs", font=FONT_CONFIG).pack(anchor='center')
tk.Checkbutton(doc_col, text="Word (.doc, .docx)", variable=word_var).pack(anchor='w')
tk.Checkbutton(doc_col, text="Excel (.xls, .xlsx)", variable=excel_var).pack(anchor='w') # Uses excel_var
tk.Checkbutton(doc_col, text="PDF (.pdf)", variable=pdf_var).pack(anchor='w')
tk.Checkbutton(doc_col, text="PSD (.psd)", variable=psd_var).pack(anchor='w')

# Custom Extensions
custom_ext_frame = tk.LabelFrame(main_frame, text="Custom Options")
custom_ext_frame.pack(pady=(5,5), padx=5, fill="x")

custom_ext_var = tk.BooleanVar()
custom_ext_entry_var = tk.StringVar()

custom_cb = tk.Checkbutton(custom_ext_frame, text="Add other extensions (comma separated):", 
                           variable=custom_ext_var, command=toggle_custom_entry_state)
custom_cb.pack(side="left", padx=5)

custom_ext_entry = tk.Entry(custom_ext_frame, textvariable=custom_ext_entry_var, width=35, state='disabled')
custom_ext_entry.pack(side="left", padx=5, pady=5, fill="x", expand=True)

# Section 3: Scan Button
scan_button = tk.Button(main_frame, text="3. Run (Scan)", command=run_scan, 
                        bg=BG_COLOR_SCAN, fg=TEXT_COLOR_LIGHT, font=FONT_CONFIG)
scan_button.pack(pady=10, fill="x")

# Section 4: Actions
action_frame = tk.LabelFrame(main_frame, text="4. Select Action")
action_frame.pack(pady=10, padx=5, fill="x")

copy_var = tk.BooleanVar()
delete_var = tk.BooleanVar()
create_excel_var = tk.BooleanVar() # Renamed from excel_var to avoid conflict

tk.Checkbutton(action_frame, text="Copy found files (creates subfolders by extension)", variable=copy_var).pack(anchor='w')
tk.Checkbutton(action_frame, text="Delete all found files (WARNING!)", variable=delete_var, fg=WARN_COLOR_DELETE).pack(anchor='w', pady=(5,0))
tk.Checkbutton(action_frame, text="Create Excel file with file paths", variable=create_excel_var).pack(anchor='w')

# Section 5: Execute Button
action_button = tk.Button(main_frame, text="5. Action (Execute)", command=perform_actions, 
                          bg=BG_COLOR_ACTION, fg=TEXT_COLOR_LIGHT, font=FONT_CONFIG)
action_button.pack(pady=10, fill="x")

root.mainloop()