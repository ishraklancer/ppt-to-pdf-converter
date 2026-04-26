import os
import argparse
import sys
import glob
import threading
import time

try:
    import comtypes.client
    import comtypes
except ImportError:
    print("Error: comtypes module not found. Please install it using 'pip install -r requirements.txt'")
    sys.exit(1)

def convert_ppt_to_pdf_batch(input_folder, delete_source=False, recursive=False, progress_callback=None, progress_update_callback=None):
    # Find all .ppt and .pptx files in the directory
    if recursive:
        ppt_files = glob.glob(os.path.join(input_folder, "**", "*.ppt"), recursive=True) + glob.glob(os.path.join(input_folder, "**", "*.pptx"), recursive=True)
    else:
        ppt_files = glob.glob(os.path.join(input_folder, "*.ppt")) + glob.glob(os.path.join(input_folder, "*.pptx"))
    
    if not ppt_files:
        msg = f"No .ppt or .pptx files found in {input_folder}"
        if progress_callback: progress_callback(msg)
        else: print(msg)
        return

    total_files = len(ppt_files)
    msg = f"Found {total_files} PowerPoint file(s). Starting conversion..."
    if progress_callback: progress_callback(msg)
    else: print(msg)

    if progress_update_callback:
        progress_update_callback(0, total_files)

    completed_files = 0
    progress_lock = threading.Lock()
    ppt_com_lock = threading.Lock()
    
    # Use up to 50% of available CPU threads, minimum 1
    cpu_count = os.cpu_count() or 2
    max_workers = max(1, cpu_count // 2)
    max_workers = min(max_workers, total_files) # Don't start more threads than files
    
    msg = f"Using {max_workers} threads for conversion."
    if progress_callback: progress_callback(msg)
    else: print(msg)

    def worker_thread(files_chunk):
        nonlocal completed_files
        comtypes.CoInitialize()
        try:
            with ppt_com_lock:
                powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        except Exception as e:
            msg = f"Error starting PowerPoint in thread: {e}"
            if progress_callback: progress_callback(msg)
            else: print(msg)
            comtypes.CoUninitialize()
            
            # Even if failed, mark files as completed so progress bar doesn't hang
            with progress_lock:
                completed_files += len(files_chunk)
                if progress_update_callback:
                    progress_update_callback(completed_files, total_files)
            return

        for ppt_path in files_chunk:
            ppt_path = os.path.abspath(ppt_path)
            pdf_path = os.path.splitext(ppt_path)[0] + ".pdf"
            
            max_retries = 3
            for attempt in range(max_retries):
                deck = None
                try:
                    if attempt == 0:
                        msg = f"Converting: {os.path.basename(ppt_path)} -> {os.path.basename(pdf_path)}"
                    else:
                        msg = f"Retrying conversion ({attempt+1}/{max_retries}): {os.path.basename(ppt_path)}"
                        
                    if progress_callback: progress_callback(msg)
                    else: print(msg)
                    
                    with ppt_com_lock:
                        deck = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
                        deck.SaveAs(pdf_path, 32)
                        deck.Close()
                    deck = None
                    
                    msg = f"Successfully converted: {os.path.basename(ppt_path)}"
                    if progress_callback: progress_callback(msg)
                    else: print(msg)
                    
                    if delete_source:
                        os.remove(ppt_path)
                        msg = f"Deleted source file: {os.path.basename(ppt_path)}"
                        if progress_callback: progress_callback(msg)
                        else: print(msg)
                        
                    break # Success, exit retry loop
                    
                except Exception as e:
                    if deck is not None:
                        try: 
                            with ppt_com_lock:
                                deck.Close()
                        except Exception: 
                            pass # nosec B110
                        
                    if attempt < max_retries - 1:
                        time.sleep(2) # Give PowerPoint a moment to unblock
                        continue
                        
                    msg = f"Error converting {os.path.basename(ppt_path)}: {e}"
                    if progress_callback: progress_callback(msg)
                    else: print(msg)
                    break # Reached max retries, exit retry loop

            # Update progress
            with progress_lock:
                completed_files += 1
                if progress_update_callback:
                    progress_update_callback(completed_files, total_files)

        try:
            with ppt_com_lock:
                powerpoint.Quit()
        except Exception:
            pass # nosec B110
        comtypes.CoUninitialize()

    # Split files into chunks
    chunk_size = (total_files + max_workers - 1) // max_workers
    chunks = [ppt_files[i:i + chunk_size] for i in range(0, total_files, chunk_size)]
    
    threads = []
    for chunk in chunks:
        t = threading.Thread(target=worker_thread, args=(chunk,))
        t.start()
        threads.append(t)
        
    for t in threads:
        t.join()

    msg = "Conversion batch finished."
    if progress_callback: progress_callback(msg)
    else: print(msg)

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext
    from tkinter import ttk

    def select_folder():
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            folder_var.set(folder_selected)

    def start_conversion():
        folder = folder_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Please select a valid folder.")
            return
            
        delete_src = delete_var.get()
        is_recursive = recursive_var.get()
        
        btn_convert.config(state=tk.DISABLED)
        btn_browse.config(state=tk.DISABLED)
        log_text.delete(1.0, tk.END)
        progress_bar['value'] = 0
        
        def log_msg(msg):
            def append_log():
                log_text.insert(tk.END, msg + "\n")
                log_text.see(tk.END)
            root.after(0, append_log)

        def update_progress(current, total):
            def set_progress():
                if total > 0:
                    progress_bar['maximum'] = total
                    progress_bar['value'] = current
            root.after(0, set_progress)
            
        def worker():
            convert_ppt_to_pdf_batch(folder, delete_src, is_recursive, log_msg, update_progress)
            root.after(0, lambda: btn_convert.config(state=tk.NORMAL))
            root.after(0, lambda: btn_browse.config(state=tk.NORMAL))

        threading.Thread(target=worker, daemon=True).start()

    root = tk.Tk()
    root.title("Batch PPT to PDF Converter")
    root.geometry("550x500")

    folder_var = tk.StringVar()
    # Default is Keep (delete_var = False)
    delete_var = tk.BooleanVar(value=False)
    recursive_var = tk.BooleanVar(value=False)

    frame_top = tk.Frame(root, pady=10, padx=10)
    frame_top.pack(fill=tk.X)

    tk.Label(frame_top, text="Input Folder:").pack(side=tk.LEFT)
    tk.Entry(frame_top, textvariable=folder_var, width=45).pack(side=tk.LEFT, padx=5)
    btn_browse = tk.Button(frame_top, text="Browse", command=select_folder)
    btn_browse.pack(side=tk.LEFT)

    frame_mid = tk.Frame(root, pady=5, padx=10)
    frame_mid.pack(fill=tk.X)

    tk.Checkbutton(frame_mid, text="Include nested folders", variable=recursive_var).pack(anchor=tk.W)
    tk.Checkbutton(frame_mid, text="Delete source .ppt/.pptx files after successful conversion (Default: Keep)", variable=delete_var).pack(anchor=tk.W)

    frame_btn = tk.Frame(root, pady=10, padx=10)
    frame_btn.pack(fill=tk.X)

    btn_convert = tk.Button(frame_btn, text="Start Conversion", command=start_conversion, bg="green", fg="white", font=("Arial", 10, "bold"), pady=5)
    btn_convert.pack(fill=tk.X)

    frame_progress = tk.Frame(root, pady=5, padx=10)
    frame_progress.pack(fill=tk.X)
    progress_bar = ttk.Progressbar(frame_progress, orient='horizontal', mode='determinate')
    progress_bar.pack(fill=tk.X)

    frame_log = tk.Frame(root, pady=5, padx=10)
    frame_log.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame_log, text="Log Output:").pack(anchor=tk.W)
    log_text = scrolledtext.ScrolledText(frame_log, width=60, height=15)
    log_text.pack(fill=tk.BOTH, expand=True)

    root.mainloop()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Batch convert PPT/PPTX files to PDF.")
    parser.add_argument("-f", "--folder", help="Input folder containing the PPT/PPTX files.", default=None)
    parser.add_argument("--recursive", action="store_true", help="Include nested folders")
    parser.add_argument("--delete", action="store_true", help="Delete the source PPT/PPTX files after successful conversion. (Default: Keep)")
    parser.add_argument("--cli", action="store_true", help="Run in CLI mode without GUI")
    
    args = parser.parse_args()
    
    if args.folder or args.cli:
        if not args.folder:
            print("Error: You must specify a folder with -f/--folder when running in CLI mode.")
            sys.exit(1)
        if not os.path.isdir(args.folder):
            print(f"Error: The directory '{args.folder}' does not exist.")
            sys.exit(1)
        convert_ppt_to_pdf_batch(args.folder, args.delete, args.recursive)
    else:
        # Run GUI mode by default if no arguments are provided
        run_gui()
