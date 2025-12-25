import os
import tkinter as tk
import threading
import tempfile
import shutil
from tkinter import ttk
from tkinter import filedialog, messagebox
from docx2pdf import convert as docx_to_pdf
from PIL import Image
from PyPDF2 import PdfMerger

HIGHLIGHT_COLOUR = "#d0e0ff"

# ---------------------------------------------------------
# FILE CONVERSION LOGIC
# ---------------------------------------------------------
def convert_to_pdf(input_path, temp_folder):
    filename = os.path.splitext(os.path.basename(input_path))[0]
    ext = os.path.splitext(input_path)[1].lower()

    output_pdf = os.path.join(temp_folder, f"{filename}.pdf")

    if ext == ".pdf":
        return input_path

    if ext == ".docx":
        docx_to_pdf(input_path, output_pdf)
        return output_pdf

    if ext in [".jpg", ".jpeg", ".png"]:
        image = Image.open(input_path).convert("RGB")
        image.save(output_pdf)
        return output_pdf

    raise ValueError(f"Unsupported file type: {ext}")


def merge_pdfs(pdf_list, output_path):
    

    # If user cancels, stop the merge
    if not output_path:
        return

    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    
    merger.write(output_path)
    merger.close()
    return output_path


# ---------------------------------------------------------
# GUI APPLICATION
# ---------------------------------------------------------
class PDFMergerGUI:
    def __init__(self, root):
        self.output_path = None

        self.root = root
        self.root.title("Multi‑File to PDF Merger")
        self.root.geometry("650x450")

        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=0)
        self.root.rowconfigure(2, weight=0)

        self.root.columnconfigure(0, weight=1)   # canvas expands
        self.root.columnconfigure(1, weight=0)   # scrollbar stays fixed

        # Scrollable canvas with white background
        # Container frame for canvas + scrollbar
        canvas_frame = tk.Frame(root)
        # canvas_frame.pack(fill="both", expand=True)
        canvas_frame.grid(row=0, column=0, sticky="nsew")


        # Scrollable canvas
        self.canvas = tk.Canvas(canvas_frame, bg="white", highlightthickness=0)
        self.canvas.pack(side="left", fill="both", expand=True)

        # Scrollbar (sibling of canvas, not child)
        self.scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Internal frame
        self.frame = tk.Frame(self.canvas, bg="white")
        self.canvas_window = self.canvas.create_window((0, 0), window=self.frame, anchor="nw")

        # Update scrollregion
        self.frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        # Enable mousewheel scrolling
        root.bind_all("<MouseWheel>", self._on_mousewheel)      # Windows + trackpad
        root.bind_all("<Button-4>", self._on_mousewheel)        # Linux scroll up
        root.bind_all("<Button-5>", self._on_mousewheel)        # Linux scroll down

        # Store file rows
        self.file_rows = []
        self.selected_index = None

        # Buttons
        btn_frame = tk.Frame(root, height=35)
        # btn_frame.pack(fill="x", side="bottom")
        btn_frame.grid(row=1, sticky="ew")

        btn_frame.pack_propagate(False)   # prevents the frame from shrinking

        # Inner frame that stays centered
        inner = tk.Frame(btn_frame)
        inner.pack(side="top")

        inner.grid_columnconfigure(2, weight=0)

        tk.Button(inner, text="Add Files", bg="lightgreen", command=self.add_files).grid(row=0, column=0, padx=5, pady=5)
        tk.Button(inner, text="Merge to PDF", bg="lightblue", command=self.start_merge).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(inner, text="Delete All", bg="red", fg="white", command=self.delete_all_rows).grid(row=0, column=2, padx=5, pady=5)
        self.status_frame = tk.Frame(root, height=25)
        self.status_frame.grid(row=2, sticky="ew")
        self.status_frame.grid_propagate(True)

        # Keyboard shortcuts
        root.bind("<Up>", lambda e: self.move_up())
        root.bind("<Down>", lambda e: self.move_down())

        root.focus_set()

    # -----------------------------------------------------
    # SCROLL WHEEL HANDLER
    # -----------------------------------------------------
    def _on_mousewheel(self, event):
        if not getattr(self, "scroll_enabled", True):
            return  # ignore scrollwheel when everything fits

        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # -----------------------------------------------------
    # FILE HANDLING
    # -----------------------------------------------------
    def add_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Select files",
            filetypes=[
                ("Supported files", "*.pdf *.jpg *.jpeg *.png *.docx"),
                ("PDF", "*.pdf"),
                ("Images", "*.jpg *.jpeg *.png"),
                ("Word Documents", "*.docx")
            ]
        )
        for f in file_paths:
            self.add_file_row(f)

    def add_file_row(self, filepath):
        row_frame = tk.Frame(self.frame, bg="white")
        row_frame.pack(fill="x")

        delete_btn = tk.Button(row_frame, text="✖", bg="red", fg="white", width=3,
                               command=lambda rf=row_frame: self.delete_row(rf))
        delete_btn.pack(side="left")

        # NEW: move buttons on the row
        up_btn = tk.Button(row_frame, text="↑", width=2,
                        command=lambda rf=row_frame: self.move_row_up(rf))
        up_btn.pack(side="left")

        down_btn = tk.Button(row_frame, text="↓", width=2,
                            command=lambda rf=row_frame: self.move_row_down(rf))
        down_btn.pack(side="left")

        label = tk.Label(row_frame, text=filepath, anchor="w", bg="white")
        label.pack(side="left", fill="x", expand=True)

        row_frame.bind("<Button-1>", lambda e, rf=row_frame: self.select_row(rf))
        label.bind("<Button-1>", lambda e, rf=row_frame: self.select_row(rf))

        self.file_rows.append(row_frame)
        self._update_scroll_state()

    # -----------------------------------------------------
    # ROW SELECTION
    # -----------------------------------------------------
    def highlight_selected(self):
        for idx, rf in enumerate(self.file_rows):
            delete_btn = rf.winfo_children()[0]  # Delete btn is the first widget in the row
            move_up = rf.winfo_children()[1]  # Move up btn is the second widget in the row
            move_down = rf.winfo_children()[2]  # Move down btn is the third widget in the row
            filename = rf.winfo_children()[-1]  # filename is the last widget in the row

            self._highlight_helper(delete_btn, idx, "red", "#D60000")
            self._highlight_helper(move_up, idx, "lightgrey", "#939393")
            self._highlight_helper(move_down, idx, "lightgrey", "#939393")
            self._highlight_helper(filename, idx, "white", HIGHLIGHT_COLOUR)

    def _highlight_helper(self, element, idx, original_bg, highlight_bg):
        if idx == self.selected_index:
            element.config(bg=highlight_bg)   # highlight colour
        else:
            element.config(bg=original_bg)     # normal colour

    def select_row(self, row_frame):
        self.selected_index = self.file_rows.index(row_frame)
        self.highlight_selected()

    def delete_all_rows(self):
        for row in list(self.file_rows):
            self.delete_row(row)

    # -----------------------------------------------------
    # SORTING LOGIC
    # -----------------------------------------------------
    def move_up(self):
        if self.selected_index is None or self.selected_index == 0:
            return

        i = self.selected_index
        self.file_rows[i], self.file_rows[i - 1] = self.file_rows[i - 1], self.file_rows[i]

        self.refresh_rows()
        self.selected_index -= 1
        self.highlight_selected()
        self._scroll_to_row(self.file_rows[i])

    def move_down(self):
        if self.selected_index is None or self.selected_index == len(self.file_rows) - 1:
            return

        i = self.selected_index
        self.file_rows[i], self.file_rows[i + 1] = self.file_rows[i + 1], self.file_rows[i]

        self.refresh_rows()
        self.selected_index += 1
        self.highlight_selected()
        self._scroll_to_row(self.file_rows[i])

    def move_row_up(self, row_frame):
        idx = self.file_rows.index(row_frame)
        if idx == 0:
            return
        self.file_rows[idx], self.file_rows[idx - 1] = self.file_rows[idx - 1], self.file_rows[idx]
        self.refresh_rows()
        self.selected_index = idx - 1
        self.highlight_selected()
        self._scroll_to_row(row_frame)

    def move_row_down(self, row_frame):
        idx = self.file_rows.index(row_frame)
        if idx == len(self.file_rows) - 1:
            return
        self.file_rows[idx], self.file_rows[idx + 1] = self.file_rows[idx + 1], self.file_rows[idx]
        self.refresh_rows()
        self.selected_index = idx + 1
        self.highlight_selected()
        self._scroll_to_row(row_frame)

    def refresh_rows(self):
        for rf in self.frame.winfo_children():
            rf.pack_forget()
        for rf in self.file_rows:
            rf.pack(fill="x")
        self._update_scroll_state()

    def _update_scroll_state(self):
        self.canvas.update_idletasks()
        scrollregion = self.canvas.bbox("all")
        if scrollregion is None:
            return

        content_height = scrollregion[3] - scrollregion[1]
        canvas_height = self.canvas.winfo_height()

        # Enable scrolling only if content is taller than canvas
        if content_height > canvas_height:
            self.scroll_enabled = True
        else:
            self.scroll_enabled = False

    def _scroll_to_row(self, row_frame):
        self.canvas.update_idletasks()

        # Get the row's position relative to the canvas
        row_y = row_frame.winfo_y()

        # Height of the canvas
        canvas_height = self.canvas.winfo_height()

        # Scroll so the row is roughly centered
        target = row_y - canvas_height // 3
        if target < 0:
            target = 0

        # Convert to scroll fraction
        scrollregion = self.canvas.bbox("all")
        if scrollregion:
            total_height = scrollregion[3] - scrollregion[1]
            if total_height > 0:
                fraction = target / total_height
                self.canvas.yview_moveto(fraction)

    # -----------------------------------------------------
    # DELETE ROW
    # -----------------------------------------------------
    def delete_row(self, row_frame):
        idx = self.file_rows.index(row_frame)
        row_frame.destroy()
        del self.file_rows[idx]

        if self.file_rows:
            self.selected_index = max(0, idx - 1)
            self.highlight_selected()
        else:
            self.selected_index = None

    # -----------------------------------------------------
    # MERGE PROCESS
    # -----------------------------------------------------

    def start_merge(self):
        # Clear previous status
        for widget in self.status_frame.winfo_children():
            widget.destroy()

        # Create progress bar
        self.progress = ttk.Progressbar(self.status_frame, mode="indeterminate", length=200)
        self.progress.pack(pady=3)
        self.progress.start(10)

        self.output_path = filedialog.asksaveasfilename(
            title="Save merged PDF as...",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="merged.pdf"
        )

        # Run merge in background thread
        threading.Thread(target=self._do_merge, daemon=True).start()

    def _merge_complete(self):
        self.progress.stop()
        for widget in self.status_frame.winfo_children():
            widget.destroy()

        tk.Label(self.status_frame, text="Merge completed!", fg="green").pack(pady=3)

    def _do_merge(self):
        try:
            if not self.file_rows:
                messagebox.showerror("Error", "No files selected.")
                return

            files = []
            for rf in self.file_rows:
                label = rf.winfo_children()[-1]
                files.append(label.cget("text"))

            try:
                temp_folder = tempfile.mkdtemp(prefix="pdfmerge_")

                converted = []
                for f in files:
                    try:
                        out = convert_to_pdf(f, temp_folder)
                        converted.append(out)
                    except Exception as e:
                        # Make the hidden error visible
                        self.root.after(0, lambda: messagebox.showerror("Conversion Error", f"{f}\n\n{e}"))
                        return

                merge_pdfs(converted, self.output_path)
                shutil.rmtree(temp_folder, ignore_errors=True)

            except Exception as e:
                messagebox.showerror("Error", str(e))

        finally:
            # When done, update UI on the main thread
            self.root.after(0, self._merge_complete)

    def show_progress(self):
        # Create popup window
        self.progress_win = tk.Toplevel(self.root)
        self.progress_win.title("Merging...")
        self.progress_win.geometry("300x80")
        self.progress_win.resizable(False, False)

        # Prevent user from closing it
        self.progress_win.transient(self.root)
        self.progress_win.grab_set()

        label = tk.Label(self.progress_win, text="Merging files, please wait...")
        label.pack(pady=10)

        # Progress bar
        self.progress = ttk.Progressbar(self.progress_win, mode="indeterminate", length=250)
        self.progress.pack(pady=5)
        self.progress.start(10)  # speed of animation


# ---------------------------------------------------------
# RUN THE APP
# ---------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Multi‑File to PDF Merger")

    # Minimum window size
    root.minsize(400, 300)

    app = PDFMergerGUI(root)
    root.mainloop()
