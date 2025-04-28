# docx_processor/gui.py
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox
import sys
import os
import threading
from processor import main as process_document


class TextRedirector:
    """Redirects stdout to a text widget."""

    def __init__(self, text_widget: scrolledtext.ScrolledText):
        self.text_widget = text_widget
        self.buffer = ""

    def write(self, string: str):
        self.buffer += string
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        self.text_widget.config(state=tk.DISABLED)

    def flush(self):
        pass


class DocxProcessorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DOCX Processor")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)

        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready")

        self._create_widgets()
        self._setup_layout()

    def _create_widgets(self):
        # Frame for file selection
        self.file_frame = ttk.LabelFrame(self.root, text="File Selection")

        # Input file selection
        ttk.Label(self.file_frame, text="Input DOCX File:").grid(
            row=0, column=0, sticky=tk.W, padx=5, pady=5
        )
        ttk.Entry(self.file_frame, textvariable=self.input_file_path, width=50).grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Button(
            self.file_frame, text="Browse...", command=self._browse_input_file
        ).grid(row=0, column=2, padx=5, pady=5)

        # Output file selection
        ttk.Label(self.file_frame, text="Output File:").grid(
            row=1, column=0, sticky=tk.W, padx=5, pady=5
        )
        ttk.Entry(self.file_frame, textvariable=self.output_file_path, width=50).grid(
            row=1, column=1, padx=5, pady=5
        )
        ttk.Button(
            self.file_frame, text="Browse...", command=self._browse_output_file
        ).grid(row=1, column=2, padx=5, pady=5)

        # Button frame
        button_frame = ttk.Frame(self.root)

        # Process button
        self.process_button = ttk.Button(
            button_frame, text="Process Document", command=self._process_document
        )
        self.process_button.pack(side=tk.LEFT, padx=5)

        # Clear log button
        self.clear_log_button = ttk.Button(
            button_frame, text="Clear Log", command=self._clear_log
        )
        self.clear_log_button.pack(side=tk.LEFT, padx=5)

        # Log area
        log_frame = ttk.LabelFrame(self.root, text="Log")
        self.log_text = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, width=80, height=20
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.config(state=tk.DISABLED)

        # Status bar
        status_bar = ttk.Frame(self.root)
        ttk.Label(status_bar, textvariable=self.status_text).pack(side=tk.LEFT, padx=5)
        self.progress_bar = ttk.Progressbar(
            status_bar, mode="indeterminate", length=100
        )
        self.progress_bar.pack(side=tk.RIGHT, padx=5)

        # Store frames for layout
        self.button_frame = button_frame
        self.log_frame = log_frame
        self.status_bar = status_bar

    def _setup_layout(self):
        # Configure the grid layout
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        # Place the widgets
        self.file_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.button_frame.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.log_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.status_bar.grid(row=3, column=0, padx=10, pady=5, sticky="ew")

    def _browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Select DOCX File",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if file_path:
            self.input_file_path.set(file_path)
            # Set default output path
            if not self.output_file_path.get():
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                output_path = os.path.join(
                    os.path.dirname(file_path), f"{base_name}_processed.xlsx"
                )
                self.output_file_path.set(output_path)

    def _browse_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="Save Output File",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Files", "*.xlsx"),
                ("CSV Files", "*.csv"),
                ("All Files", "*.*"),
            ],
        )
        if file_path:
            self.output_file_path.set(file_path)

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _process_document(self):
        input_path = self.input_file_path.get()
        output_path = self.output_file_path.get()

        if not input_path or not output_path:
            messagebox.showerror("Error", "Please select both input and output files.")
            return

        # Validate file extensions
        if not input_path.lower().endswith(".docx"):
            messagebox.showerror("Error", "Input file must be a .docx file.")
            return

        if not output_path.lower().endswith((".xlsx", ".csv")):
            messagebox.showerror("Error", "Output file must be an .xlsx or .csv file.")
            return

        # Disable UI elements during processing
        self.process_button.config(state=tk.DISABLED)
        self.status_text.set("Processing...")
        self.progress_bar.start(10)

        # Clear the log
        self._clear_log()

        # Run processing in a separate thread to keep the UI responsive
        threading.Thread(
            target=self._run_processing, args=(input_path, output_path), daemon=True
        ).start()

    def _run_processing(self, input_path, output_path):
        # Redirect stdout to the log text widget
        original_stdout = sys.stdout
        sys.stdout = TextRedirector(self.log_text)

        try:
            # Process the document
            result = process_document(input_path, output_path)

            if result == 0:
                final_status = "Document processed successfully!"
                self.root.after(0, lambda: messagebox.showinfo("Success", final_status))
            else:
                final_status = f"Document processing failed with code: {result}"
                self.root.after(0, lambda: messagebox.showerror("Error", final_status))

            self._log(f"\n{final_status}")

        except Exception as e:
            error_message = f"Error during processing: {str(e)}"
            self._log(f"\n{error_message}")
            self.root.after(0, lambda: messagebox.showerror("Error", error_message))
        finally:
            # Restore stdout
            sys.stdout = original_stdout

            # Re-enable UI elements
            self.root.after(0, lambda: self._reset_ui())

    def _reset_ui(self):
        self.process_button.config(state=tk.NORMAL)
        self.status_text.set("Ready")
        self.progress_bar.stop()

    def _log(self, message: str):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)


def run():
    root = tk.Tk()
    app = DocxProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    run()
