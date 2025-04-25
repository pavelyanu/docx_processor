# gui.py
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import threading
import sys

import docx_processor.processor as processor


class DocxProcessorGUI:
    # Default values for configuration
    DEFAULT_TABLE_INDEX = 3  # Will display as 3 but use as 2 (0-based)
    DEFAULT_HEADER_LEN = 3
    DEFAULT_FOOTER_LEN = 2
    DEFAULT_ACCOUNT_INDEX = 4

    def __init__(self, root):
        self.root = root
        self.root.title("DOCX Transaction Processor")
        self.root.geometry("700x500")
        self.setup_ui()

    def setup_ui(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding=10)
        file_frame.pack(fill=tk.X, pady=5)

        # Input file
        ttk.Label(file_frame, text="Input DOCX file:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        self.input_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.input_path, width=50).grid(
            row=0, column=1, padx=5
        )
        ttk.Button(file_frame, text="Browse...", command=self.browse_input).grid(
            row=0, column=2
        )

        # Output file
        ttk.Label(file_frame, text="Output Excel file:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.output_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.output_path, width=50).grid(
            row=1, column=1, padx=5
        )
        ttk.Button(file_frame, text="Browse...", command=self.browse_output).grid(
            row=1, column=2
        )

        # Configuration section
        config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding=10)
        config_frame.pack(fill=tk.X, pady=5)

        # Table index (1-based in UI)
        ttk.Label(config_frame, text="Table index:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        self.table_index = tk.IntVar(value=self.DEFAULT_TABLE_INDEX)
        ttk.Spinbox(
            config_frame, from_=1, to=10, textvariable=self.table_index, width=5
        ).grid(row=0, column=1, sticky=tk.W)

        # Header length
        ttk.Label(config_frame, text="Header length:").grid(
            row=0, column=2, sticky=tk.W, pady=5, padx=(20, 0)
        )
        self.header_len = tk.IntVar(value=self.DEFAULT_HEADER_LEN)
        ttk.Spinbox(
            config_frame, from_=0, to=10, textvariable=self.header_len, width=5
        ).grid(row=0, column=3, sticky=tk.W)

        # Footer length
        ttk.Label(config_frame, text="Footer length:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.footer_len = tk.IntVar(value=self.DEFAULT_FOOTER_LEN)
        ttk.Spinbox(
            config_frame, from_=0, to=10, textvariable=self.footer_len, width=5
        ).grid(row=1, column=1, sticky=tk.W)

        # Account cell index
        ttk.Label(config_frame, text="Account cell index:").grid(
            row=1, column=2, sticky=tk.W, pady=5, padx=(20, 0)
        )
        self.account_index = tk.IntVar(value=self.DEFAULT_ACCOUNT_INDEX)
        ttk.Spinbox(
            config_frame, from_=0, to=10, textvariable=self.account_index, width=5
        ).grid(row=1, column=3, sticky=tk.W)

        # Restore defaults button
        ttk.Button(
            config_frame, text="Restore Defaults", command=self.restore_defaults
        ).grid(row=2, column=0, columnspan=4, pady=10)

        # Process button
        self.process_btn = ttk.Button(
            main_frame, text="Process Document", command=self.process_document
        )
        self.process_btn.pack(pady=10)

        # Status log
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Scrollable text area for log
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=10)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Make text widget read-only
        self.log_text.config(state=tk.DISABLED)

    def restore_defaults(self):
        """Reset all configuration fields to default values"""
        self.table_index.set(self.DEFAULT_TABLE_INDEX)
        self.header_len.set(self.DEFAULT_HEADER_LEN)
        self.footer_len.set(self.DEFAULT_FOOTER_LEN)
        self.account_index.set(self.DEFAULT_ACCOUNT_INDEX)
        self.log("Default configuration restored")

    def browse_input(self):
        filename = filedialog.askopenfilename(
            title="Select DOCX file",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if filename:
            self.input_path.set(filename)
            # Auto-suggest output filename
            output_name = os.path.splitext(filename)[0] + "_processed.xlsx"
            self.output_path.set(output_name)

    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Save Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")],
        )
        if filename:
            self.output_path.set(filename)

    def log(self, message):
        """Add message to log text widget"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)  # Scroll to end
        self.log_text.config(state=tk.DISABLED)
        # Force update to show message immediately
        self.root.update_idletasks()

    def process_document(self):
        """Process the document in a separate thread"""
        # Get values from UI
        input_path = self.input_path.get()
        output_path = self.output_path.get()

        # Validate inputs
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid input DOCX file")
            return

        if not output_path:
            messagebox.showerror("Error", "Please specify an output Excel file")
            return

        # Clear log
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        # Disable process button
        self.process_btn.config(state=tk.DISABLED)

        # Override print to capture output in our log
        original_print = print

        def custom_print(message):
            original_print(message)  # Still print to console
            self.root.after(0, lambda: self.log(str(message)))  # Log to UI

        # Start processing in a thread
        def process_thread():
            try:
                # Replace print function to capture output
                processor.print = custom_print

                # Get configuration from UI (convert table index from 1-based to 0-based)
                table_index = self.table_index.get() - 1  # Convert to 0-based
                header_len = self.header_len.get()
                footer_len = self.footer_len.get()
                account_index = self.account_index.get()

                # Log with user-friendly 1-based index
                self.log(
                    f"Starting processing with: Table={self.table_index.get()} (index {table_index}), Header={header_len}, Footer={footer_len}, Account Index={account_index}"
                )

                # Set up transaction row parsing config
                transaction_row_parsing_config = processor.TransactionRowParsingConfig(
                    field_count=3,
                    id_test_func=lambda s: s.isdigit(),
                )

                # Create table format with our settings
                table_format = processor.TableFormat(
                    header_len=header_len,
                    footer_len=footer_len,
                    account_cell_index=account_index,
                    transaction_row_parsing_config=transaction_row_parsing_config,
                )

                # Create document format with our settings
                document_format = processor.InputDocumentFormat(
                    path=input_path,
                    table_index=table_index,  # Using 0-based index
                )

                # Set up configurations
                loading_config, processing_config, export_config = (
                    processor.setup_configuration(input_path, output_path)
                )

                # Override with our custom settings
                loading_config.document_format = document_format
                processing_config.table_format = table_format

                # Process the document - adapted version of main()
                try:
                    # Load the document
                    custom_print(
                        f"Loading document: {loading_config.document_format.path}"
                    )
                    document = loading_config.loading_strategy(
                        loading_config.document_format
                    )

                    # Choose the table
                    custom_print(
                        f"Selecting table {self.table_index.get()} (index {loading_config.document_format.table_index})"
                    )
                    table = loading_config.table_choose_strategy(
                        document, loading_config.document_format
                    )

                    # Process the table
                    custom_print("Processing transactions...")
                    transactions = processor.extract_transactions(
                        processing_config, table
                    )

                    # Export the result
                    custom_print(
                        f"Exporting to {export_config.output_document_format.path}"
                    )
                    export_config.export_strategy(
                        transactions, export_config.output_document_format
                    )

                    custom_print(
                        f"✓ Successfully processed document and exported to {export_config.output_document_format.path}"
                    )
                    result = 0

                except processor.DocumentLoadingError as e:
                    custom_print(f"❌ Document loading error: {str(e)}")
                    result = 1
                except processor.TableProcessingError as e:
                    custom_print(f"❌ Table processing error: {str(e)}")
                    result = 1
                except processor.ExportError as e:
                    custom_print(f"❌ Export error: {str(e)}")
                    result = 1
                except Exception as e:
                    custom_print(f"❌ Unexpected error: {str(e)}")
                    result = 1

                # Restore original print
                processor.print = original_print

                # Show result
                if result == 0:
                    self.root.after(
                        0,
                        lambda: messagebox.showinfo(
                            "Success",
                            f"Document processed successfully!\nOutput saved to: {output_path}",
                        ),
                    )
                else:
                    self.root.after(
                        0,
                        lambda: messagebox.showerror(
                            "Error", "Processing failed. See log for details."
                        ),
                    )

            except Exception as e:
                # Restore original print
                processor.print = original_print

                # Log error and show message box
                error_message = f"Error: {str(e)}"
                self.root.after(0, lambda: self.log(error_message))
                self.root.after(0, lambda: messagebox.showerror("Error", error_message))

            finally:
                # Re-enable process button
                self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL))

        # Start thread
        threading.Thread(target=process_thread, daemon=True).start()


def main():
    root = tk.Tk()
    app = DocxProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
