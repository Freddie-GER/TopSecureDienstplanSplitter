import os
import re
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
from PyPDF2 import PdfReader, PdfWriter
import threading
import queue

def remove_problematic_strings(line: str) -> str:
    # Remove specific unwanted single characters
    line = line.replace("T", "", 1)
    line = line.replace("e", "", 1)
    line = line.replace("x", "", 1)
    line = line.replace("t", "", 1)
    line = line.replace("5", "", 1)
    
    # Remove ": " only if it's followed by a lowercase letter (including umlauts)
    line = re.sub(r': (?=[a-zäöüß])', '', line, count=1)
    
    # Remove remaining ":" characters
    line = line.replace(":", "", 1)
    
    return line.strip()

def insert_space_before_upper(name: str) -> str:
    return re.sub(r'(?<=[a-zäöüß])(?=[A-ZÄÖÜ])', ' ', name)

def handle_compound_names(line: str) -> str:
    line = re.sub(r'\s*-\s*', '-', line)
    return line

class DienstplanSplitter:
    def __init__(self, root):
        self.root = root
        self.root.title("Dienstplan Splitter")

        window_width = 600
        window_height = 400
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # File selection frame
        self.file_frame = ttk.LabelFrame(main_frame, text="PDF-Datei", padding="5")
        self.file_frame.pack(fill=tk.X, pady=(0, 10))

        self.file_label = ttk.Label(self.file_frame, text="Keine Datei ausgewählt")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.output_button = ttk.Button(
            self.file_frame, text="Zielverzeichnis wählen", command=self.select_output_dir
        )
        self.output_button.pack(side=tk.RIGHT, padx=(10, 0))

        self.new_file_button = ttk.Button(
            self.file_frame, text="Andere Datei wählen", command=self.select_file
        )
        self.new_file_button.pack(side=tk.RIGHT, padx=(10, 0))
        self.new_file_button.pack_forget()  # Initially hidden

        self.select_button = ttk.Button(
            self.file_frame, text="PDF auswählen", command=self.select_file
        )
        self.select_button.pack(side=tk.RIGHT, padx=(10, 0))

        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(
            main_frame, variable=self.progress_var, maximum=100
        )
        self.progress.pack(fill=tk.X, pady=(0, 10))

        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.pack(fill=tk.X)

        self.output_text = tk.Text(main_frame, height=10, wrap=tk.WORD)
        self.output_text.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.output_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.output_text.configure(yscrollcommand=scrollbar.set)

        self.process_button = ttk.Button(
            main_frame, text="Dienstpläne aufteilen", command=self.process_file
        )
        self.process_button.pack(pady=(20, 0))

        self.selected_file = None
        self.output_dir = None
        self.processing = False

        # Variables to track the last generated PDF
        self.last_writer = None
        self.last_info = None

        # Queue for thread-safe logging
        self.queue = queue.Queue()
        self.root.after(100, self.process_queue)

    def handle_drop(self, event):
        try:
            # Get the dropped data
            data = event.data if hasattr(event, 'data') else event.widget.selection_get()
            
            # Clean up the file path (remove {} and quotes if present)
            file_path = data.strip('{}').strip('"')
            
            if file_path.lower().endswith('.pdf'):
                self.selected_file = file_path
                self.file_label.config(text=os.path.basename(file_path))
                self.new_file_button.config(state=tk.NORMAL)
                self.status_label.config(text="")
                self.output_text.delete(1.0, tk.END)
                self.progress_var.set(0)
            else:
                messagebox.showerror("Fehler", "Bitte wählen Sie eine PDF-Datei aus.")
        except:
            # If drag and drop fails, silently ignore
            pass

    def select_output_dir(self):
        dir_path = filedialog.askdirectory(title="Ausgabeverzeichnis wählen")
        if dir_path:
            self.output_dir = dir_path

    def get_output_dir(self):
        if not self.selected_file:
            return "split_schedules"

        # Get the base name of the original file without extension
        base_name = os.path.splitext(os.path.basename(self.selected_file))[0]
        
        if self.output_dir:
            # If user selected an output directory, create a subdirectory there
            return os.path.join(self.output_dir, f"{base_name} Splitted")
        else:
            # If no output directory selected, create in default location
            return os.path.join("split_schedules", f"{base_name} Splitted")

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="PDF-Datei auswählen", filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            self.selected_file = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.select_button.pack_forget()  # Hide the initial select button
            self.new_file_button.pack(side=tk.RIGHT, padx=(10, 0))  # Show the new file button
            self.status_label.config(text="")
            self.output_text.delete(1.0, tk.END)
            self.progress_var.set(0)

    def log_message(self, message):
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)

    def log_message_threadsafe(self, message):
        self.queue.put(lambda: self.log_message(message))

    def process_queue(self):
        try:
            while True:
                task = self.queue.get_nowait()
                task()
        except queue.Empty:
            pass
        self.root.after(100, self.process_queue)

    def process_pdf(self):
        try:
            if not os.path.exists(self.selected_file):
                self.log_message_threadsafe("Fehler: Datei nicht gefunden!")
                return

            output_dir = self.get_output_dir()
            os.makedirs(output_dir, exist_ok=True)

            reader = PdfReader(self.selected_file)
            total_pages = len(reader.pages)
            processed_pages = 0

            with pdfplumber.open(self.selected_file) as pdf:
                for current_page in range(total_pages):
                    try:
                        progress = ((current_page + 1) / total_pages) * 100
                        self.progress_var.set(progress)
                        self.log_message_threadsafe(f"Verarbeite Seite {current_page + 1} von {total_pages}")

                        page = pdf.pages[current_page]
                        text = page.extract_text()

                        self.log_message_threadsafe(f"\n--- Debug: Text von Seite {current_page + 1} ---")
                        if text:
                            preview_text = text[:500].replace("\n", "\\n")
                            self.log_message_threadsafe(preview_text)
                        else:
                            self.log_message_threadsafe("Keine Textinformationen gefunden.")
                        self.log_message_threadsafe("-------------------\n")

                        # Attempt to extract name and date
                        last_name, first_name, date_str = self.extract_name_and_date(text if text else "")

                        if all([last_name, first_name, date_str]):
                            # Successfully extracted information; create a new PDF
                            writer = PdfWriter()
                            writer.add_page(reader.pages[current_page])

                            safe_last_name = re.sub(r'[\\/*?:"<>|]', "", last_name)
                            safe_first_name = re.sub(r'[\\/*?:"<>|]', "", first_name)
                            output_filename = f"{safe_last_name}_{safe_first_name}_{date_str}.pdf"
                            output_path = os.path.join(output_dir, output_filename)

                            with open(output_path, "wb") as output_file:
                                writer.write(output_file)

                            self.log_message_threadsafe(f"Erstellt: {output_filename}")

                            # Update last_writer and last_info
                            self.last_writer = writer
                            self.last_info = {
                                'last_name': last_name,
                                'first_name': first_name,
                                'date_str': date_str
                            }

                        else:
                            # Failed to extract information; check for "Seite 2"
                            if text and "Seite 2" in text:
                                self.log_message_threadsafe("Hinweis: Fortsetzungsseite erkannt. Anhängen an die vorherige PDF.")

                                if self.last_writer and self.last_info:
                                    # Append the current page to the last_writer
                                    self.last_writer.add_page(reader.pages[current_page])

                                    # Overwrite the existing PDF with the appended page
                                    safe_last_name = re.sub(r'[\\/*?:"<>|]', "", self.last_info['last_name'])
                                    safe_first_name = re.sub(r'[\\/*?:"<>|]', "", self.last_info['first_name'])
                                    output_filename = f"{safe_last_name}_{safe_first_name}_{self.last_info['date_str']}.pdf"
                                    output_path = os.path.join(output_dir, output_filename)

                                    with open(output_path, "wb") as output_file:
                                        self.last_writer.write(output_file)

                                    self.log_message_threadsafe(f"Erstellt (mit Anhang): {output_filename}")
                                else:
                                    self.log_message_threadsafe("Warnung: Keine vorherige PDF zum Anhängen gefunden.")
                            else:
                                self.log_message_threadsafe(f"Warnung: Keine vollständigen Informationen auf Seite {current_page + 1} gefunden")

                        processed_pages += 1

                    except Exception as e:
                        self.log_message_threadsafe(f"Fehler auf Seite {current_page + 1}: {e}")

            self.progress_var.set(100)
            self.log_message_threadsafe("\nVerarbeitung abgeschlossen!")

        except Exception as e:
            self.log_message_threadsafe(f"Unerwarteter Fehler: {e}")

        finally:
            self.processing = False
            self.process_button.config(state=tk.NORMAL)
            self.select_button.config(state=tk.NORMAL)
            self.select_button.config(state=tk.NORMAL)

    def process_file(self):
        if not self.selected_file:
            messagebox.showerror("Fehler", "Bitte wählen Sie zuerst eine PDF-Datei aus.")
            return

        if self.processing:
            return

        self.processing = True
        self.process_button.config(state=tk.DISABLED)
        self.select_button.config(state=tk.DISABLED)
        self.output_text.delete(1.0, tk.END)
        self.progress_var.set(0)

        thread = threading.Thread(target=self.process_pdf)
        thread.daemon = True
        thread.start()

    def extract_name_and_date(self, text):
        lines = text.split("\n")
        name_line = None

        for i, line in enumerate(lines):
            if "Herr/Frau" in line and i + 1 < len(lines):
                # Get the next line and clean it
                next_line = remove_problematic_strings(lines[i + 1].strip())
                
                # If it's a salutation, try the line after
                if next_line.lower() in ["herrn", "frau", "herr"]:
                    if i + 2 < len(lines):
                        name_line = lines[i + 2].strip()
                else:
                    name_line = lines[i + 1].strip()
                break

        if not name_line:
            self.log_message_threadsafe("Warnung: Kein Name gefunden.")
            return None, None, None

        name_line = remove_problematic_strings(name_line)
        name_line = handle_compound_names(name_line)

        # Split the name into words and handle multiple names
        name_parts = name_line.split()
        if len(name_parts) >= 2:
            last_name = name_parts[-1]  # Last word is the last name
            first_name = "_".join(name_parts[:-1])  # All other words are first names, joined with underscore
        else:
            # Try to match single name consisting of uppercase letters
            single_name_match = re.search(r"([A-ZÄÖÜÉÈÊËÀÁÂÃÌÍÎÏÒÓÔÕÙÚÛÝ]{2,})", name_line)
            if single_name_match:
                first_name = single_name_match.group(1)
                last_name = ""
            else:
                self.log_message_threadsafe(f"Warnung: Ungültiger Name: {name_line}")
                return None, None, None

        # Extract and parse date
        date_match = re.search(r"Zeitraum:\s*(.*?)(?=\s*\d{2}\.\d{2}\.\d{4}|$)", text)
        if not date_match:
            self.log_message_threadsafe("Warnung: Kein Zeitraum gefunden.")
            return None, None, None

        date_line = date_match.group(1).strip()

        german_months = {
            "Januar": "01", "Februar": "02", "März": "03", "April": "04",
            "Mai": "05", "Juni": "06", "Juli": "07", "August": "08",
            "September": "09", "Oktober": "10", "November": "11", "Dezember": "12"
        }

        # Get the last date in the line first
        full_date_match = re.search(r"Zeitraum:.*(\d{2}\.\d{2}\.\d{4})(?:\s|$)", text)
        if not full_date_match:
            # Try to find any date at the end of the line
            full_date_match = re.search(r".*(\d{2}\.\d{2}\.\d{4})(?:\s|$)", text)
            if not full_date_match:
                self.log_message_threadsafe("Warnung: Kein Datum am Ende der Zeile gefunden.")
                return None, None, None

        last_date = full_date_match.group(1)
        try:
            date_obj = datetime.strptime(last_date, "%d.%m.%Y")
        except ValueError as e:
            self.log_message_threadsafe(f"Fehler beim Parsen des Datums: {e}")
            return None, None, None

        # Try word + year format
        month_match = re.search(
            r"(" + "|".join(german_months.keys()) + r")\s+(\d{4})", date_line
        )
        
        if month_match:
            month_word = month_match.group(1)
            year = month_match.group(2)
            month_number = german_months[month_word]
            
            # If the date at the end is in a different month than the word month
            # OR in a different year, use only month and year
            if year != str(date_obj.year) or month_number != f"{date_obj.month:02d}":
                name_for_file = first_name if not last_name else last_name
                return name_for_file, first_name, f"{year}_{month_number}"
        
        # If no month word found OR the date is in the same month,
        # use the full date from the end
        name_for_file = first_name if not last_name else last_name
        return name_for_file, first_name, f"{date_obj.year}_{date_obj.month:02d}_{date_obj.day:02d}"

def main():
    root = tk.Tk()
    app = DienstplanSplitter(root)
    root.mainloop()

if __name__ == "__main__":
    main()