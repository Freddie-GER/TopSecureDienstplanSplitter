"""
DienstplanSplitter Virtual Printer Service

This module implements a virtual printer service that:
1. Uses Windows Print to PDF as base driver
2. Monitors the output directory for new PDFs
3. Automatically processes them with DienstplanSplitter logic
"""

import os
import sys
import time
import shutil
import win32print
import win32api
from pathlib import Path
from typing import Optional
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Import our existing splitting logic
from ..gui.dienstplan_splitter_gui import DienstplanSplitter

class PDFHandler(FileSystemEventHandler):
    """Handles PDF files created by Windows Print to PDF."""
    
    def __init__(self, output_dir: Path):
        self.output_dir = output_dir
        # Initialize our existing splitter logic
        self.splitter = DienstplanSplitter(None)  # None because we don't need GUI
        
    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith('.pdf'):
            # Wait a moment to ensure the file is completely written
            time.sleep(1)
            self.process_pdf(Path(event.src_path))
            
    def process_pdf(self, pdf_path: Path):
        """Process a newly created PDF using our existing splitting logic."""
        try:
            # Create output directory
            output_dir = self.output_dir / pdf_path.stem
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Use our existing splitting logic
            self.splitter.selected_file = str(pdf_path)
            self.splitter.output_dir = str(output_dir)
            self.splitter.process_pdf()
            
            # Clean up the original PDF
            pdf_path.unlink()
        except Exception as e:
            print(f"Error processing PDF {pdf_path}: {e}", file=sys.stderr)

class DienstplanPrinterService:
    """Main printer service class that handles printer installation and monitoring."""
    
    def __init__(self, printer_name: str = "Dienstplan Splitter", output_dir: Optional[Path] = None):
        """Initialize the printer service.
        
        Args:
            printer_name: Name of the virtual printer as it appears in Windows
            output_dir: Directory where split PDFs will be saved
        """
        self.printer_name = printer_name
        self.output_dir = output_dir or Path.home() / "Documents" / "Dienstplan Splitter"
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Create temp directory for initial PDFs
        self.temp_dir = self.output_dir / "temp"
        self.temp_dir.mkdir(exist_ok=True)
        
        # Set up file monitoring
        self.event_handler = PDFHandler(self.output_dir)
        self.observer = Observer()
        self.observer.schedule(self.event_handler, str(self.temp_dir), recursive=False)
        
    def install_printer(self) -> bool:
        """Install the printer using Windows Print to PDF driver.
        
        Returns:
            bool: True if installation was successful
        """
        try:
            # Get the Windows Print to PDF driver
            driver_name = "Microsoft Print to PDF"
            driver_info = win32print.GetPrinterDriverDirectory()
            
            # Add our printer port (points to temp directory)
            port_name = f"DS_PORT:{self.temp_dir}"
            
            # Create the printer
            win32print.AddPrinter(
                self.printer_name,
                2,  # Level 2 for detailed printer info
                {
                    "pDriverName": driver_name,
                    "pPrinterName": self.printer_name,
                    "pPortName": port_name,
                }
            )
            return True
        except Exception as e:
            print(f"Error installing printer: {e}", file=sys.stderr)
            return False
            
    def uninstall_printer(self) -> bool:
        """Remove the virtual printer from Windows.
        
        Returns:
            bool: True if uninstallation was successful
        """
        try:
            win32print.DeletePrinter(self.printer_name)
            return True
        except Exception as e:
            print(f"Error uninstalling printer: {e}", file=sys.stderr)
            return False
            
    def start_monitoring(self):
        """Start monitoring for new PDFs."""
        self.observer.start()
        
    def stop_monitoring(self):
        """Stop monitoring for new PDFs."""
        self.observer.stop()
        self.observer.join()
            
def main():
    """Main entry point when running as a service."""
    service = DienstplanPrinterService()
    if service.install_printer():
        print(f"Successfully installed printer: {service.printer_name}")
        try:
            service.start_monitoring()
            print("Monitoring for print jobs...")
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            service.stop_monitoring()
            service.uninstall_printer()
    else:
        print("Failed to install printer", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main() 