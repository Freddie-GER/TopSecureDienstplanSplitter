"""
DienstplanSplitter Virtual Printer Service

This module implements a virtual printer service that:
1. Creates a redirected printer port
2. Installs a printer using a standard Windows driver
3. Captures and converts print jobs to PDF
4. Automatically processes them with DienstplanSplitter logic
"""

import os
import sys
import time
import shutil
import win32print
import win32api
import win32con
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
        
        # Port and driver names
        self.port_name = "DSPLITTER:"
        self.driver_name = "Generic / Text Only"
        
    def create_port(self) -> bool:
        """Create a redirected printer port.
        
        Returns:
            bool: True if port creation was successful
        """
        try:
            # Get the port monitor
            monitors = win32print.EnumMonitors(None, 1)
            monitor_name = None
            for monitor in monitors:
                if "Redirected Port" in monitor[1]:
                    monitor_name = monitor[1]
                    break
                    
            if not monitor_name:
                print("Error: Redirected Port monitor not found", file=sys.stderr)
                return False
                
            # Create port info structure
            port_info = {
                "Port": self.port_name,
                "Monitor": monitor_name,
                "Description": "Dienstplan Splitter Port",
                "Command": f'cmd.exe /c copy /b "%%1" "{self.temp_dir}\\%%2.pdf"'
            }
            
            # Add the port
            win32print.AddPort(None, None, port_info)
            return True
        except Exception as e:
            print(f"Error creating port: {e}", file=sys.stderr)
            return False
        
    def install_printer(self) -> bool:
        """Install the printer using a standard Windows driver.
        
        Returns:
            bool: True if installation was successful
        """
        try:
            # First create our port
            if not self.create_port():
                return False
                
            # Get system printer driver directory
            driver_dir = win32print.GetPrinterDriverDirectory()
            
            # Create printer
            printer_info = {
                "pServerName": None,
                "pPrinterName": self.printer_name,
                "pShareName": "",
                "pPortName": self.port_name,
                "pDriverName": self.driver_name,
                "pComment": "Dienstplan Splitter Virtual Printer",
                "pLocation": "",
                "pDevMode": None,
                "pSepFile": "",
                "pPrintProcessor": "winprint",
                "pDatatype": "RAW",
                "pParameters": "",
                "pSecurityDescriptor": None,
                "Attributes": win32print.PRINTER_ATTRIBUTE_LOCAL | win32print.PRINTER_ATTRIBUTE_SHARED,
                "Priority": 1,
                "DefaultPriority": 1,
                "StartTime": 0,
                "UntilTime": 0,
                "Status": 0,
                "cJobs": 0,
                "AveragePPM": 0
            }
            
            # Add the printer
            handle = win32print.AddPrinter(None, 2, printer_info)
            if handle:
                win32print.ClosePrinter(handle)
                return True
            return False
            
        except Exception as e:
            print(f"Error installing printer: {e}", file=sys.stderr)
            return False
            
    def uninstall_printer(self) -> bool:
        """Remove the virtual printer from Windows.
        
        Returns:
            bool: True if uninstallation was successful
        """
        try:
            # Delete printer
            handle = win32print.OpenPrinter(self.printer_name)
            if handle:
                win32print.DeletePrinter(handle)
                win32print.ClosePrinter(handle)
            
            # Delete port
            win32print.DeletePort(None, None, self.port_name)
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