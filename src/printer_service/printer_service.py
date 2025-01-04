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
import logging
from pathlib import Path
from typing import Optional
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

def setup_logging():
    """Set up logging configuration."""
    logger = logging.getLogger("DienstplanSplitter")
    logger.setLevel(logging.DEBUG)
    
    # Create formatters
    detailed_formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s"
    )
    
    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(detailed_formatter)
    logger.addHandler(console_handler)
    
    # File handler - will be set up when we have the output directory
    return logger

# Create logger instance
logger = setup_logging()

# Import our existing splitting logic
try:
    from gui.dienstplan_splitter_gui import DienstplanSplitter  # Try direct import first
except ImportError:
    try:
        from src.gui.dienstplan_splitter_gui import DienstplanSplitter  # Try with src prefix
    except ImportError:
        logger.error("Could not import DienstplanSplitter. Please check the installation.")
        sys.exit(1)

class PDFHandler(FileSystemEventHandler):
    """Handles PDF files created by Windows Print to PDF."""
    
    def __init__(self, output_dir: Path):
        self.output_dir = output_dir
        # Initialize our existing splitter logic
        self.splitter = DienstplanSplitter(None)  # None because we don't need GUI
        
    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith('.pdf'):
            logger.info(f"New PDF detected: {event.src_path}")
            # Wait a moment to ensure the file is completely written
            time.sleep(1)
            self.process_pdf(Path(event.src_path))
            
    def process_pdf(self, pdf_path: Path):
        """Process a newly created PDF using our existing splitting logic."""
        try:
            logger.info(f"Processing PDF: {pdf_path}")
            # Create output directory
            output_dir = self.output_dir / pdf_path.stem
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Use our existing splitting logic
            self.splitter.selected_file = str(pdf_path)
            self.splitter.output_dir = str(output_dir)
            self.splitter.process_pdf()
            
            # Clean up the original PDF
            pdf_path.unlink()
            logger.info(f"Successfully processed PDF: {pdf_path}")
        except Exception as e:
            logger.error(f"Error processing PDF {pdf_path}: {e}", exc_info=True)

class DienstplanPrinterService:
    """Main printer service class that handles printer installation and monitoring."""
    
    def __init__(self, output_dir: str = None):
        """Initialize the printer service."""
        logger.info("==================================================")
        logger.info("Starting DienstplanSplitter Virtual Printer")
        logger.info("==================================================")
        
        # Set up the output directory
        if output_dir is None:
            # Default to Desktop/Pläne_YYYY_MM_DD
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            current_date = datetime.now().strftime("%Y_%m_%d")
            output_dir = os.path.join(desktop, f"Pläne_{current_date}")
        
        self.output_dir = output_dir
        logger.info(f"Output directory: {self.output_dir}")
        
        # Create directories
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Set up file logging now that we have the output directory
        log_file = os.path.join(self.output_dir, "printer_service.log")
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logger.addHandler(file_handler)
        
        # Create temp directory
        self.temp_dir = os.path.join(self.output_dir, "temp")
        logger.info(f"Temp directory: {self.temp_dir}")
        os.makedirs(self.temp_dir, exist_ok=True)
        
        # Printer configuration
        self.printer_name = "Dienstplan Splitter"
        self.port_name = "IP_Dienstplan_Splitter"
        self.driver_name = "Microsoft XPS Document Writer"
        logger.info(f"Using printer driver: {self.driver_name}")
        logger.info(f"Installing printer: {self.printer_name}")
        
        # Set up file monitoring
        self.event_handler = PDFHandler(self.output_dir)
        self.observer = Observer()
        self.observer.schedule(self.event_handler, str(self.temp_dir), recursive=False)
        self.observer.start()
        
    def list_available_drivers(self):
        """List all available printer drivers on the system."""
        try:
            drivers = win32print.EnumPrinterDrivers(None, None, 2)
            logger.info("Available printer drivers:")
            for driver in drivers:
                logger.info(f"  - {driver['Name']}")
            return drivers
        except Exception as e:
            logger.error(f"Error listing drivers: {e}", exc_info=True)
            return []
        
    def create_port(self) -> bool:
        """Create a redirected printer port."""
        logger.info(f"Creating printer port: {self.port_name}")
        try:
            # First remove any existing port
            try:
                logger.info("Attempting to remove existing port...")
                # Use EnumPorts to check if port exists
                ports = win32print.EnumPorts(None, 1)
                if any(p[1] == self.port_name for p in ports):
                    # Port exists, but we can't delete it, so we'll reuse it
                    logger.info("Port already exists, will reuse it")
                    return True
            except Exception as e:
                logger.debug(f"Port check failed: {e}")
            
            # Add the port using AddPortEx
            try:
                # Get the port monitor
                monitors = win32print.EnumMonitors(None, 1)
                local_monitor = next((m for m in monitors if m[1].lower() == "local port"), None)
                
                if not local_monitor:
                    logger.error("Local Port monitor not found")
                    return False
                
                # Create the port using the monitor
                handle = win32print.OpenPrinter(None)
                try:
                    win32print.AddPortEx(None, local_monitor[1], 1, self.port_name)
                    logger.info("Port created successfully")
                    return True
                finally:
                    win32print.ClosePrinter(handle)
                    
            except Exception as e:
                logger.error(f"Error creating port: {e}")
                return False
                
        except Exception as e:
            logger.error(f"Error managing port: {e}", exc_info=True)
            return False
        
    def install_printer(self) -> bool:
        """Install the virtual printer."""
        try:
            # Create the port first
            if not self.create_port():
                logger.error("Failed to create printer port")
                return False

            # Remove any existing printer with the same name
            try:
                win32print.DeletePrinter(win32print.OpenPrinter(self.printer_name))
                logger.info("Removed existing printer")
            except Exception as e:
                logger.debug(f"Printer removal failed (this is normal if it didn't exist): {e}")

            # Get the driver info
            try:
                drivers = win32print.EnumPrinterDrivers(None, None, 2)
                driver = next((d for d in drivers if d["Name"].lower() == self.driver_name.lower()), None)
                
                if not driver:
                    logger.error(f"Driver '{self.driver_name}' not found")
                    return False
                
                # Create printer using PRINTER_INFO_2 structure
                printer_info = {
                    "pServerName": None,
                    "pPrinterName": self.printer_name,
                    "pShareName": None,
                    "pPortName": self.port_name,
                    "pDriverName": driver["Name"],
                    "pComment": "Virtual printer for Dienstplan Splitter",
                    "pLocation": None,
                    "pDevMode": None,
                    "pSepFile": None,
                    "pPrintProcessor": "winprint",
                    "pDatatype": "RAW",
                    "pParameters": None,
                    "pSecurityDescriptor": None,
                    "Attributes": 0x00000002,  # PRINTER_ATTRIBUTE_LOCAL
                    "Priority": 0,
                    "DefaultPriority": 0,
                    "StartTime": 0,
                    "UntilTime": 0,
                }
                
                win32print.AddPrinter(None, 2, printer_info)
                logger.info("Printer installed successfully")
                return True
                
            except Exception as e:
                logger.error(f"Error installing printer: {e}", exc_info=True)
                return False
                
        except Exception as e:
            logger.error(f"Error in printer installation: {e}", exc_info=True)
            return False
            
    def uninstall_printer(self) -> bool:
        """Remove the virtual printer from Windows."""
        logger.info(f"Uninstalling printer: {self.printer_name}")
        try:
            # Delete printer
            try:
                logger.info("Attempting to remove printer...")
                handle = win32print.OpenPrinter(self.printer_name)
                if handle:
                    win32print.DeletePrinter(handle)
                    win32print.ClosePrinter(handle)
                    logger.info("Printer removed successfully")
            except Exception as e:
                logger.error(f"Error removing printer: {e}", exc_info=True)
            
            # Delete port
            try:
                logger.info("Attempting to remove port...")
                win32print.DeletePort(None, None, self.port_name)
                logger.info("Port removed successfully")
            except Exception as e:
                logger.error(f"Error removing port: {e}", exc_info=True)
                
            return True
        except Exception as e:
            logger.error(f"Error during uninstallation: {e}", exc_info=True)
            return False
            
    def start_monitoring(self):
        """Start monitoring for new PDFs."""
        logger.info("Starting file monitoring")
        self.observer.start()
        
    def stop_monitoring(self):
        """Stop monitoring for new PDFs."""
        logger.info("Stopping file monitoring")
        self.observer.stop()
        self.observer.join()
            
def main():
    """Main entry point when running as a service."""
    try:
        logger.info("Starting DienstplanSplitter printer service")
        service = DienstplanPrinterService()
        if service.install_printer():
            logger.info(f"Successfully installed printer: {service.printer_name}")
            try:
                service.start_monitoring()
                logger.info("Monitoring for print jobs...")
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                logger.info("Received shutdown signal")
                service.stop_monitoring()
                service.uninstall_printer()
        else:
            logger.error("Failed to install printer")
            return False
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        return False
    finally:
        # Clean up logging handlers
        for handler in logger.handlers[:]:
            handler.flush()
            handler.close()
            logger.removeHandler(handler)
        return True

if __name__ == "__main__":
    success = main()
    if not success:
        print("\nPress Enter to exit...")
        input() 