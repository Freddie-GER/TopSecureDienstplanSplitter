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

# Set up logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('printer_service.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

logger.info("="*50)
logger.info("Starting DienstplanSplitter Virtual Printer")
logger.info("="*50)

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
    
    def __init__(self, printer_name: str = "Dienstplan Splitter", output_dir: Optional[Path] = None):
        """Initialize the printer service."""
        logger.info("Initializing DienstplanPrinterService")
        self.printer_name = printer_name
        
        # Create output directory on desktop with today's date
        today = datetime.now()
        desktop = Path(os.path.expandvars("%USERPROFILE%\\Desktop"))
        self.output_dir = desktop / f"Pläne_{today.strftime('%Y_%m_%d')}"
        self.output_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"Output directory: {self.output_dir}")
        
        # Create temp directory for initial PDFs
        self.temp_dir = self.output_dir / "temp"
        self.temp_dir.mkdir(exist_ok=True)
        logger.info(f"Temp directory: {self.temp_dir}")
        
        # Set up file monitoring
        self.event_handler = PDFHandler(self.output_dir)
        self.observer = Observer()
        self.observer.schedule(self.event_handler, str(self.temp_dir), recursive=False)
        
        # Port and driver names
        self.port_name = "IP_" + self.printer_name.replace(" ", "_")  # Standard port naming convention
        self.driver_name = "Generic / Text Only"  # Standard Windows driver that's always available
        logger.info(f"Using printer driver: {self.driver_name}")
        
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
                win32print.DeletePort(None, None, self.port_name)
                logger.info("Existing port removed successfully")
            except Exception as e:
                logger.debug(f"Port deletion failed (this is normal if it didn't exist): {e}")
            
            # Add the port using the Local Port monitor
            win32print.AddPort(None, None, {
                "PortName": self.port_name,
                "MonitorName": "Local Port",
                "Description": "Dienstplan Splitter Port"
            })
            logger.info("Port created successfully")
            
            return True
        except Exception as e:
            logger.error(f"Error creating port: {e}", exc_info=True)
            return False
        
    def install_printer(self) -> bool:
        """Install the printer using a standard Windows driver."""
        logger.info(f"Installing printer: {self.printer_name}")
        try:
            # First create our port
            if not self.create_port():
                return False
                
            # Remove any existing printer
            try:
                logger.info("Attempting to remove existing printer...")
                handle = win32print.OpenPrinter(self.printer_name)
                if handle:
                    win32print.DeletePrinter(handle)
                    win32print.ClosePrinter(handle)
                    logger.info("Existing printer removed successfully")
            except Exception as e:
                logger.debug(f"Printer deletion failed (this is normal if it didn't exist): {e}")
            
            # Create printer using level 2 info structure
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
                "Attributes": win32print.PRINTER_ATTRIBUTE_LOCAL | win32print.PRINTER_ATTRIBUTE_QUEUED,
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
                logger.info(f"Successfully installed printer: {self.printer_name}")
                return True
                
            logger.error("Failed to get printer handle")
            return False
            
        except Exception as e:
            logger.error(f"Error installing printer: {e}", exc_info=True)
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