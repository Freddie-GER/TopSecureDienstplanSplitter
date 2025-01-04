"""
Management script for the DienstplanSplitter virtual printer.
Provides commands to install, uninstall, and check the status of the printer.
Must be run as administrator on Windows.
"""

import sys
import ctypes
import argparse
import logging
from pathlib import Path
from printer_service.printer_service import DienstplanPrinterService

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

def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def install_printer():
    """Install the virtual printer."""
    logger.info("Installing printer...")
    service = DienstplanPrinterService()
    if service.install_printer():
        logger.info("✅ Printer installed successfully!")
        logger.info(f"Printer name: {service.printer_name}")
        logger.info(f"Output directory: {service.output_dir}")
        return True
    else:
        logger.error("❌ Failed to install printer")
        return False

def uninstall_printer():
    """Uninstall the virtual printer."""
    logger.info("Uninstalling printer...")
    service = DienstplanPrinterService()
    if service.uninstall_printer():
        logger.info("✅ Printer uninstalled successfully!")
        return True
    else:
        logger.error("❌ Failed to uninstall printer")
        return False

def check_status():
    """Check if the printer is installed and working."""
    import win32print
    logger.info("Checking printer status...")
    service = DienstplanPrinterService()
    
    # Check if printer exists
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
    printer_exists = any(p[2] == service.printer_name for p in printers)
    
    # Check if port exists
    ports = win32print.EnumPorts(None, 1)
    port_exists = any(p[1] == service.port_name for p in ports)
    
    # Check output directory
    output_dir_exists = service.output_dir.exists()
    temp_dir_exists = service.temp_dir.exists()
    
    logger.info("\n📊 Printer Status:")
    logger.info(f"{'✅' if printer_exists else '❌'} Printer installed: {service.printer_name}")
    logger.info(f"{'✅' if port_exists else '❌'} Port configured: {service.port_name}")
    logger.info(f"{'✅' if output_dir_exists else '❌'} Output directory: {service.output_dir}")
    logger.info(f"{'✅' if temp_dir_exists else '❌'} Temp directory: {service.temp_dir}")
    
    if printer_exists and port_exists and output_dir_exists and temp_dir_exists:
        logger.info("\n✅ Printer is fully configured and ready to use!")
    else:
        logger.warning("\n⚠️  Some components are missing or not configured correctly")

def main():
    """Main entry point for the management script."""
    try:
        logger.info("Starting printer management script")
        parser = argparse.ArgumentParser(
            description="Manage the DienstplanSplitter virtual printer"
        )
        parser.add_argument(
            'command',
            nargs='?',  # Make command optional
            choices=['install', 'uninstall', 'status'],
            help='Command to execute'
        )
        
        args = parser.parse_args()
        
        # Check for admin rights
        if not is_admin():
            logger.error("❌ This script must be run as administrator!")
            logger.error("Please right-click and select 'Run as administrator'")
            return False
        
        # If no command provided, show help and install by default
        if not args.command:
            logger.info("No command provided. Available commands:")
            parser.print_help()
            logger.info("\nInstalling printer by default...")
            return install_printer()
        
        if args.command == 'install':
            return install_printer()
        elif args.command == 'uninstall':
            return uninstall_printer()
        elif args.command == 'status':
            check_status()
            return True
            
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        return False

if __name__ == "__main__":
    try:
        success = main()
    except Exception as e:
        logger.error(f"Critical error: {e}", exc_info=True)
        success = False
    finally:
        # Clean up logging handlers
        for handler in logging.getLogger().handlers[:]:
            handler.flush()
            handler.close()
        if not success:
            print("\nPress Enter to continue...")
            input() 