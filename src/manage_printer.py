"""
Management script for the DienstplanSplitter virtual printer.
Provides commands to install, uninstall, and check the status of the printer.
Must be run as administrator on Windows.
"""

import sys
import ctypes
import argparse
from pathlib import Path
from printer_service.printer_service import DienstplanPrinterService

def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def install_printer():
    """Install the virtual printer."""
    service = DienstplanPrinterService()
    if service.install_printer():
        print("✅ Printer installed successfully!")
        print(f"Printer name: {service.printer_name}")
        print(f"Output directory: {service.output_dir}")
        return True
    else:
        print("❌ Failed to install printer")
        return False

def uninstall_printer():
    """Uninstall the virtual printer."""
    service = DienstplanPrinterService()
    if service.uninstall_printer():
        print("✅ Printer uninstalled successfully!")
        return True
    else:
        print("❌ Failed to uninstall printer")
        return False

def check_status():
    """Check if the printer is installed and working."""
    import win32print
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
    
    print("\n📊 Printer Status:")
    print(f"{'✅' if printer_exists else '❌'} Printer installed: {service.printer_name}")
    print(f"{'✅' if port_exists else '❌'} Port configured: {service.port_name}")
    print(f"{'✅' if output_dir_exists else '❌'} Output directory: {service.output_dir}")
    print(f"{'✅' if temp_dir_exists else '❌'} Temp directory: {service.temp_dir}")
    
    if printer_exists and port_exists and output_dir_exists and temp_dir_exists:
        print("\n✅ Printer is fully configured and ready to use!")
    else:
        print("\n⚠️  Some components are missing or not configured correctly")

def main():
    """Main entry point for the management script."""
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
        print("❌ This script must be run as administrator!")
        print("Please right-click and select 'Run as administrator'")
        return  # Changed from sys.exit(1) to allow pause at end
    
    # If no command provided, show help and install by default
    if not args.command:
        print("No command provided. Available commands:")
        parser.print_help()
        print("\nInstalling printer by default...")
        install_printer()
        return
    
    if args.command == 'install':
        install_printer()
    elif args.command == 'uninstall':
        uninstall_printer()
    elif args.command == 'status':
        check_status()

if __name__ == "__main__":
    try:
        main()
    finally:
        print("\nPress Enter to continue...")
        input() 