import os
import sys
import time
from datetime import datetime
from pathlib import Path

import pyperclip

# Dynamic runtime validation for required dependencies.
REQUIRED_PACKAGES = ["pyperclip", "pymsgbox", "send2trash", "openpyxl", "ezsheets"]
for package_name in REQUIRED_PACKAGES:
    try:
        __import__(package_name)
    except ModuleNotFoundError:
        print(f"\n[CRITICAL ERROR] Required library missing: '{package_name}'.")
        print(f"Please run this command in your terminal: pip install {package_name}")
        input("\nPress Enter to close the system...")
        sys.exit()

import Logs
import pymsgbox
from Billing import Billing, SpreadsheetExport
from Function import Functions


class LFSystem(Functions):
    """Main class for the TUI ecosystem and unified operational control."""

    def __init__(self):
        self.drive = Path.home()
        self.system_folder = self.drive / "SISTEMA_LF"
        self.backups_folder = self.system_folder / "Backups"
        self.spreadsheets_folder = self.system_folder / "Spreadsheets"
        self.data_folder = self.system_folder / "Data"

        self.current_path = Path.cwd()
        self.pages = []
        self.current_page = 0

        self.backups_folder.mkdir(parents=True, exist_ok=True)
        self.data_folder.mkdir(parents=True, exist_ok=True)
        self.spreadsheets_folder.mkdir(parents=True, exist_ok=True)

    def modern_loading_screen(self):
        """Modern terminal loading bar."""
        os.system("cls" if os.name == "nt" else "clear")
        bar_width = 40
        print("\n" * 2)
        print(" INITIALIZING LF OPERATIONAL ECOSYSTEM ".center(65, "="))
        print("\n Mapping local buffers and starting core modules...\n")

        for index in range(bar_width + 1):
            percent = int((index / bar_width) * 100)
            filled = "■" * index
            empty = "-" * (bar_width - index)
            print(f"\r Syncing tables: |{filled}{empty}| {percent}% Complete", end="", flush=True)
            time.sleep(0.04)

        print("\n\n" + " INTEGRATED ENVIRONMENT READY ".center(65, "="))
        time.sleep(0.8)
        os.system("cls" if os.name == "nt" else "clear")

    def prepare_pages(self, target_path=None):
        """Prepare paginated directory entries. Accepts an alternate target path."""
        try:
            target = target_path if target_path is not None else self.current_path
            items = sorted([path for path in target.glob("*")], key=lambda path: path.name.lower())
            self.pages = [items[i : i + 10] for i in range(0, len(items), 10)]
        except Exception:
            self.pages = []

    def render_tui(self, custom_menu=None, custom_path=None):
        """Render the base TUI. Optional arguments support the spreadsheet module."""
        os.system("cls" if os.name == "nt" else "clear")
        display_address = custom_path if custom_path is not None else self.current_path

        print(f" FILE OPERATIONS LF | Time: {datetime.now().strftime('%H:%M:%S')}")
        print(f" Active Path: {display_address}")
        print(f" Item Page: {self.current_page + 1}/{max(1, len(self.pages))}")
        print("=" * 85)

        if not self.pages:
            print("   (Directory is empty or scan permissions are unavailable)")
        else:
            if self.current_page >= len(self.pages):
                self.current_page = 0
            for file_path in self.pages[self.current_page]:
                item_type = "[DIR]" if file_path.is_dir() else "[FILE]"
                print(f"  {item_type} {file_path.name}")

        print("=" * 85)
        if custom_menu:
            print(f" {custom_menu}")
        else:
            print(" [1] Next Page   | [2] Previous Page | [B] Parent Folder | [S] Zip Item (Clipboard)")
            print(" [G] Open Folder | [C] Backup        | [F] Copy Items    | [E] Billing | [P] Create Excel")
            print(" [ESC] End Session")
        print("=" * 85)

    def interface_loop(self):
        while True:
            self.prepare_pages(self.current_path)
            self.render_tui()

            command = input("Enter a panel command and press Enter: ").strip().lower()

            if command == "1" and self.current_page < len(self.pages) - 1:
                self.current_page += 1
            elif command == "2" and self.current_page > 0:
                self.current_page -= 1
            elif command == "b":
                self.current_path = self.current_path.parent
                self.current_page = 0
            elif command == "esc":
                print("\nDisconnecting from local storage and closing processes safely...")
                time.sleep(0.6)
                sys.exit()

            elif command == "g":
                print("\n>> OPEN FOLDER MODE: Copy the Windows folder name...")
                input("Press Enter here after copying the name:")
                clipboard_content = pyperclip.paste().strip()
                target = self.current_path / clipboard_content
                if target.exists() and target.is_dir():
                    self.current_path = target
                    self.current_page = 0
                else:
                    print("Target directory was not found in the active path.")
                    time.sleep(1.5)

            elif command == "s":
                print("\n>> ZIP MODE: Copy the exact file name to compress...")
                input("Press Enter here after copying the name:")
                clipboard_content = pyperclip.paste().strip()
                target = self.current_path / clipboard_content
                if target.exists():
                    if target.is_file():
                        self.create_zip_backup(target, self.backups_folder)
                    else:
                        print("The target is a folder. Use the dedicated [C] tool.")
                        time.sleep(1.5)

            elif command == "c":
                user_input = self.request_path("Enter the full path of the item to back up:")
                if user_input:
                    self.create_zip_backup(user_input, self.backups_folder)

            elif command == "f":
                choice = self.show_menu(
                    "File Replication Module",
                    {
                        "1": "Direct isolated copy (single item)",
                        "2": "Grouped smart copy (file tree)",
                    },
                )
                if choice:
                    source = self.request_path("SOURCE path:")
                    destination = self.request_path("DESTINATION path:")
                    if source and destination:
                        if choice == "1":
                            self.copy_single_item(source, destination)
                        elif choice == "2":
                            self.copy_folder_contents(source, destination)

            elif command == "e":
                billing = Billing(self.data_folder)
                while True:
                    option = self.show_menu(
                        "Audit and Accounting Panel",
                        {
                            "1": "Register New Account",
                            "2": "Show Account Metadata",
                            "3": "Update Internal Records",
                            "4": "Monitor Database",
                            "5": "Post Purchase Entry",
                            "6": "Reverse Entry by NC Code",
                        },
                    )
                    if option == "1":
                        billing.create_billing_account()
                    elif option == "2":
                        billing.show_account_details()
                    elif option == "3":
                        billing.edit_account()
                    elif option == "4":
                        sub_option = self.show_menu(
                            "Monitor Filter Selection",
                            {"1": "Active Customers", "2": "General Purchase History"},
                        )
                        if sub_option == "1":
                            billing.monitor_customers()
                        elif sub_option == "2":
                            billing.monitor_purchases()
                    elif option == "5":
                        account_name = pymsgbox.prompt("Linked Account Name:")
                        item = pymsgbox.prompt("Purchased Item:")
                        price = pymsgbox.prompt("Item Value:")
                        if account_name and item and price:
                            billing.add_purchase(account_name.strip(), item.strip(), price.strip())
                    elif option == "6":
                        billing.delete_purchase()
                    else:
                        break

            elif command == "p":
                spreadsheet_export = SpreadsheetExport(self.spreadsheets_folder, self.data_folder)
                spreadsheet_export.spreadsheet_interface_loop(self)

if __name__ == "__main__":
    system = LFSystem()
    system.modern_loading_screen()
    system.interface_loop()
