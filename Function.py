import logging
import os
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

import pymsgbox
import send2trash


class Functions:
    """File, backup, and input helpers for the LF system."""

    def delete_item(self, item_path):
        logging.info("Starting delete process for: %s", item_path)
        if not item_path.exists():
            pymsgbox.alert("The selected file or folder does not exist.", "Error")
            return False

        key = self.show_menu(
            "Delete Options",
            {
                "1": f"Delete permanently (no backup) -> {item_path.name}",
                "2": f"Send to the Windows Recycle Bin -> {item_path.name}",
            },
        )

        try:
            if key == "1":
                confirmation = pymsgbox.prompt(
                    text=(
                        "WARNING: This action cannot be undone.\n\n"
                        f"To permanently delete:\n'{item_path.name}'\n"
                        "type DELETE below:"
                    ),
                    title="Critical Confirmation",
                )
                if confirmation and confirmation.strip() == "DELETE":
                    if item_path.is_file():
                        item_path.unlink()
                    elif item_path.is_dir():
                        shutil.rmtree(item_path)
                    logging.info("Item permanently deleted: %s", item_path.name)
                    pymsgbox.alert(f"'{item_path.name}' was permanently deleted.", "Success")
                    return True

                pymsgbox.alert("Action canceled.", "Canceled")
                return False

            if key == "2":
                send2trash.send2trash(item_path)
                logging.info("Item sent to Recycle Bin: %s", item_path.name)
                pymsgbox.alert(f"'{item_path.name}' was sent to the Recycle Bin.", "Success")
                return True

        except Exception as exc:
            logging.error("Error deleting item %s: %s", item_path, exc, exc_info=True)
            pymsgbox.alert(f"Error while deleting: {exc}", "Error")
            return False

        return False

    def generate_timestamp(self):
        return datetime.now().strftime("%m-%d-%y_%H-%M-%S")

    def copy_single_item(self, source, destination):
        logging.info("Simple copy: %s -> %s", source, destination)
        try:
            if source.is_dir():
                shutil.copytree(source, destination / source.name, dirs_exist_ok=True)
            else:
                shutil.copy2(source, destination)
            return True
        except Exception as exc:
            logging.error("Simple copy failed: %s", exc, exc_info=True)
            pymsgbox.alert(f"Copy failed: {exc}", "Error")
            return False

    def copy_folder_contents(self, source, destination):
        logging.info("Smart copy: %s -> %s", source, destination)
        try:
            items = list(source.glob("*"))
            if not items:
                return False

            for item in items:
                destination_path = destination / item.name
                if item.is_file():
                    shutil.copy2(item, destination_path)
                elif item.is_dir():
                    shutil.copytree(item, destination_path, dirs_exist_ok=True)
            return True
        except Exception as exc:
            logging.error("Smart copy failed: %s", exc, exc_info=True)
            pymsgbox.alert(f"Smart copy failed: {exc}", "Error")
            return False

    def create_zip_backup(self, item_path, backups_folder):
        zip_name = f"Backup_{item_path.stem}_{self.generate_timestamp()}.zip"
        final_zip_path = backups_folder / zip_name

        try:
            with zipfile.ZipFile(final_zip_path, "w") as zip_file:
                if item_path.is_file():
                    zip_file.write(
                        item_path,
                        arcname=item_path.name,
                        compress_type=zipfile.ZIP_DEFLATED,
                        compresslevel=9,
                    )
                else:
                    for root_folder, _, files in os.walk(item_path):
                        for file_name in files:
                            full_path = Path(root_folder) / file_name
                            relative_path = full_path.relative_to(item_path.parent)
                            zip_file.write(
                                full_path,
                                arcname=relative_path,
                                compress_type=zipfile.ZIP_DEFLATED,
                                compresslevel=9,
                            )
            pymsgbox.alert(f"Backup created:\n{zip_name}", "Success")
            return True
        except Exception as exc:
            logging.error("Zip backup failed: %s", exc, exc_info=True)
            pymsgbox.alert(f"Backup failed: {exc}", "Error")
            return False

    def request_path(self, message):
        while True:
            user_input = pymsgbox.prompt(
                text=f"{message}\nor type 'S' to exit",
                title="Path Input",
            )
            if user_input is None or user_input.strip().lower() == "s":
                return None

            path = Path(user_input.strip().replace('"', ""))
            if path.exists():
                return path

            pymsgbox.alert("Invalid path or path not found.", "Error")

    def show_menu(self, title, options):
        """Terminal-based menu helper."""
        print(f"\n==== {title.upper()} ====")
        for key, description in options.items():
            print(f" [{key}] {description}")
        print(" [S] Exit / Cancel")
        print("======================")

        while True:
            choice = input("Choose an option and press Enter: ").strip().lower()
            if choice == "s":
                return None
            if choice in [key.lower() for key in options.keys()]:
                return choice.upper() if choice.isalpha() else choice
            print("Invalid option. Please try again.")
