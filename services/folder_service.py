import os
from datetime import datetime


class FolderService:
    @staticmethod
    def create_month_folder(base_path):
        now = datetime.now()
        folder_name = now.strftime("%B_%Y")
        full_path = os.path.join(base_path, folder_name)

        os.makedirs(full_path, exist_ok=True)
        return full_path
