import os
from pathlib import Path
import pandas as pd

class FilesInfo:
    """
    Find all files in script folder path and subfolders. Store data to "result.xlsx".
    """
    data_list = []
    count = 0

    def __init__(self):
        self.script_folder_path = os.getcwd()

        
    def collect_object_info(self, path: str):
        """
        iterates through folders recursively and list of dictionaries
        """
        for object_name in os.listdir(path):
            path_to_object = f"{path}/{object_name}"
            if os.path.isfile(path_to_object):
                curr_object_name = os.path.basename(path_to_object)
                self.count += 1
                self.data_list.append(
                    {
                        "number": self.count, 
                        "folder": os.path.basename(path), 
                        "fileName": Path(curr_object_name).stem,
                        "extension": Path(curr_object_name).suffix
                    }
                )
            elif os.path.isdir(path_to_object):
                self.collect_object_info(path_to_object)

    def make_xlsx(self):
        """
        make xlsx file from list of dictionaries using pandas
        """
        data_frame = pd.DataFrame(self.data_list)
        data_frame.to_excel('result.xlsx', index=False)

if __name__ == "__main__":
    files_info_collector = FilesInfo()
    files_info_collector.collect_object_info(files_info_collector.script_folder_path)
    files_info_collector.make_xlsx()