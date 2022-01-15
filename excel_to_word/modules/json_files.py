import json
import os


class JSON:
    """Class for read and write json-files."""
    def __init__(self, file_name: str):
        """
        Initialisation all parameters.
        :param file_name: *.json
        """
        self.file_name = file_name
    
    def read(self) -> dict:
        """
        Read data from file.
        :return: dict data
        """
        check_file = os.path.exists(self.file_name)
        if not check_file:
            with open(self.file_name, 'w', encoding="utf-8") as file:
                file.write('{}')
        with open(self.file_name, 'r') as file:
            data = json.load(file)
        return data
    
    def write(self, data: dict):
        """
        Write dict data to file.
        :param data: dict with data
        :return: None
        """
        check_file = os.path.exists(self.file_name)
        if not check_file:
            with open(self.file_name, 'w', encoding="utf-8") as file:
                file.write('{}')
        with open(self.file_name, "w", encoding="utf-8") as file:
            json.dump(data,
                      file,
                      ensure_ascii=False,
                      sort_keys=False,
                      indent=2,
                      separators=(',', ': ')
                      )
