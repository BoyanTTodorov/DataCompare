import os
import pandas as pd

path_protime = r'./Protime'
path_agency = r'./Agency'

class ConcatFiles:
    def __init__(self, path) -> None:
        """
        Initializes the class with the provided path.
        """
        self.path = path
        self.extension = '.xlsx'  # Corrected spelling
        self.df = None
        self.folder = None

    def combine_files(self):
        """
        Combines all the Excel sheets with the extension '.xlsx' into a single DataFrame
        from the folder passed as an argument.
        """
        try:
            # List all files in the directory with '.xlsx' extension
            self.folder = [file for file in os.listdir(self.path) if file.endswith(self.extension)]

            # Check if there are any files found
            if len(self.folder) == 1:
                # If there is only one file, read it into a DataFrame
                self.df = pd.read_excel(os.path.join(self.path, self.folder[0]))
            elif len(self.folder) > 1:
                # If there are multiple files, concatenate them into a single DataFrame
                self.df = pd.concat([pd.read_excel(os.path.join(self.path, file)) for file in self.folder], ignore_index=True)
            else:
                return f'No files with extension {self.extension} found in {self.path}'
        except Exception as e:
            print(f'Exception occure as {e}')
        finally:     
            print(f'Data loaded from:{self.path}')
            print("\n".join(f"{file}" for file in self.folder))

