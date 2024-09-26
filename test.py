import pandas as pd
import os
import time
import numpy as np
import re

# Function to normalize and simplify names
def normalize_name(name):
    if isinstance(name, str):
        # Convert to lowercase
        name = name.lower()
        # Remove non-alphabetic characters (excluding spaces and hyphens)
        name = re.sub(r'[^a-z\s\-]', '', name)
        # Remove extra whitespace
        name = ' '.join(name.split())
        # Remove middle names or initials
        parts = name.split()
        if len(parts) > 2:
            name = f"{parts[0]} {parts[-1]}"
        return name
    else:
        # Return empty string for non-string values (e.g., NaN)
        return ''

# DataLoader class to load and clean data
class DataLoader():
    def __init__(self, path) -> None:
        self.path = path
        self.df = None

    def load_data(self):
        folder = os.listdir(self.path)
        print(f'List of the data found in {folder}')
        try:
            files = [file for file in folder if file.endswith('.xlsx')]
            if len(files) == 1:
                self.df = pd.read_excel(os.path.join(self.path, files[0]))
                print('One file loaded')
            elif len(files) > 1:
                data_frames = [pd.read_excel(os.path.join(self.path, file)) for file in files]
                self.df = pd.concat(data_frames, ignore_index=True)
                print(f'{len(files)} files loaded')
            else:
                print('There are no files with .xlsx extension')
                self.df = pd.DataFrame()
            return self.df
        except Exception as e:
            print(f'Exception occurred: {e}')
            return pd.DataFrame()
        finally:
            print(f'Dataframe created from {self.path}')

    def clean_protime_data(self):
        try:
            self.df = self.df[pd.to_numeric(self.df['Week'], errors='coerce').notna()]
            self.df = self.df[['Date', 'Week', 'Temp Agency', 'Full Name', 'Hours (Dec)', 'Invoice incl ADV']]
            self.df['Hours (Dec)'] = self.df['Hours (Dec)'].astype(float).round(2)
            self.df['Week'] = self.df['Week'].astype(int)
            filter_mask = (self.df['Temp Agency'] == 'OTTO') & (self.df['Week'].between(3, 37))
            self.df = self.df[filter_mask]
            self.df['Date'] = pd.to_datetime(self.df['Date']).dt.strftime('%Y-%m-%d')
            self.df['Invoice incl ADV'] = pd.to_numeric(self.df['Invoice incl ADV'], errors='coerce').round(2)
            self.df['Full Name'] = self.df['Full Name'].astype(str).str.strip()
            self.df = self.df.reset_index(drop=True)
            return self.df
        except Exception as e:
            print(f'Exception occurred: {e}')
            return pd.DataFrame()
        finally:
            print(f'Protime data is prepared')

    def clean_agency_data(self):
        column_renames = {
            'Datum': 'Date',
            'Gewerkte week': 'Week',
            'Naam Medewerker': 'Name',
            'Uren': 'Hours',
            'Nettowaarde': 'Invoice'
        }
        try:
            self.df = self.df.rename(columns=column_renames)
            self.df = self.df[pd.to_numeric(self.df['Week'], errors='coerce').notna()]
            self.df = self.df[['Date', 'Week', 'Name', 'Hours', 'Invoice']]
            self.df['Hours'] = self.df['Hours'].astype(float).round(2)
            self.df['Week'] = self.df['Week'].astype(int)
            self.df['Date'] = pd.to_datetime(self.df['Date']).dt.strftime('%Y-%m-%d')
            self.df['Invoice'] = pd.to_numeric(self.df['Invoice'], errors='coerce').round(2)
            self.df['Name'] = self.df['Name'].astype(str).str.strip()
            self.df = self.df.reset_index(drop=True)
            return self.df
        except Exception as e:
            print(f'Exception occurred: {e}')
            return pd.DataFrame()
        finally:
            print(f'Agency data is prepared')

if __name__ == '__main__':
    # Start timer
    start_time = time.time()

    # Load agency data
    path_agency = r'C:\Users\boyan.todorov\Desktop\DataCompareProject\Agency'
    agency_data = DataLoader(path_agency)
    agency = agency_data.load_data()
    if agency is not None and not agency.empty:
        agency = agency_data.clean_agency_data()
    else:
        print("Agency data is empty or None. Exiting.")
        exit()

    # Load protime data
    path_protime = r'C:\Users\boyan.todorov\Desktop\DataCompareProject\Protime'
    protime_data = DataLoader(path_protime)
    protime = protime_data.load_data()
    if protime is not None and not protime.empty:
        protime = protime_data.clean_protime_data()
    else:
        print("Protime data is empty or None. Exiting.")
        exit()

    # Proceed only if both DataFrames are not empty
    if not protime.empty and not agency.empty:
        # Normalize names in both DataFrames
        protime['Normalized Name'] = protime['Full Name'].apply(normalize_name)
        agency['Normalized Name'] = agency['Name'].apply(normalize_name)

        # Remove entries with empty 'Normalized Name'
        protime = protime[protime['Normalized Name'] != '']
        agency = agency[agency['Normalized Name'] != '']

        # Aggregate hours by Date, Week, and Normalized Name
        protime_agg = protime.groupby(['Date', 'Week', 'Normalized Name'], as_index=False)['Hours (Dec)'].sum()
        agency_agg = agency.groupby(['Date', 'Week', 'Normalized Name'], as_index=False)['Hours'].sum()

        # Merge the aggregated DataFrames
        merged_df = pd.merge(
            protime_agg,
            agency_agg,
            how='outer',
            on=['Date', 'Week', 'Normalized Name'],
            suffixes=('_Protime', '_Agency')
        )

        # Calculate the absolute difference in hours and convert to minutes
        merged_df['Hours_Diff_Hours'] = merged_df['Hours (Dec)'] - merged_df['Hours']
        merged_df['Hours_Diff_Minutes'] = abs(merged_df['Hours_Diff_Hours']) * 60

        # Introduce a threshold in minutes
        threshold_minutes = 5  # You can adjust this value as needed

        # Filter the differences based on the threshold
        diff_df = merged_df[merged_df['Hours_Diff_Minutes'] > threshold_minutes]

        # Merge back with original names for clarity
        # Get unique combinations of Date, Week, Normalized Name, and Full Name
        protime_names = protime[['Date', 'Week', 'Normalized Name', 'Full Name']].drop_duplicates()
        agency_names = agency[['Date', 'Week', 'Normalized Name', 'Name']].drop_duplicates()

        # Merge to get the original names
        diff_only_columns = diff_df.merge(
            protime_names,
            on=['Date', 'Week', 'Normalized Name'],
            how='left'
        ).merge(
            agency_names,
            on=['Date', 'Week', 'Normalized Name'],
            how='left',
            suffixes=('_Protime', '_Agency')
        )

        # Rearrange columns
        diff_only_columns = diff_only_columns[['Date', 'Week', 'Full Name', 'Name', 'Hours (Dec)', 'Hours', 'Hours_Diff_Minutes']]

        # Display the differences
        print(diff_only_columns)
        diff_only_columns.to_excel('diff.xlsx', index=False)
    else:
        print("One or both of the DataFrames are empty. Cannot proceed with merging.")

    # End timer and print elapsed time
    end_time = time.time()
    print(f"Time taken for name matching and merging: {end_time - start_time} seconds")
