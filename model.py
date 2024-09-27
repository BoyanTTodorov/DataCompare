import pandas as pd
import os
import numpy as np
import re

class DataModel:
    def __init__(self):
        """
        Initializes the DataModel with empty paths and dataframes.
        The threshold for filtering based on differences in minutes can be set later.
        """
        self.path_agency = ''
        self.path_protime = ''
        self.threshold_minutes = None  # Optional threshold for filtering differences
        self.start_week = None
        self.end_week = None
        self.protime_data = pd.DataFrame()
        self.agency_data = pd.DataFrame()
        self.differences = pd.DataFrame()

    # Setters for paths and parameters
    def set_agency_path(self, path):
        """
        Set the path for agency data and validate it.
        """
        self.path_agency = path
        print(f"Model: Agency path set to {path}")

    def set_protime_path(self, path):
        """
        Set the path for Protime data and validate it.
        """
        self.path_protime = path
        print(f"Model: Protime path set to {path}")

    def set_threshold_minutes(self, minutes):
        """
        Set the threshold in minutes for filtering output data based on differences.
        If invalid input is provided, raise a ValueError.
        """
        try:
            if minutes is not None and minutes != '':
                self.threshold_minutes = float(minutes)
            else:
                self.threshold_minutes = None  # Optional
            print(f"Model: Threshold minutes set to {self.threshold_minutes}")
        except ValueError:
            raise ValueError("Invalid input for threshold minutes. Please enter a numeric value.")

    def set_week_range(self, start_week, end_week):
        """
        Set the week range for filtering data.
        Both weeks must be valid numbers. Otherwise, raise an error.
        """
        try:
            self.start_week = int(start_week) if start_week else None
            self.end_week = int(end_week) if end_week else None

            # Ensure the weeks are valid numbers
            if self.start_week is None or self.end_week is None:
                raise ValueError("Start week and End week cannot be empty.")

            print(f"Model: Week range set from {self.start_week} to {self.end_week}")
        except ValueError:
            raise ValueError("Invalid week numbers. Please enter numeric values for both start and end weeks.")

    def normalize_name(self, name):
        """
        Normalize the given name by:
        - Converting to lowercase.
        - Removing any special characters (except hyphens and spaces).
        - Removing extra spaces.
        - Removing middle names (including hyphenated ones) to standardize names for comparison.
        """
        if isinstance(name, str):
            # Convert the name to lowercase and normalize spacing
            name = name.lower()
            name = name.replace('-', ' ')
            name = ' '.join(name.split())
            # Remove middle names by keeping only the first and last part
            parts = name.split()
            if len(parts) > 2:
                name = f"{parts[0]} {parts[-1]}"
            return name
        else:
            return ''  # Return empty string for invalid names

    def load_data(self):
        """
        Load data from the provided paths for both agency and Protime datasets.
        Raise errors if paths are invalid or files are missing.
        """
        try:
            print("Model: Loading agency data")
            # Load Agency Data
            if not os.path.exists(self.path_agency):
                raise FileNotFoundError(f"Agency path does not exist: {self.path_agency}")
            agency_files = [f for f in os.listdir(self.path_agency) if f.endswith('.xlsx')]
            if not agency_files:
                raise FileNotFoundError(f"No .xlsx files found in agency path: {self.path_agency}")
            agency_dfs = [pd.read_excel(os.path.join(self.path_agency, f)) for f in agency_files]
            self.agency_data = pd.concat(agency_dfs, ignore_index=True)
            print(f"Model: Agency data loaded with {len(self.agency_data)} records")
            self.clean_agency_data()

            print("Model: Loading protime data")
            # Load Protime Data
            if not os.path.exists(self.path_protime):
                raise FileNotFoundError(f"Protime path does not exist: {self.path_protime}")
            protime_files = [f for f in os.listdir(self.path_protime) if f.endswith('.xlsx')]
            if not protime_files:
                raise FileNotFoundError(f"No .xlsx files found in protime path: {self.path_protime}")
            protime_dfs = [pd.read_excel(os.path.join(self.path_protime, f)) for f in protime_files]
            self.protime_data = pd.concat(protime_dfs, ignore_index=True)
            print(f"Model: Protime data loaded with {len(self.protime_data)} records")
            self.clean_protime_data()

            return True
        except Exception as e:
            print(f"Model: Error loading data - {e}")
            raise e

    def clean_protime_data(self):
        """
        Clean the Protime data by:
        - Filtering valid weeks.
        - Renaming columns for clarity.
        - Applying week and other necessary filters.
        """
        try:
            print("Model: Cleaning protime data")
            df = self.protime_data.copy()
            df = df[pd.to_numeric(df['Week'], errors='coerce').notna()]
            df = df.rename(columns={
                'Hours (Dec)': 'Protime Hours',
                'Invoice incl ADV': 'Protime Invoice'
            })
            df = df[['Week', 'Temp Agency', 'Full Name', 'Protime Hours', 'Protime Invoice']]
            df['Protime Hours'] = df['Protime Hours'].astype(float).round(2)
            df['Protime Invoice'] = pd.to_numeric(df['Protime Invoice'], errors='coerce').round(2)
            df['Week'] = df['Week'].astype(int)

            # Apply week filter if set
            if self.start_week and self.end_week:
                df = df[(df['Week'] >= self.start_week) & (df['Week'] <= self.end_week)]
                print(f"Protime data week range after filtering: {df['Week'].min()} to {df['Week'].max()}")
            # Filter by Temp Agency (e.g., 'OTTO')
            df = df[df['Temp Agency'] == 'OTTO']
            df['Full Name'] = df['Full Name'].astype(str).str.strip()
            self.protime_data = df.reset_index(drop=True)
            print(f"Model: Protime data cleaned with {len(df)} records")
        except Exception as e:
            print(f"Model: Error cleaning protime data - {e}")
            raise Exception(f'Exception occurred while cleaning Protime data: {e}')

    def clean_agency_data(self):
        """
        Clean the Agency data by:
        - Filtering valid weeks.
        - Renaming columns for clarity.
        - Applying week and other necessary filters.
        """
        try:
            print("Model: Cleaning agency data")
            df = self.agency_data.copy()
            column_renames = {
                'Gewerkte week': 'Week',
                'Naam Medewerker': 'Name',
                'Uren': 'Agency Hours',
                'Nettowaarde': 'Agency Invoice'
            }
            df = df.rename(columns=column_renames)
            df = df[pd.to_numeric(df['Week'], errors='coerce').notna()]
            df = df[['Week', 'Name', 'Agency Hours', 'Agency Invoice']]
            df['Agency Hours'] = df['Agency Hours'].astype(float).round(2)
            df['Agency Invoice'] = pd.to_numeric(df['Agency Invoice'], errors='coerce').round(2)
            df['Week'] = df['Week'].astype(int)

            # Apply week filter if set
            if self.start_week and self.end_week:
                df = df[(df['Week'] >= self.start_week) & (df['Week'] <= self.end_week)]
                print(f"Agency data week range after filtering: {df['Week'].min()} to {df['Week'].max()}")
            df['Name'] = df['Name'].astype(str).str.strip()
            self.agency_data = df.reset_index(drop=True)
            print(f"Model: Agency data cleaned with {len(df)} records")
        except Exception as e:
            print(f"Model: Error cleaning agency data - {e}")
            raise Exception(f'Exception occurred while cleaning Agency data: {e}')

    def process_data(self):
        """
        Process the data by:
        - Normalizing names.
        - Filtering weeks that exist in both datasets.
        - Aggregating and merging the datasets.
        - Applying thresholds for filtering.
        - Providing useful feedback and error handling to the user.
        """
        try:
            print("Model: Processing data")

            if self.protime_data.empty or self.agency_data.empty:
                raise Exception("One or both of the DataFrames are empty. Cannot proceed with merging.")

            # Normalize names in both datasets
            self.protime_data['Normalized Name'] = self.protime_data['Full Name'].apply(self.normalize_name)
            self.agency_data['Normalized Name'] = self.agency_data['Name'].apply(self.normalize_name)

            # Remove entries with empty 'Normalized Name'
            self.protime_data = self.protime_data[self.protime_data['Normalized Name'] != '']
            self.agency_data = self.agency_data[self.agency_data['Normalized Name'] != '']

            # Check if the week range exists in both datasets
            protime_weeks = set(self.protime_data['Week'].unique())
            agency_weeks = set(self.agency_data['Week'].unique())
            common_weeks = protime_weeks.intersection(agency_weeks)

            # If no common weeks exist, inform the user and display week ranges from both datasets
            if not common_weeks:
                min_protime_week = self.protime_data['Week'].min()
                max_protime_week = self.protime_data['Week'].max()
                min_agency_week = self.agency_data['Week'].min()
                max_agency_week = self.agency_data['Week'].max()
                raise Exception(
                    f"No matching weeks found between datasets. "
                    f"Protime weeks range from {min_protime_week} to {max_protime_week}, "
                    f"while Agency weeks range from {min_agency_week} to {max_agency_week}. "
                    f"Please check your week selection."
                )

            # Filter the data to include only matching weeks in the selected range
            protime_agg = self.protime_data[self.protime_data['Week'].between(self.start_week, self.end_week)]
            agency_agg = self.agency_data[self.agency_data['Week'].between(self.start_week, self.end_week)]

            if protime_agg.empty or agency_agg.empty:
                raise Exception(
                    "No data available for the selected week range. "
                    f"Please select a range between Protime weeks {min_protime_week}-{max_protime_week} and "
                    f"Agency weeks {min_agency_week}-{max_agency_week}."
                )

            # Proceed with aggregation and comparison
            protime_agg = protime_agg.groupby(['Week', 'Normalized Name'], as_index=False).agg({
                'Protime Hours': 'sum',
                'Protime Invoice': 'sum'
            })

            agency_agg = agency_agg.groupby(['Week', 'Normalized Name'], as_index=False).agg({
                'Agency Hours': 'sum',
                'Agency Invoice': 'sum'
            })

            # Merge the aggregated DataFrames
            merged_df = pd.merge(
                protime_agg,
                agency_agg,
                how='outer',
                on=['Week', 'Normalized Name'],
                suffixes=('_Protime', '_Agency')
            )

            # Calculate the differences
            merged_df['Hours Difference'] = merged_df['Protime Hours'] - merged_df['Agency Hours']
            merged_df['Difference in Minutes'] = abs(merged_df['Hours Difference']) * 60
            merged_df['Invoice Difference'] = merged_df['Protime Invoice'] - merged_df['Agency Invoice']
            merged_df['Overpay Request'] = np.where((merged_df['Protime Invoice'] - merged_df['Agency Invoice']) < 0, abs(merged_df['Protime Invoice'] - merged_df['Agency Invoice']), 0)
            merged_df['GXO Overpays'] = np.where((merged_df['Protime Invoice'] - merged_df['Agency Invoice']) > 0, abs(merged_df['Protime Invoice'] - merged_df['Agency Invoice']), 0)
            # Apply threshold if it is set
            conditions = pd.Series(True, index=merged_df.index)
            if self.threshold_minutes is not None and self.threshold_minutes > 0:
                conditions &= (merged_df['Difference in Minutes'] > self.threshold_minutes)
            diff_df = merged_df[conditions]

            # **Add Totals Row**
            totals = diff_df[['Protime Hours', 'Agency Hours',
                              'Protime Invoice', 'Agency Invoice',
                              'Hours Difference', 'Difference in Minutes', 'Invoice Difference', 'Overpay Request', 'GXO Overpays']].sum(numeric_only=True)
            totals_row = pd.DataFrame({
                'Week': ['Total'],
                'Normalized Name': [''],
                'Protime Hours': [totals['Protime Hours']],
                'Agency Hours': [totals['Agency Hours']],
                'Protime Invoice': [totals['Protime Invoice']],
                'Agency Invoice': [totals['Agency Invoice']],
                'Hours Difference': [totals['Hours Difference']],
                'Difference in Minutes': [totals['Difference in Minutes']],
                'Invoice Difference': [totals['Invoice Difference']],
                'Overpay Request': [totals['Overpay Request']],
                'GXO Overpays': [totals['GXO Overpays']]

            })

            # Append the totals row to the DataFrame
            self.differences = pd.concat([diff_df, totals_row], ignore_index=True)
            print(f"Model: Data processing completed with {len(self.differences)} records (including totals)")
            return True
        except Exception as e:
            print(f"Model: Error processing data - {e}")
            raise e

    def save_report(self, output_path):
        """
        Save the report to the specified output path.
        """
        try:
            print(f"Model: Saving report to {output_path}")
            self.differences.to_excel(output_path, index=False)
            print("Model: Report saved successfully")
        except Exception as e:
            print(f"Model: Error saving report - {e}")
            raise e
