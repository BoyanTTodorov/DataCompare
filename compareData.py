import pandas as pd

class AgencyData:
    def __init__(self, agency_df):
        """
        Initializes the class with the provided DataFrame for Agency.
        """
        self.agency_df = agency_df

        # Dictionary to rename columns from Dutch to English
        self.column_renames = {
            'Invoice nr': 'Invoice Number',
            'Kostenplaats site': 'Cost Center',
            'Naam Medewerker': 'Employee Name',
            'Personeels nummer': 'Employee Number',
            'Toeslag': 'Allowance',
            'Uren': 'Hours',
            'Subtotaal excl Reiskosten': 'Subtotal (Excl. Travel Costs)',
            'Reiskosten': 'Travel Costs',
            'Tarief': 'Rate'
        }

    def rename_columns(self):
        """
        Renames the columns in the Agency DataFrame to English,
        while keeping the 'Datum', 'Gewerkte week', and 'Nettowaarde' columns unchanged.
        """
        try:
            self.agency_df.rename(columns=self.column_renames, inplace=True)
            print("Columns renamed successfully, except for 'Datum', 'Gewerkte week', and 'Nettowaarde'.")
        except Exception as e:
            print(f"Error renaming columns: {e}")

    def fix_date_format(self):
        """
        Fixes the date format in the 'Datum' column to 'yyyy-mm-dd'.
        """
        try:
            self.agency_df['Datum'] = pd.to_datetime(self.agency_df['Datum']).dt.strftime('%Y-%m-%d')
            print("Date format fixed to yyyy-mm-dd.")
        except Exception as e:
            print(f"Error fixing date format: {e}")

    def get_agency_data(self):
        """
        Returns the agency DataFrame after renaming columns and fixing date format.
        """
        return self.agency_df
    
    
class ProtimeData:
    def __init__(self, protime_df):
        """
        Initializes the class with the provided DataFrame for Protime.
        """
        self.protime_df = protime_df

        # Dictionary to rename columns in Protime DataFrame
        self.column_renames = {
            'Date': 'Date',
            'Week': 'Week',
            'Temp Agency': 'Temp Agency',
            'Full Name': 'Full Name',
            'Invoice incl ADV': 'Invoice incl ADV'
        }

    def rename_columns(self):
        """
        Renames the columns in the Protime DataFrame.
        """
        try:
            self.protime_df.rename(columns=self.column_renames, inplace=True)
            print("Columns renamed successfully.")
        except Exception as e:
            print(f"Error renaming columns: {e}")

    def fix_date_format(self):
        """
        Fixes the date format in the 'Date' column to 'yyyy-mm-dd'.
        """
        try:
            self.protime_df['Date'] = pd.to_datetime(self.protime_df['Date']).dt.strftime('%Y-%m-%d')
            print("Date format fixed to yyyy-mm-dd.")
        except Exception as e:
            print(f"Error fixing date format: {e}")

    def filter_temp_agency(self):
        """
        Filters the Protime DataFrame to include only rows where 'Temp Agency' equals 'OTTO'.
        """
        try:
            self.protime_df = self.protime_df[self.protime_df['Temp Agency'] == 'OTTO']
            print("Filtered Protime data for Temp Agency 'OTTO'.")
        except Exception as e:
            print(f"Error filtering by Temp Agency: {e}")

    def get_protime_data(self):
        """
        Returns the Protime DataFrame after renaming columns, fixing date format, and filtering by 'OTTO'.
        """
        return self.protime_df
