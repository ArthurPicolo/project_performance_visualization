import pandas as pd

def process_dataframe(df, column_mappings):
    """
    Replaces names in specified columns of DataFrames using unique mappings and prefixes.
    """
    for column, prefix in column_mappings.items():
        # Get unique names from the column
        unique_names = df[column].dropna().unique()
        
        # Create mapping
        name_mapping = {
            name: f"{prefix} {i + 1}" if not str(name).startswith("TOTAL") else name
            for i, name in enumerate(unique_names)
        }
        
        # Apply mapping to the column
        df[column] = df[column].map(name_mapping)

    return df

def main():
    # Load Excel files
    pnl = pd.read_excel('P L By Account Manager _PL By AM.xlsx')
    project_list = pd.read_excel('ProjectListing.xlsx')

    # Define column mappings and prefixes for each DataFrame
    pnl_mappings = {
        "PROJECT NAME": "Project",
        "ACCOUNT MANAGER": "Account Manager",
        "CUSTOMER": "Customer"
    }
    project_list_mappings = {
        "Project Name": "Project",
        "Customer": "Customer",
        "Project Manager": "Project Manager"
    }

    # Process the DataFrames separately
    pnl_processed = process_dataframe(pnl, pnl_mappings)
    project_list_processed = process_dataframe(project_list, project_list_mappings)

    # Save the processed DataFrames
    pnl_processed.to_excel('Processed_PNL.xlsx', index=False)
    project_list_processed.to_excel('Processed_ProjectListing.xlsx', index=False)

    print("Processed files saved:")
    print("- Processed_PNL.xlsx")
    print("- Processed_ProjectListing.xlsx")

if __name__ == '__main__':
    main()
