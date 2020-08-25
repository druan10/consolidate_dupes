import pandas as pd
from datetime import datetime
import os

if __name__ == "__main__":
    
    root_directory = './'
    file_name = 'your_file.xlsx'
    dupe_column = 'standardized_name'
    id_column = 'id'

    print('Loading report')
    
    report_df = pd.read_excel(os.path.join('.',file_name), na_values=['NA'])

    print('Loaded. Size of original file is:')
    print(report_df.shape)
    
    # Group by dupe_column
    # Get dupe counts, and group ids
    count_df = report_df.groupby(dupe_column).size().reset_index(name='dupe_count')

    dupe_item_ids_df = report_df.groupby([dupe_column])[id_column].apply(','.join).reset_index()
    print('Saving Report')

    writer = pd.ExcelWriter(f'./dupe_checks_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx')
    # Save full report in one sheet
    report_df.to_excel(writer,'Original')
    # Save dupe count to separate sheet
    count_df.to_excel(writer, "Dupe Count")
    # Save old ids to separate sheet
    dupe_item_ids_df.to_excel(writer, "Merged")

    writer.save()