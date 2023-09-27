from Pandas import PandasObject


if __name__ == "__main__":
    
    project1_obj = PandasObject('your_excel_sheet.xlsm')
    df1 = project1_obj.read_excel_data(sheet_name='Sheet1', header_rows=24, columns_range='C:U,Z') # Change header_row and coulmn_range according to your preference
    project1_obj.write_to_database(df1, databaseName='my_sqlite_database1', sheet_name='Sheet1')

    project2_obj = PandasObject('your_excel_sheet_2.xlsm')
    df2 = project2_obj.read_excel_data(sheet_name='Sheet1', header_rows=24, columns_range='C:U,Z')
    project2_obj.write_to_database(df2, databaseName='my_sqlite_database2', sheet_name='Sheet1')
  
    primary_db = 'my_sqlite_database1.db'  
    secondary_db = 'my_sqlite_database2.db'  


    project1_obj.compare_differences(primary_db, secondary_db)

   