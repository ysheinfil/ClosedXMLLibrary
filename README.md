# ClosedXMLLibrary

### Description
- This solution contains a console application project called XslxToCsv. When you run the executable, it takes all the Excel files (that are in xlsx format) within a folder and copies each one of them to csv format. You have the ability to require certain columns and ignore certain columns. 
- The following components are expected to be in the app.config file:
  - `xlsxFolder` - Folder that contains the Excel (xlsx) files.
  - `csvFolder` - Folder that you want the csv files to go.
  - `NecessaryFields` - list of column names that must be in the Excel files. Each name is separated by a comma.
  - `SkippedFields` - list of column names in the Excel files that will be skipped and will not end up in the csv file. Each name is separated by a comma.
