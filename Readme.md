# Pandas XL Writer functions
### Writing Pandas projects to an Excel workbook
***
This notebook contains Python functions useful for outputting an Excel, project workbook whose sheets are one or more Pandas DataFrames.  This is useful when a Python project builds up and completes several, related DataFrames whose data needs to be shared with a consulting client or other user.  Usage is to call the XLWriterPrep function for each DataFrame that will later be included in the Excel workbook.  The XLWriter function is then called at the end of the Python code to create the Excel workbook. The XLWriter functions facilitate using this for organized creation of a formatted workbook whose columns use specified number formats and which have specified column widths to control the Excel data appearance. XLWriter uses [Pandas ExcelWriter](https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.ExcelWriter.html). [Here is a useful ExcelWriter code example](https://xlsxwriter.readthedocs.io/example_pandas_column_formats.html) from the documentation.
​
The XLPrep Python function updates four lists used later by XLWriter.  XLPrep adds the DataFrame to a DataFrames list and adds a sheet name to a list of those. Two lists of lists hold formatting specifications for the columns on each sheet.  XLPrep adds a blank list for Excel number formats and a blank Excel column width list to these lists of lists for those items.  These list elements contain the same number of blank items as the DataFrame has columns.  The list elements can then be manually updated to specify Excel number formats and column widths for each DataFrame's columns.

J.D. Landgrebe,
October 25, 2019
