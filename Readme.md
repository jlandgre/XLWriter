# Pandas XL Writer functions
### Writing Pandas projects to an Excel workbook

The Jupyter Notebook in this repository contains Python functions useful for outputting an Excel, project workbook whose sheets are one or more Pandas DataFrames.  This is useful when a Python project builds several, related DataFrames whose data needs to be shared with a consulting client or other user.  Usage is to call the XLWriterPrep function for each DataFrame that will later be included in the Excel workbook.  The XLWriter function is then called at the end of the Python code to create the Excel workbook. The XLWriter functions facilitate creating a workbook whose columns use specified number formats and which have specified column widths to control the Excel data appearance. XLWriter uses [Pandas ExcelWriter](https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.ExcelWriter.html). [The XLWriter functions build on this published ExcelWriter code example](https://xlsxwriter.readthedocs.io/example_pandas_column_formats.html) from the documentation.
​
The XLWriterPrep function updates four lists for later use by the XLWriter function.  XLWriterPrep adds a specified DataFrame to a DataFrames list and adds a sheet name to a list of those. Two lists of lists hold formatting specifications for the columns on each sheet.  XLWriterPrep adds a blank list for a shhet's Excel number formats and a blank Excel column width list with zero values as null defaults.  These list elements contain the same number of blank items as the DataFrame has columns.  The example Jupyter Notebook shows how to manuallly update list elements to specify Excel number formats and column widths for each DataFrame's columns.

J.D. Landgrebe,

October 25, 2019
