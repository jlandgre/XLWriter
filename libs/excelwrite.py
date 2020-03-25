import pandas as pd
import numpy as np
PathCIFile = 'col_info.csv'

#Loads new item into lists of DataFrames, sheet names.  Initializes column formats
def XLWriterPrep(lst_dfs, lst_shts, lst_fmts, lst_colwidths, df,sht):
    lst_fmts.append([])
    lst_colwidths.append([])
    for i in range(len(df.index.names) + len(df.columns)):
        lst_fmts[len(lst_fmts) - 1].append('')
        lst_colwidths[len(lst_colwidths) - 1].append(0)
    return lst_dfs, lst_shts, lst_fmts, lst_colwidths

# Write list of DataFrames to Excel workbook as separate worksheets
def XLWriter(wkbk, lst_dfs, lst_shts, lst_fmts, lst_colwidths):
    writer = pd.ExcelWriter(wkbk, engine='xlsxwriter')
    worksheet = []
    workbook = writer.book
    for i in range(len(lst_dfs)):
        lst_dfs[i].to_excel(writer, sheet_name=lst_shts[i])
        worksheet.append(writer.sheets[lst_shts[i]])

    #Add all uniqueformats to a dict
    dict_fmts = {}
    format = []
    k = 0
    for i in range(len(lst_fmts)):
        for j in range(len(lst_fmts[i])):
            curfmt = lst_fmts[i][j]
            if len(curfmt) > 0 and curfmt not in dict_fmts:
                dict_fmts[curfmt] = k #Save the index, k, as dictionary value for later
                format.append(workbook.add_format({'num_format': curfmt}))
                k += 1

    #Assign specified formats and column widths to each sheet
    for i in range(len(lst_shts)):

        #create pd.ExcelWriter object for each sheet
        worksheet = writer.sheets[lst_shts[i]]

        #Assign any specified column widths and number formats
        for j in range(1,len(lst_fmts[i])):
            colstr = XLColString(j + 1)
            colwidth = None
            fmt = None

            if lst_colwidths[i][j] > 0: colwidth = lst_colwidths[i][j]
            if len(lst_fmts[i][j]) > 0: fmt = lst_fmts[i][j]

            if fmt != None:
                worksheet.set_column(colstr, colwidth, format[dict_fmts[fmt]])
            else:
                worksheet.set_column(colstr, colwidth, None)

    #Write the workbook and return
    writer.save()
    return()

#Converts integer, icol, into Excel column range (Example: icol = 30 --> 'AD:AD')
def XLColString(icol):
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    letters =''
    while True:
        letter = ''
        Q = (icol - 1) // 26
        R = (icol - 1) % 26
        letter = alphabet[R]
        letters = letter + letters
        if Q == 0: break
        icol = Q
    return letters + ':' + letters

#Global Variables
NameCI = 'Name'
DescCI = 'Description'
UnitsCI = 'Units'
XLFormatCI = 'XLFormat'
XLWidthCI = 'XLWidth'
ParentCI = 'Parent'

path_ColInfo = 'libs/'

#ReadColInfo - reads the col_info DataFrame to a nested dictionary
#Column 1 of the col_info.csv contains column headings used as key for the other column info items (For example
#data column 'Revenue' has Description, XLFormat and XLWidth equal to 'Annual Revenue', '$0.00' and '11',
#respectively)
def ReadColInfoFromFile(PathCIFile):

    #Read the file, and Create the nested dictionary empty strucure
    df_CI = pd.read_csv(PathCIFile, index_col=NameCI)
    ColInfo = {DescCI: {}, UnitsCI: {}, XLFormatCI: {}, XLWidthCI: {}}

    #Read each col_info DataFrame row and translate to dictionary values
    for var_name, row in df_CI.iterrows():
        for dictCI in ColInfo:
            ColInfo[dictCI][var_name] = row[dictCI]
    return ColInfo

#RefreshColInfoToFile - refreshes col_info.csv based on dictionary contents; adds file rows as needed
def RefreshColInfoToFile(PathCIFile, ColInfo):

    #Read the col_info file to be updated
    df_CI = pd.read_csv(PathCIFile, index_col=NameCI)

    #Iterate over dictionaries in ColInfo and over variables (keys); update/append to df_CI
    for CI_dict in ColInfo:
        for var in CI_dict:
            df_CI.loc[var,CI_dict] = ColInfo[CI_dict][var]
    df_CI.to_csv(PathCIFile)
    return

#BuildXLWriterLists - Uses number format and column width ColInfo to build needed XLWriter lists
def BuildXLWriterLists(df, ColInfo):

    #Initialize lists with placeholder for index - Works for single index
    lst_fmts = ['0']
    lst_widths = [10]

    #Populate format and width for index; xxx Need to address case of multiindex
    if df.index.names[0] is None:
        df.index.name = 'index'
    elif df.index.name in ColInfo:
        lst_fmts = [ColInfo[XLFormatCI][df.index.name]]
        lst_widths = [ColInfo[XLWidthCI][df.index.name]]

    #Iterate through df columns and add number format and column width to the lists
    for col in df.columns:
        if col in ColInfo[XLFormatCI] and col in ColInfo[XLWidthCI]:
            lst_fmts.append(ColInfo[XLFormatCI][col])
            lst_widths.append(ColInfo[XLWidthCI][col])
        else:
            lst_fmts.append('0')
            lst_widths.append(10)

    #Get rid of quotation marks that prevent formats from working with ExcelWriter
    lst_fmts = ListReplaceNaN(lst_fmts,'')
    lst_fmts = [s.replace('"', '') for s in lst_fmts]
    lst_widths = ListReplaceNaN(lst_widths,0)
    return lst_fmts, lst_widths

#Toggles val to blank string if val is nan values
def SetVal(val):
    retval = val
    if isinstance(val , float) and np.isnan(val): retval = ''
    return str(retval)

def CreateExcelStepsDF(lst_dfs, lst_XLshts, ColInfo):

    #Add needed ColInfo entries for ExcelSteps worksheet
    ColInfo[XLFormatCI]['Formula/List Name/Sort-by'] = '@'
    ColInfo[XLFormatCI]['Number Format'] = '@'

    dictC = {1:'Sheet', 2:'Column',3:'Step',4:'Comment',5:'Number Format',6:'Width'}
    df_ExcelSteps = pd.DataFrame(columns=['Sheet', 'Column', 'Step', 'Formula/List Name/Sort-by',
                                          'After or End Column', 'Keep Formulas', 'Comment',
                                          'Number Format', 'Width'])
    for dframe, sht in zip(lst_dfs, lst_XLshts):
        lst_dfCols = list(dframe.index.names) + dframe.columns.tolist()

        #Iterate through columns and append ExcelSteps recipe rows
        for i, col in enumerate(lst_dfCols):
            if not col in ColInfo[DescCI]: continue
            comment = SetVal(ColInfo[DescCI][col])
            ucomment = SetVal(ColInfo[UnitsCI][col])
            if len(ucomment) > 0: comment = comment + ' in ' + ucomment

            row = {dictC[1]:sht, dictC[2]:col, dictC[3]:'Col_Format', dictC[4]:comment,
                   dictC[5]: ColInfo[XLFormatCI][col],dictC[6]: ColInfo[XLWidthCI][col]}

            df_ExcelSteps = df_ExcelSteps.append(row, ignore_index=True)

        #Add Tbl_FreezeRow1 to end of recipe
        df_ExcelSteps = df_ExcelSteps.append({dictC[1]:np.nan}, ignore_index=True)
        freeze = {dictC[1]:sht, dictC[3]:'Tbl_FreezeRow1'}
        df_ExcelSteps = df_ExcelSteps.append(freeze, ignore_index=True)

        #add blank row to create spacing in the recipe
        df_ExcelSteps = df_ExcelSteps.append({dictC[1]:np.nan}, ignore_index=True)

        #Name the ExcelSteps df's index
        df_ExcelSteps.index.name = 'row'

    lst_dfs.append(df_ExcelSteps)
    lst_XLshts.append('ExcelSteps')
    return lst_dfs, lst_XLshts, ColInfo

def WriteExcelWorkbook(lst_dfs, lst_shts, fName_xlsx, ColInfo, IsExcelSteps):
    lst_fmts = []
    lst_colwidths = []
    retval = fName_xlsx + ' Written Successfully'
    if len(lst_dfs) != len(lst_shts): return 'ERROR: Must be same number of dfs and shts'

    #Make local copy so ColInfo doesn't get modified
    Col_Info_l = ColInfo

    if IsExcelSteps:
        lst_dfs, lst_shts, Col_Info_l = CreateExcelStepsDF(lst_dfs, lst_shts, Col_Info_l)

    i=0
    for dframe, sht in zip(lst_dfs, lst_shts):
        XLWriterPrep(lst_dfs, lst_shts, lst_fmts, lst_colwidths, dframe, sht)
        lst_fmts[i], lst_colwidths[i] = BuildXLWriterLists(dframe, ColInfo)
        i = i+1

    #Text format for ExcelSteps Formula and Number Format columns
    if IsExcelSteps:
        i = len(lst_dfs) - 1
        lst_fmts[i][4], lst_fmts[i][8] = '@', '@'
    XLWriter(fName_xlsx, lst_dfs, lst_shts, lst_fmts, lst_colwidths)
    return retval

def ListReplaceNaN(lst, val_replace):
    for i, v in enumerate(lst):
        if v is np.nan: lst[i] = val_replace
    return lst
