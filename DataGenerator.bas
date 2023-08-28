' Use this script to generate data from excel file and convert it into test scripts
' Procedurtes to run from the UI:
'
' MakeDataExtract
'   Extract data into XML file.
'   XML file will be saved in the same folder as original Excel file with name <filename>.testdata.xml
'
' MakeDataExtractAndTransform
'   Extracts data and transforms using the stylesheet
'
' FileList_MakeDataExtractAndTransform
'   Extracts data and transforms using list of stylesheets

'   Default stylesheet name is:
Const DEFAULT_STYLESHEET_NAME = "TestData_to_DMCTestSQL.xsl"
'   File will be saved under the same folder with the name <file name>.testclass.sql

Const DEFAULT_FILE_EXTENSION = ".testclass.sql"

Const DEFAULT_DATA_EXTENSION = ".TestData.xml"

Const IS_EXPORT_DATA_EXTTRACT = "1" ' Flag to export intermediate data.xml files (1/0)

Const MAX_BLANK_LINES_BETWEEN_BLOCKS = 2

'-------------------------------------------------------------------------------
Sub MakeDataExtract()
    Dim WB As Workbook
    Dim Data As String
    
    'MsgBox "MakeDataExtract" '----------
    
    Set WB = Application.ActiveWorkbook
    
    Data = "<?xml version='1.0' encoding='UTF-8'?>"
    
    AppendLine Data, ""
    AppendLine Data, GetWorkbookData(WB)
    
    ExportFileName = WB.FullName
    
    ExportFileName = Mid(ExportFileName, 1, InStr(ExportFileName, ".") - 1)
    
    ExportFileName = ExportFileName + DEFAULT_DATA_EXTENSION    '.TestData.xml
    
    Open ExportFileName For Output As #1
    Print #1, Data
    Close #1
    
    'MsgBox "done save - " + ExportFileName '----------
    
End Sub

'-------------------------------------------------------------------------------
Sub MakeDataExtractAndTransformOne()
    Dim xml As DOMDocument60
    Dim xslt As DOMDocument60
    Dim WB As Workbook
    Dim W As Workbook
    Dim Data As String
    Dim StylesheetName As String
    Dim FileExtension As String
    Dim N As name
    Dim R As Range
        
    'MsgBox "MakeDataExtractAndTransformOne" '----------
    
    Set WB = Application.ActiveWorkbook
    StylesheetName = DEFAULT_STYLESHEET_NAME
    FileExtension = DEFAULT_FILE_EXTENSION
    
    On Error Resume Next
    
    Set N = WB.Names("stylesheet_name")
    Set R = N.RefersToRange
    If Not (R Is Nothing) Then
        ssName = R.value
        If ssName <> "" Then
            StylesheetName = ssName
        End If
    End If
    
    'MsgBox "StylesheetName - " + StylesheetName '----------
    
    On Error GoTo 0
    
    On Error Resume Next
    
    Set N = WB.Names("ouput_file_extenstion")
    Set R = N.RefersToRange
    If Not (R Is Nothing) Then
        feName = R.value
        If feName <> "" Then
            FileExtension = feName
        End If
    End If
    
    'MsgBox "FileExtension - " + FileExtension '----------
    
    On Error GoTo 0
    
    
    Data = "<?xml version='1.0' encoding='UTF-8'?>"
    
    AppendLine Data, ""
    AppendLine Data, GetWorkbookData(WB)
    
    '---------------------------------------------------------------------------
    If IS_EXPORT_DATA_EXTTRACT = "1" Then
    
        ExportFileName = WB.FullName
        
        ExportFileName = Mid(ExportFileName, 1, InStr(ExportFileName, ".") - 1)
        
        ExportFileName = ExportFileName + DEFAULT_DATA_EXTENSION    '.TestData.xml
        
        Open ExportFileName For Output As #1
        Print #1, Data
        Close #1
        
        'MsgBox "done save - " + ExportFileName '----------
        
    End If
    '---------------------------------------------------------------------------
    
    Set xml = New DOMDocument60
    
    If xml.LoadXML(Data) Then
        
        Set xslt = New DOMDocument60
        
        xsltPath = Application.ThisWorkbook.Path
        
        If xslt.Load(xsltPath + "/" + StylesheetName) Then
            
            Data = xml.transformNode(xslt)
            
            ExportFileName = WB.FullName
            
            ExportFileName = Mid(ExportFileName, 1, InStr(ExportFileName, ".") - 1)
            
            ExportFileName = ExportFileName + FileExtension
            
            Open ExportFileName For Output As #1
            Print #1, Data
            Close #1
            
            'MsgBox "done save - " + ExportFileName '----------
    
        End If
    End If
    
End Sub


'-------------------------------------------------------------------------------
Sub MakeDataExtractAndTransform()
    Dim xml As DOMDocument60
    Dim xslt As DOMDocument60
    Dim WB As Workbook
    Dim Opts As Worksheet
    Dim Data As String
    Dim StylesheetName As String
    Dim FileExtension As String
    Dim FileName As String
    Dim N As name
    Dim R As Range
    Dim LO As ListObject
            
    'MsgBox "MakeDataExtractAndTransform" '----------
    
    Set WB = Application.ActiveWorkbook
    
    Set Opts = WB.Sheets("Options")
    
    On Error Resume Next
    
    Set LO = Opts.ListObjects("TransformationOptions")
    
    '---------------------------------------------------------------------------
    'If Not (LO Is Nothing) Then
    '    MsgBox "NOT LO Is Nothing"
    'Else
    '    MsgBox "LO Is Nothing"
    'End If
    '---------------------------------------------------------------------------
    
    
    On Error GoTo 0
   
    If Not (LO Is Nothing) Then
        Data = "<?xml version='1.0' encoding='UTF-8'?>"
        
        AppendLine Data, ""
        AppendLine Data, GetWorkbookData(WB)
            
            
        '---------------------------------------------------------------------------
        If IS_EXPORT_DATA_EXTTRACT = "1" Then
        
            ExportFileName = WB.FullName
            
            ExportFileName = Mid(ExportFileName, 1, InStr(ExportFileName, ".") - 1)
            
            ExportFileName = ExportFileName + DEFAULT_DATA_EXTENSION    '.data.xml
            
            Open ExportFileName For Output As #1
            Print #1, Data
            Close #1
            
            'MsgBox "done save - " + ExportFileName '----------
            
        End If
        '---------------------------------------------------------------------------
        
        Set xml = New DOMDocument60
        
        If xml.LoadXML(Data) Then
            Set xslt = New DOMDocument60
            
            xsltPath = Application.ThisWorkbook.Path
            
            
            'MsgBox "xsltPath - " + xsltPath '----------
            'MsgBox "LO.ListRows.Count - " + Str(LO.ListRows.Count) '----------
            
            For idx = 1 To LO.ListRows.Count
                
                'MsgBox "idx - " + Str(idx) '----------
                StylesheetName = LO.ListRows(idx).Range(1) 'Name of XSLT Template
                FileExtension = LO.ListRows(idx).Range(2) 'Output file extention
				
                FileName = "" '(Optional) Output Export Path and FileName
                If Not (LO.ListRows(idx).Range(3) Is Nothing) Then
                   ssName = LO.ListRows(idx).Range(3).value
                   If ssName <> "" Then
                       FileName = ssName
                   End If
                End If
                
                If xslt.Load(xsltPath + "/" + StylesheetName) Then
                
                    Data = xml.transformNode(xslt)
                    
                    'MsgBox "do transformNode - " + xsltPath + "/" + StylesheetName '----------
                    
                    ExportFileName = WB.FullName
                        
						
                    If FileName = "" Then
                        
                        'Replace File Extention
                        
                        ExportFileName = Mid(ExportFileName, 1, InStr(ExportFileName, ".") - 1)
                        
                        ExportFileName = ExportFileName + FileExtension
                        
                    Else
                        
                        'Replace File Name
                        
                        ExportFileName = Mid(ExportFileName, 1, InStrRev(ExportFileName, "\"))
                        ExportFileName = ExportFileName + FileName
                        
                    End If
                    
                    'MsgBox "ExportFileName - " + ExportFileName '----------
                    
                    Open ExportFileName For Output As #1
                    Print #1, Data
                    Close #1
                    
                    'MsgBox "done save - " + ExportFileName '----------
                
                End If
            Next idx
        End If
    Else
        MakeDataExtractAndTransformOne
    End If

End Sub

'-------------------------------------------------------------------------------
Sub FileList_MakeDataExtractAndTransform()
    Dim WB As Workbook
    Dim SList As Worksheet
    Dim aName As String
    
    
    Set SList = Application.ActiveWorkbook.ActiveSheet
    
    row = 1
    cnt = 0
    cntTotal = 0
    While SList.Cells(row, 1).value <> "" Or row = 1
        aName = SList.Cells(row, 1).value
        'MsgBox "aName - " + aName '----------
                
        On Error Resume Next
        If aName > "" Then
            Set WB = Application.Workbooks.Open(aName)
            cntTotal = cntTotal + 1
            If Not (WB Is Empty) Then
                WB.Activate
                MakeDataExtractAndTransform
                WB.Close
                cnt = cnt + 1
            End If
        End If
        On Error GoTo 0
        
        row = row + 1
    Wend
    
    MsgBox Str(cnt) + " out of " + Str(cntTotal) + " tests has been processed"
End Sub


'-------------------------------------------------------------------------------
Function GetWorkbookData(WB As Workbook) As String
    Dim S As Worksheet
    Dim OptionsWS As Worksheet
    Dim Res As String
    Dim dataset_name As String
    Dim R As Range
    Dim N As name
    
    Dim optionName As String
    Dim optionValue As String
    
    'On Error Resume Next
    
    dataset_name = WB.name
    dataset_name = Mid(dataset_name, 1, InStr(dataset_name, ".") - 1)
    
    GetWorkbookData = "<dataset>"
    
    AppendLine GetWorkbookData, "<source-info>"
    AppendLine GetWorkbookData, getTextElementString("DatasetName", dataset_name)
    AppendLine GetWorkbookData, getTextElementString("FileName", WB.FullName)
    AppendLine GetWorkbookData, getTextElementString("GeneratedAt", Format(Now, "yyyy-MM-dd hh:mm:ss"))
    AppendLine GetWorkbookData, "</source-info>"
        
    On Error Resume Next
    
    AppendLine GetWorkbookData, "<options>"
    For Each N In WB.Names
        Set R = N.RefersToRange
        If Not (R Is Nothing) Then
            If R.Worksheet.name = "Options" Then
                optionName = N.name
                optionValue = R.value
                If optionValue <> "" Then
                    AppendLine GetWorkbookData, getTextElementString(optionName, optionValue)
                End If
            End If
        End If
    Next N
    AppendLine GetWorkbookData, "</options>"
    
    On Error GoTo 0
    
    For Each S In WB.Worksheets
        If Not (S.name = "Options" Or S.name = "Documentation" Or Mid(S.name, 1, 1) = "_") Then
            AppendLine GetWorkbookData, GetSheetData(S)
        End If
    Next S
    
    AppendLine GetWorkbookData, "</dataset>"
    
End Function

'-------------------------------------------------------------------------------
Function GetSheetData(S As Worksheet) As String
    Dim R As Range
    Dim rowIdx As Integer
    Dim colIdx As Integer
    Dim endOfSheet As Boolean
    Dim blankLinesSkipped As Integer
    Dim Data As String
    
    colIdx = 1
    rowIdx = 1
    endOfSheet = False
    blankLinesSkipped = 0
    
    Data = "<data ref='" + S.name + "'>"
    
    blankLinesSkipped = skipBlankLine(S, colIdx, rowIdx, MAX_BLANK_LINES_BETWEEN_BLOCKS)
    endOfSheet = blankLinesSkipped > MAX_BLANK_LINES_BETWEEN_BLOCKS
    
    While Not endOfSheet
        Set R = S.Cells(rowIdx, colIdx)
        
        If R.value = "Name" Or R.value = "Description" Then
            AppendLine Data, GetDescriptionData(R)
        ElseIf Mid(R.value, 1, 1) <> "_" Then
            AppendLine Data, GetBlockData(R)
        End If

        skipTillBlankLine S, colIdx, rowIdx
        blankLinesSkipped = skipBlankLine(S, colIdx, rowIdx, MAX_BLANK_LINES_BETWEEN_BLOCKS)

        endOfSheet = blankLinesSkipped > MAX_BLANK_LINES_BETWEEN_BLOCKS
    Wend
    
    AppendLine Data, "</data>"
    
    GetSheetData = Data
End Function

'-------------------------------------------------------------------------------
Private Sub skipTillBlankLine(ByRef S As Worksheet, ByRef col As Integer, ByRef row As Integer)
    Dim C As Range
    
    While S.Cells(row, col).value <> ""
        row = row + 1
    Wend
End Sub

'-------------------------------------------------------------------------------
Private Function skipBlankLine(ByRef S As Worksheet, ByRef col As Integer, ByRef row As Integer, maxLines As Integer) As Integer
    Dim skippedLines As Integer
    skippedLines = 0
    
    While S.Cells(row, col).value = "" And skippedLines <= maxLines
        row = row + 1
        skippedLines = skippedLines + 1
    Wend
    
    skipBlankLine = skippedLines
    
End Function

'-------------------------------------------------------------------------------
Private Function GetDescriptionData(Start As Range) As String
    Dim S As Worksheet
    Dim DescriptionID As String
    Dim DescriptionValue As String
    
    DescriptionID = Start.value
    
    Set S = Start.Worksheet
    DescriptionValue = S.Cells(Start.row, Start.Column + 1).value
    
    GetDescriptionData = getTextElementString(DescriptionID, DescriptionValue)
End Function

'-------------------------------------------------------------------------------
Private Function GetBlockData(Start As Range) As String
    Dim S As Worksheet
    Dim Result As String
    Dim BlockID As String
    Dim BlockTypeID As String
    
    Set S = Start.Worksheet
    
    BlockID = Start.value
    BlockTypeID = S.Cells(Start.row, Start.Column + 1).value
    
    Result = "<" + BlockID
    If BlockTypeID <> "" Then
        Append Result, " type='" + BlockTypeID + "'"
    Else
        BlockTypeID = BlockID
    End If
    Append Result, ">"
   
    AppendLine Result, GetColumnDescription(Start)
    AppendLine Result, GetData(Start, BlockTypeID)
    AppendLine Result, "</" + BlockID + ">"

    GetBlockData = Result
End Function

'-------------------------------------------------------------------------------
Private Function GetColumnDescription(Start As Range) As String
    Dim S As Worksheet
    
    Dim col As Integer
    Dim DataTypeRow As Integer
    Dim ColumnNameRow As Integer
    Dim CaptionRow As Integer
    
    Dim columnType As String
    Dim columnName As String
    Dim columnCaption As String
    
    Dim cellStr As String
    
    DataTypeRow = Start.row + 1
    ColumnNameRow = Start.row + 2
    CaptionRow = Start.row + 3
    
    Set S = Start.Worksheet
    
    GetColumnDescription = "<columns>"
    
    col = Start.Column
    
    While S.Cells(DataTypeRow, col).value <> ""
        columnType = S.Cells(DataTypeRow, col).value
        columnName = S.Cells(ColumnNameRow, col).value
        columnCaption = S.Cells(CaptionRow, col).value
        
        If Mid(columnType, 1, 1) <> "_" Then
            cellStr = "<column>" _
                      + getTextElementString("type", columnType) _
                      + getTextElementString("name", columnName) _
                      + getTextElementString("caption", columnCaption) _
                      + "</column>"
        Else
            cellStr = ""
        End If
        AppendLine GetColumnDescription, cellStr
        
        col = col + 1
    Wend
    
    AppendLine GetColumnDescription, "</columns>"
    
End Function


'-------------------------------------------------------------------------------
Private Function GetData(Start As Range, parentID As String) As String
    Dim S As Worksheet
    
    Dim col As Integer
    Dim DataTypeRow As Integer
    Dim ColumnNameRow As Integer
    Dim DataRow As Integer
    
    Dim columnType As String
    Dim name As String
    Dim value As String
    
    Dim cellStr As String
    
    DataTypeRow = Start.row + 1
    ColumnNameRow = Start.row + 2
    DataRow = Start.row + 4
    
    Set S = Start.Worksheet
    
    While S.Cells(DataRow, Start.Column).value <> ""
        AppendLine GetData, "<row id='" + parentID + "_" + Format(DataRow - Start.row - 3) + "'>"
        
        col = Start.Column
        
        While S.Cells(DataTypeRow, col).value <> ""
            columnType = S.Cells(DataTypeRow, col).value
            name = S.Cells(ColumnNameRow, col).value
            value = S.Cells(DataRow, col).value
            
            If Mid(columnType, 1, 1) <> "_" Then
                cellStr = getElementString(columnType, name, value)
            Else
                cellStr = ""
            End If
            AppendLine GetData, cellStr
            
            col = col + 1
        Wend
        
        AppendLine GetData, "</row>"
        
        DataRow = DataRow + 1
    Wend
    
End Function


'-------------------------------------------------------------------------------
Private Function getElementValue(value As String) As String
    Dim badStrings As String
    badstring = "<>?/\&%@"
    
    doEscape = False
    
    For i = 1 To Len(badstring)
        searchChr = Mid(badstring, i, 1)
        If InStr(value, searchChr) > 0 Then
            doEscape = True
            Exit For
        End If
    Next i
    
    getElementValue = value
    
    If doEscape Then
        getElementValue = "<![CDATA[" + getElementValue + "]]>"
    End If
    
End Function

'-------------------------------------------------------------------------------
Private Function getElementString(elementName As String, name As String, value As String) As String
    getElementString = "<" + elementName
    If name <> "" Then
        getElementString = getElementString + " name='" + name + "'"
    End If
    getElementString = getElementString + ">" + getElementValue(value) + "</" + elementName + ">"
End Function

'-------------------------------------------------------------------------------
Private Function getTextElementString(elementName As String, value As String) As String
    getTextElementString = getElementString(elementName, "", value)
End Function

'-------------------------------------------------------------------------------
Sub Append(ByRef Trg As String, ByRef newString As String)
    Trg = Trg + newString
End Sub

'-------------------------------------------------------------------------------
Sub AppendLine(ByRef Trg As String, ByRef newString As String)
    Append Trg, newString
    Append Trg, Chr(10)
End Sub

'-------------------------------------------------------------------------------
Sub AppendLineOffsetTab(ByRef Trg As String, ByRef newString As String)
    Append Trg, Chr(9)
    AppendLine Trg, newString
End Sub

'-------------------------------------------------------------------------------
Sub AppendLineOffset(ByRef Trg As String, ByRef newString As String)
    Append Trg, "  "
    AppendLine Trg, newString
End Sub



