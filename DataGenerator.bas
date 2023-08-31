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

' Converts XML to UTF-8 without BOM / otherwise Windows-1251
Const C_ENCODING_WIN = "WIN"
Const C_ENCODING_UTF8 = "UTF8"
Const DEFAULT_ENCODING = C_ENCODING_UTF8

' Pretty print output XMLs (1/0)
Const C_FORMAT_PRETTYPRINT = "PP"
Const C_FORMAT_PLAIN = "PLAIN"
Const DEFAULT_FORMAT = C_FORMAT_PRETTYPRINT


Const IS_EXPORT_DATA_EXTRACT = "1" ' Flag to export intermediate data.xml files (1/0)

' Режим отладки - показывает больше данных при завершени операции
Const IS_DEBUG_MODE = "0" ' Debug mode - more info (1/0)

Const MAX_BLANK_LINES_BETWEEN_BLOCKS = 2


Dim FilesList As String

'-------------------------------------------------------------------------------
Sub MakeDataExtract()
    Dim WB As Workbook
    Dim Data As String
    
    'MsgBox "MakeDataExtract" '----------
    
    Set WB = Application.ActiveWorkbook
    
    Data = "<?xml version='1.0' encoding='windows-1251'?>"
    
    AppendLine Data, ""
    AppendLine Data, GetWorkbookData(WB)
    
    ExportFileName = WB.FullName
    
    ExportFileName = Mid(ExportFileName, 1, InStr(ExportFileName, ".") - 1)
    
    ExportFileName = ExportFileName + DEFAULT_DATA_EXTENSION    '.TestData.xml
    
    SaveToXML ExportFileName, Data, C_ENCODING_WIN 'Save in Win-1251
    
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
    
    
    'Data = "<?xml version='1.0' encoding='UTF-8'?>"
    Data = "<?xml version='1.0' encoding='windows-1251'?>"
    
    AppendLine Data, ""
    AppendLine Data, GetWorkbookData(WB)
    
    '---------------------------------------------------------------------------
    If IS_EXPORT_DATA_EXTRACT = "1" Then
    
        ExportFileName = WB.FullName
        
        ExportFileName = Mid(ExportFileName, 1, InStr(ExportFileName, ".") - 1)
        
        ExportFileName = ExportFileName + DEFAULT_DATA_EXTENSION    '.TestData.xml
        
        SaveToXML ExportFileName, Data, C_ENCODING_WIN 'Save in Win-1251
        
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
            
    
            If DEFAULT_FORMAT = C_FORMAT_PRETTYPRINT Then ' Pretty print output XMLs
                Data = PrettyPrintXml(Data)
            End If
            
            'Save in DEF format
            SaveToXML ExportFileName, Data, DEFAULT_ENCODING
            
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
    Dim sEncoding As String
    Dim sFormat As String
    Dim N As name
    Dim R As Range
    Dim LO As ListObject
    
    
    FilesList = ""
            
    'MsgBox "MakeDataExtractAndTransform" '----------
    
    Set WB = Application.ActiveWorkbook
    
    Set Opts = WB.Sheets("Options")
    
    On Error Resume Next
    
    'Find in excel NamedTable - TransformationOptions
    Set LO = Opts.ListObjects("TransformationOptions")
    
    On Error GoTo 0
   
    'If NOT FOUND
    If (LO Is Nothing) Then
        'Convert One File
        MakeDataExtractAndTransformOne
    Else
        'Use Table TransformationOptions
        Data = "<?xml version='1.0' encoding='windows-1251'?>"
        
        AppendLine Data, ""
        AppendLine Data, GetWorkbookData(WB)
            
            
        '---------------------------------------------------------------------------
        If IS_EXPORT_DATA_EXTRACT = "1" Then
        
            ExportFileName = WB.FullName
            
            ExportFileName = Mid(ExportFileName, 1, InStr(ExportFileName, ".") - 1)
            
            ExportFileName = ExportFileName + DEFAULT_DATA_EXTENSION    '.data.xml
            
            SaveToXML ExportFileName, Data, C_ENCODING_WIN 'Save in Win-1251
            
            'MsgBox "done save - " + ExportFileName '----------
            
        End If
        '---------------------------------------------------------------------------
        
        Set xml = New DOMDocument60
        
        If xml.LoadXML(Data) Then
            Set xslt = New DOMDocument60
            
            xsltPath = Application.ThisWorkbook.Path
            
            For idx = 1 To LO.ListRows.Count
                
                'MsgBox "idx - " + Str(idx) '----------
                
                'Reading table - TransformationOptions
                StylesheetName = LO.ListRows(idx).Range(1) 'Name of XSLT Template
                FileExtension = LO.ListRows(idx).Range(2) 'Output file extention
                
                '(Optional) Output Export Path and FileName
                FileName = ""
                If Not (LO.ListRows(idx).Range(3) Is Nothing) Then
                   ssName = LO.ListRows(idx).Range(3).value
                   If ssName <> "" Then
                       FileName = ssName
                   End If
                End If
                
                '(Optional) Pretty print Export
                sFormat = DEFAULT_FORMAT
                If Not (LO.ListRows(idx).Range(4) Is Nothing) Then
                   ssName = LO.ListRows(idx).Range(4).value
                   If (ssName = C_FORMAT_PRETTYPRINT) Then 'PP
                       sFormat = C_FORMAT_PRETTYPRINT
                   Else
                       sFormat = C_FORMAT_PLAIN
                   End If
                End If
                
                '(Optional) Convert to UTF-8
                sEncoding = DEFAULT_ENCODING
                If Not (LO.ListRows(idx).Range(5) Is Nothing) Then
                   ssName = LO.ListRows(idx).Range(5).value
                   If (ssName = C_ENCODING_WIN) Then 'WIN
                       sEncoding = C_ENCODING_WIN
                   Else
                       sEncoding = C_ENCODING_UTF8
                   End If
                End If
                
                If xslt.Load(xsltPath + "\" + StylesheetName) Then
                
                    
                    'MsgBox "do transformNode - " + xsltPath + "\" + StylesheetName '----------
                    
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
                    
                    
                    Data = xml.transformNode(xslt)
                    
                    'MsgBox "ExportFileName - " + ExportFileName '----------
                    
                    If sFormat = C_FORMAT_PRETTYPRINT Then ' Pretty print output XMLs
                        Data = PrettyPrintXml(Data)
                    End If
                                
                    SaveToXML ExportFileName, Data, sEncoding
                    
                    AppendLine FilesList, " - " + ExportFileName
                    
                    'MsgBox "done save - " + ExportFileName '----------
                
                End If
            Next idx
        End If
    End If

End Sub


'-------------------------------------------------------------------------------
Private Sub SaveToXML(inExportFileName As Variant, inData As Variant, inIS_CONVERT_XML_TO_UTF8 As Variant)

    If inIS_CONVERT_XML_TO_UTF8 = C_ENCODING_UTF8 Then
        'Converts XML to UTF-8 without BOM
        'sEncoding Pretty print output XMLs (1/0)
            
        '-----------------------------------
        'Способ сохранения в кодировке UTF-8 Without!!! BOM
        'Option Explicit

        Const adSaveCreateNotExist = 1
        Const adSaveCreateOverWrite = 2
        Const adTypeBinary = 1
        Const adTypeText = 2
        
        Dim objStreamUTF8: Set objStreamUTF8 = CreateObject("ADODB.Stream")
        Dim objStreamUTF8NoBOM: Set objStreamUTF8NoBOM = CreateObject("ADODB.Stream")
        
        With objStreamUTF8
          .Charset = "UTF-8"
          .Open
          .WriteText inData
          .Position = 0
          .SaveToFile inExportFileName, adSaveCreateOverWrite
          .Type = adTypeText
          .Position = 3
        End With
        
        With objStreamUTF8NoBOM
          .Type = adTypeBinary
          .Open
          objStreamUTF8.CopyTo objStreamUTF8NoBOM
          .SaveToFile inExportFileName, adSaveCreateOverWrite
        End With
                    
        objStreamUTF8.Close
        objStreamUTF8NoBOM.Close
    Else
            
        'Старый способ сохранения (Выпускал кодировку Windows-1251
        Open inExportFileName For Output As #1
        Print #1, inData
        Close #1
        
    End If
End Sub

'-------------------------------------------------------------------------------
Private Function PrettyPrintXml(dom As Variant) As String 'из xml-строки делает форматированный xml

   Dim writer As New MXXMLWriter60

   With writer
       .omitXMLDeclaration = False
       .indent = True
       .byteOrderMark = False
       .standalone = False
   End With

   Dim reader As New SAXXMLReader60

   Set reader.contentHandler = writer
   reader.Parse dom

   PrettyPrintXml = writer.output
   'PrettyPrintXml = Replace(PrettyPrintXml, "encoding=""UTF-16""", "encoding=""UTF-8""") ' windows-1251
   PrettyPrintXml = Replace(PrettyPrintXml, "encoding=""UTF-16""", "") ' windows-1251
   PrettyPrintXml = Replace(PrettyPrintXml, " standalone=""no""", "")

End Function

'-------------------------------------------------------------------------------
Sub FileList_MakeDataExtractAndTransform()
    Dim WB As Workbook
    Dim SList As Worksheet
    Dim aName As String
    Dim MsgStr As String
    
    FilesList = ""
    
    
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
    
    MsgStr = Str(cnt) + " out of " + Str(cntTotal) + " tests has been processed"
    
    If IS_DEBUG_MODE = "1" Then
        AppendLine MsgStr, FilesList
    End If
    
    MsgBox MsgStr
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

