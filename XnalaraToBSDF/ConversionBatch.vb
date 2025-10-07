Sub Main
    Dim sFilePath As String
    Dim sFileFoldr As String
    Dim i As Integer
    Dim maxElements As Long
    Dim oSheet1 As Object
    Dim oSheet2 As Object

    ' Get document path
    sFilePath = ThisComponent.URL
    sFilePath = ConvertFromURL(sFilePath)

    If sFilePath = "" Then
        MsgBox "Please save the spreadsheet before running this macro."
        Exit Sub
    End If

    sFileFoldr = GetPadreFolder(sFilePath)

    ' Ask for max elements to process
    maxElements = InputBox("Enter maximum number of elements to process per column:", "Max Elements", "100")
    If maxElements <= 0 Then
        MsgBox "Operation cancelled."
        Exit Sub
    End If

    ' Ensure destination sheet
    EnsureSheetExists("Sheet2")
    oSheet1 = ThisComponent.Sheets(0)
    oSheet2 = ThisComponent.Sheets.getByName("Sheet2")

    ' Process columns
    CopyColumnWithSuffix "A", "nones", "", maxElements
    CopyColumnWithSuffix "B", "copia", "_D", maxElements
    CopyColumnWithSuffix "D", "invertir", "_R", maxElements
    CopyColumnWithSuffix "F", "invertgreen", "_NG", maxElements
    CopyColumnWithSuffix "H", "extractalfa", "_A", maxElements
        
    ExportSheet2AsCSV
End Sub


' --------------------------- Utilities ---------------------------

Function GetPadreFolder(sFilePath As String) As String
    Dim i As Long, sFileFoldr As String
    For i = Len(sFilePath) To 1 Step -1
        If Mid(sFilePath, i, 1) = "\" Then
            sFileFoldr = Left(sFilePath, i - 1)
            Exit For
        End If
    Next i
    GetPadreFolder = sFileFoldr
End Function


Sub EnsureSheetExists(sheetName As String)
    Dim oSheets As Object
    Dim i As Long, found As Boolean
    oSheets = ThisComponent.Sheets
    found = False
    For i = 0 To oSheets.getCount() - 1
        If oSheets.getByIndex(i).Name = sheetName Then
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        oSheets.insertNewByName(sheetName, oSheets.getCount())
    End If
End Sub


Function LastDotPos(s As String) As Long
    Dim p As Long
    For p = Len(s) To 1 Step -1
        If Mid(s, p, 1) = "." Then
            LastDotPos = p
            Exit Function
        End If
    Next p
    LastDotPos = 0
End Function


' --------------------------- Core Function ---------------------------
Sub CopyColumnWithSuffix(colLetter As String, actionType As String, suffix As String, maxElements As Long)
    Dim oDoc As Object, oSheetSrc As Object, oSheetDest As Object
    Dim colIndex As Long, lastRow As Long, i As Long
    Dim data As Variant, resultData() As Variant
    Dim sFileFolder As String, sTextureFolder As String
    Dim sValue As String, baseName As String, ext As String, dotPos As Long
    Dim batchFile As String, batchFileNum As Integer
    Dim processed As Long
    Dim fso As Object, shellObj As Object, cmdLine As String
    Dim outPath As String

    oDoc = ThisComponent
    oSheetSrc = oDoc.CurrentController.ActiveSheet
    EnsureSheetExists "Sheet2"
    oSheetDest = oDoc.Sheets.getByName("Sheet2")

    sFileFolder = GetPadreFolder(ConvertFromURL(oDoc.getURL()))
    sTextureFolder = sFileFolder & "\texturas"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sTextureFolder) Then fso.CreateFolder(sTextureFolder)

    colIndex = Asc(UCase(colLetter)) - Asc("A")
    lastRow = maxElements
    data = oSheetSrc.getCellRangeByPosition(colIndex, 1, colIndex, lastRow).getDataArray()
    ReDim resultData(UBound(data)) As Variant

    batchFile = sFileFolder & "\_bulk_" & actionType & ".cmd"
    batchFileNum = FreeFile
    Open batchFile For Output As #batchFileNum
    Print #batchFileNum, "@echo off"
    Print #batchFileNum, "cd /d """ & sFileFolder & """"

    processed = 0

    For i = LBound(data) To UBound(data)
        sValue = Trim(data(i)(0))
        If sValue <> "" Then
            ' <-- use LastDotPos (your helper above) instead of InStrRev
            dotPos = LastDotPos(sValue)
            If dotPos > 0 Then
                baseName = Left(sValue, dotPos - 1)
                ext = Mid(sValue, dotPos)
            Else
                baseName = sValue
                ext = ""
            End If

            Select Case LCase(actionType)
                Case "nones"
                    ' just copy as-is, no texturas, no .jpg
                    resultData(i) = Array(sValue)

                Case "copia"
                    outPath = "texturas\" & baseName & suffix & ".jpg"
                    Print #batchFileNum, "magick convert """ & sValue & """ """ & outPath & """"
                    resultData(i) = Array(outPath)

                Case "invertir"
                    outPath = "texturas\" & baseName & suffix & ".jpg"
                    Print #batchFileNum, "magick convert """ & sValue & """ -negate """ & outPath & """"
                    resultData(i) = Array(outPath)

                Case "invertgreen"
                    outPath = "texturas\" & baseName & suffix & ".jpg"
                    Print #batchFileNum, "magick convert """ & sValue & """ -channel G -negate """ & outPath & """"
                    resultData(i) = Array(outPath)

                Case "extractalfa"
                    outPath = "texturas\" & baseName & suffix & ".jpg"
                    Print #batchFileNum, "magick convert """ & sValue & """ -alpha extract """ & outPath & """"
                    resultData(i) = Array(outPath)

                Case Else
                    resultData(i) = Array("")
            End Select

            processed = processed + 1
        Else
            resultData(i) = Array("")
        End If
    Next i

    Close #batchFileNum

    oSheetDest.getCellRangeByPosition(colIndex, 1, colIndex, lastRow).setDataArray(resultData)

    Set shellObj = CreateObject("WScript.Shell")
    cmdLine = "cmd.exe /c """ & batchFile & """"
    shellObj.Run cmdLine, 0, True

    MsgBox "Action '" & actionType & "' completed for column " & colLetter & " (" & processed & " processed)."
End Sub


' ------------------------------------------------------------
Sub ExportSheet2AsCSV()
    Dim oDoc As Object
    Dim sFileFolder As String, sFileName As String
    Dim args(1) As New com.sun.star.beans.PropertyValue

    oDoc = ThisComponent

    ' Make Sheet2 active
    oDoc.CurrentController.setActiveSheet(oDoc.Sheets.getByName("Sheet2"))

    ' CSV path
    sFileFolder = GetPadreFolder(ConvertFromURL(oDoc.URL))
    sFileName = sFileFolder & "\listaOUT.csv"

    ' CSV export properties
    args(0).Name = "FilterName"
    args(0).Value = "Text - txt - csv (StarCalc)"
    args(1).Name = "FilterOptions"
    args(1).Value = "44,34,0,1"  ' comma, quote, no page break, export only active sheet

    ' Export
    oDoc.storeToURL(ConvertToURL(sFileName), args)

    MsgBox "Sheet2 saved as " & sFileName
End Sub
