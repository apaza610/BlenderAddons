Sub Main
	Dim oSheet1 As Object
    Dim oSheet2 As Object
    Dim oCell As Object
    Dim oTargetCell As Object
    
    Dim sFilePath As String		' E:\apps\art3d\XNALara XPS\11.8.9\data\SilentHill\Douglas\prueba.csv
    Dim sFileFoldr As String		' E:\apps\art3d\XNALara XPS\11.8.9\data\SilentHill\Douglas
    Dim i as Integer
    
    sFilePath = ThisComponent.URL				' Get the URL of the current document
    sFilePath = ConvertFromURL(sFilePath)		' Convert URL to system path
    'MsgBox "Current file path: " & sFilePath	 
    MsgBox GetPadreFolder(sFilePath)				' Get parent folder	
    
    EnsureSheetExists("Sheet2")
    oSheet1 = ThisComponent.Sheets(0) ' first sheet			' Get the sheets
    oSheet2 = ThisComponent.Sheets.getByName("Sheet2")
    
    Dim colValues As String
    CopyColumnWithSuffix("B", "copia", "_D")
    CopyColumnWithSuffix("H", "extractalfa","_A")
End Sub

Function GetPadreFolder(sFilePath)
	For i = Len(sFilePath) To 1 Step -1
		If Mid(sFilePath, i, 1) = "\" Then
			sFileFoldr = Left(sFilePath, i-1)
			Exit For
		End If
	Next i
	GetPadreFolder = sFileFoldr
End Function

Sub EnsureSheetExists(sheetName As String)
    Dim oSheets As Object
    Dim oSheet As Object
    Dim i As Long
    Dim found As Boolean
    
    oSheets = ThisComponent.Sheets
    found = False
    
    For i = 0 To oSheets.getCount() - 1							' Check if sheet already exists
        If oSheets.getByIndex(i).Name = sheetName Then
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then							' create a new sheet
        oSheets.insertNewByName(sheetName, oSheets.getCount())
    End If
End Sub

' ---------------------------------------------------------------------------------------------------------------
Sub CopyColumnWithSuffix(colLetter As String, actionType As String, suffix As String)
    Dim oDoc As Object
    Dim oSheetSrc As Object, oSheetDest As Object
    Dim colIndex As Long, lastRow As Long
    Dim i As Long, j As Long
    Dim oCellSrc As Object, oCellDest As Object
    Dim sValue As String, newValue As String
    Dim baseName As String, ext As String
    Dim dotPos As Long
    Dim sFileFolder As String
    Dim sourcePath As String, targetPath As String
    Dim processed As Long
    Dim cmd As String, wsh As Object

    oDoc = ThisComponent
    oSheetSrc = oDoc.CurrentController.ActiveSheet

    ' Ensure Sheet2 exists and get reference
    EnsureSheetExists "Sheet2"
    oSheetDest = oDoc.Sheets.getByName("Sheet2")

    ' Determine folder of the current document
    If oDoc.getURL() = "" Then
        MsgBox "Please save the spreadsheet before running this macro."
        Exit Sub
    End If
    sFileFolder = GetPadreFolder(ConvertFromURL(oDoc.getURL()))
    
    colIndex = Asc(UCase(colLetter)) - Asc("A")			' Convert column letter (A → 0, B → 1, etc.)

    lastRow = 1									' Find last used row
    Dim totalRows As Long
    totalRows = oSheetSrc.getRows().getCount() - 1
    For i = 1 To totalRows
        If Trim(oSheetSrc.getCellByPosition(colIndex, i).getString()) <> "" Then lastRow = i
    Next i

    processed = 0

    ' Loop rows starting from B2 (index 1)
    For i = 1 To lastRow
        oCellSrc = oSheetSrc.getCellByPosition(colIndex, i)
        sValue = Trim(oCellSrc.getString())
        If sValue <> "" Then
            ' Find last dot
            dotPos = 0
            For j = Len(sValue) To 1 Step -1
                If Mid(sValue, j, 1) = "." Then
                    dotPos = j
                    Exit For
                End If
            Next j

            If dotPos > 0 Then
                baseName = Left(sValue, dotPos - 1)
                ext = Mid(sValue, dotPos)
            Else
                baseName = sValue
                ext = ""
            End If

            newValue = baseName & suffix & ext
            sourcePath = sFileFolder & "\" & sValue
            targetPath = sFileFolder & "\" & newValue

            Select Case LCase(actionType)
                Case "copia"
                    On Error Resume Next
                    FileCopy sourcePath, targetPath
                    On Error GoTo 0

                Case "extractalfa"
                    Dim alphaName As String, alphaPath As String
                    alphaName = baseName & "_A" & ext
                    alphaPath = sFileFolder & "\" & alphaName
                    cmd = "magick convert " & Chr(34) & sourcePath & Chr(34) & " -alpha extract " & Chr(34) & alphaPath & Chr(34)

                    On Error Resume Next
                    Set wsh = CreateObject("WScript.Shell")
                    If Err.Number = 0 Then
                        wsh.Run Environ("COMSPEC") & " /c " & cmd, 0, True
                    Else
                        Shell(Environ("COMSPEC") & " /c " & cmd, 0)
                    End If
                    On Error GoTo 0

                    newValue = alphaName
            End Select
            
            oCellDest = oSheetDest.getCellByPosition(colIndex, i)	' Write result into Sheet2 same position
            oCellDest.String = newValue
            processed = processed + 1
        End If
    Next i

    MsgBox "Action '" & actionType & "' completed on column " & colLetter & " — " & processed & " entries processed."
End Sub

