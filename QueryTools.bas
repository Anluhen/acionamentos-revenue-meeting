Attribute VB_Name = "QueryTools"
' ----- Version -----
'        1.5.1
' -------------------

Option Explicit

Sub AtualizarConsultaPCP(Optional ShowOnMacroList = False)
    
    OptimizeCodeExecution True

    ' Refresh the Power Query in "PREVISÃO ESTOQUE"
    On Error Resume Next
    ThisWorkbook.Sheets("Previsão de Estoque").ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
    If Err.Number <> 0 Then
        MsgBox "A consulta da Previsão de Estoque do PCP não pode ser atualizada", vbInformation
    End If
    On Error GoTo 0
    
    OptimizeCodeExecution False
    
End Sub

Sub AtualizarConsulta(Optional ShowOnMacroList = False)
    
    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

    Dim wsConsulta As Worksheet
    Dim wsFaturamento As Worksheet
    Dim wsHistorico As Worksheet
    Dim dataRange As Range
    Dim tbl As ListObject
    Dim targetRange As Range
    Dim rowCount As Long, colCount As Long, targetColCount As Long, targetRowCount As Long
    Dim savedValues As Object
    Dim savedPMNames As Object
    Dim currentDate As Date
    Dim currentWeek As Integer
    Dim i As Long
    Dim ID As Variant
    Dim targetArray As Variant
    Dim manualDataEntryStart As Long
    Dim manualDataEntrySize As Long
    Dim statusColumn As Long
    Dim IDColumn As Long
    Dim PMColumn As Long
    Dim wsAtualizar As Worksheet
    
    OptimizeCodeExecution True
    
Dim temp As Double
Debug.Print "Timers: "
temp = Timer
    
' Maintenance
If False Then
    MsgBox "Em manutenção. Macro desabilitada.", vbInformation
    GoTo CleanExit
End If

    ' Set worksheets
    Set wsConsulta = ThisWorkbook.Sheets("Consulta")
    Set wsFaturamento = ThisWorkbook.Sheets("Faturamento")
    
ErrorSection = "RefreshAnalysis"

    ' Refresh the Power Query in "Consulta"
    On Error Resume Next
    wsConsulta.ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
    If Err.Number <> 0 Then
        If MsgBox("A consulta do Analisys não pode ser atualizada. Deseja prosseguir sem atualizar?", vbYesNo) = vbNo Then
            GoTo CleanExit
        End If
    Else
        
        Dim lastQueryDate As String
        Dim shp As Shape
        Dim txtBox As Shape
        Dim shapeFound As Boolean
    
        ' Get the last modified date from the source file
        lastQueryDate = FileDateTime("\\brjgs100\DFSWEG\GROUPS\BR_SC_JGS_WAU_ADM_CONTRATOS\ACIONAMENTOS\00-EQUIPE DE APOIO\00-BANCO DE DADOS\ANALYSIS_ADCON_WAU.xlsm")
    
        ' Loop through all shapes in the sheet
        shapeFound = False
        For Each shp In wsFaturamento.Shapes
            ' Check if the shape is a text box
            If shp.Type = msoTextBox Then
                shp.TextFrame2.TextRange.Text = "Última atualização do BI: " & vbCrLf & lastQueryDate
                shapeFound = True
                Exit For
            End If
        Next shp
        
    End If
    On Error GoTo ErrorHandler

ErrorSection = "RefreshEstoque"

    ' Refresh the Power Query in "PREVISÃO ESTOQUE"
    On Error Resume Next
    ThisWorkbook.Sheets("Previsão de Estoque").ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
    If Err.Number <> 0 Then
        If MsgBox("A consulta da Previsão de Estoque do PCP não pode ser atualizada. Deseja prosseguir sem atualizar?", vbYesNo) = vbNo Then
            GoTo CleanExit
        End If
    End If
    On Error GoTo ErrorHandler
    
ErrorSection = "GetFaturamentoTable"

    ' Get the existing table in "Faturamento"
    On Error Resume Next
    Set tbl = wsFaturamento.ListObjects("Faturamento")
    If tbl Is Nothing Then
        MsgBox "A tabela não foi encontrada na planilha 'Faturamento'.", vbCritical
        GoTo CleanExit
    End If
    On Error GoTo ErrorHandler

Debug.Print "Time to update query: " & Timer - temp
temp = Timer
   
ErrorSection = "SaveDashboardHistory"

    Call CopyDashboardTableToHistorico

Debug.Print "Time to save dashboard: " & Timer - temp
temp = Timer

ErrorSection = "SaveHistoricalData"
    
    Dim rngSource As Range
    Dim newRowCount As Long
    Dim destRange As Range
    Dim lastRow As Long, lastCol As Long
    Dim dupRange As Range
    Dim colIndices() As Variant
    
    Set wsHistorico = ThisWorkbook.Worksheets("Histórico Faturamento")
    
    ' Set the source range as the table’s data rows
    Set rngSource = tbl.DataBodyRange
    newRowCount = rngSource.Rows.Count
    
    ' Insert new rows at the top of "Histórico Faturamento" to add the new data
    wsHistorico.Rows("2:" & newRowCount).Insert Shift:=xlUp

    ' Ajust formating of new rows
    wsHistorico.Rows(newRowCount + 1).Copy
    wsHistorico.Rows("2:" & newRowCount).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ' Copy data from the table to the destination worksheet.
    ' The table data is pasted starting at cell B1
    Set destRange = wsHistorico.Range("B2").Resize(newRowCount, rngSource.Columns.Count)
    destRange.Value = rngSource.Value
    
    ' Add the current date and time (timestamp) to column A for the new rows
    wsHistorico.Range("A2").Resize(newRowCount, 1).Value = Now
    
    ' Find the overall area that now holds data on "Histórico Faturamento"
    lastRow = wsHistorico.Cells(wsHistorico.Rows.Count, "A").End(xlUp).Row
    lastCol = wsHistorico.Cells(1, wsHistorico.Columns.Count).End(xlToLeft).Column
    Set dupRange = wsHistorico.Range(wsHistorico.Cells(1, 1), wsHistorico.Cells(lastRow, lastCol))
    
    ' Build an array with columns (relative to dupRange) from column B onward.
    ' If the table data has N columns, then the duplicate-check array contains 2,3,...,N+1.
    ReDim colIndices(1 To destRange.Columns.Count)
    For i = 1 To destRange.Columns.Count
        colIndices(i) = i + 1
    Next i

    ' Remove duplicates based only on the table data columns (ignoring column A)
    dupRange.RemoveDuplicates Columns:=Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28), Header:=xlYes

Debug.Print "Time to save past data: " & Timer - temp
temp = Timer

ErrorSection = "SaveData"

    manualDataEntryStart = 23
    manualDataEntrySize = 5
    statusColumn = 5 ' As referenced by the table
    IDColumn = 1 ' As referenced by the table
    PMColumn = 16 ' As referenced by the table
    
    ' --- Columns must be ajusted in case of layout changes ---
    ' Cut manully filled columns J to N to the end of the table before pasting
    Columns("Q:V").Cut
    ' Paste it after the table (+1 column to compensate the shift of the table position
    Columns(tbl.DataBodyRange.Columns.Count + 2).Insert Shift:=xlToRight
    
        
    ' Initialize dictionary to save values
    Set savedValues = CreateObject("Scripting.Dictionary")
    
    ' Initialize dictionary to save values
    Set savedPMNames = CreateObject("Scripting.Dictionary")
    
    ' Save the current values of ID and columns Y to AE
    With tbl.DataBodyRange
        For i = 1 To .Rows.Count
            ID = .Cells(i, IDColumn).Value
            
            If Not savedValues.Exists(ID) And Join(Application.Transpose(Application.Transpose(.Cells(i, manualDataEntryStart).Resize(1, manualDataEntrySize).Value)), "") <> "" _
                And Join(Application.Transpose(Application.Transpose(.Cells(i, manualDataEntryStart).Resize(1, manualDataEntrySize).Value)), "") <> "RECONHECIDO" Then
                savedValues(ID) = .Cells(i, manualDataEntryStart).Resize(1, manualDataEntrySize).Value  ' Save columns Y to AE
            End If
            
            If UCase(.Cells(i, PMColumn).Value) <> "NÃO ATRIBUÍDO" Then
                savedPMNames(ID) = .Cells(i, PMColumn).Value
            End If
            
        Next i
    End With

Debug.Print "Time to save IDs: " & Timer - temp
temp = Timer
    
' ErrorSection = "BackUpData"
'
'    ' Look for old IDs that are not kept and save them to a wokrsheet
'    Dim oldIDs As Object, newIDs As Object
'    Dim wsBackup As Worksheet
'    Dim r As Long
'    Dim idKey As Variant
'    Dim headerCount As Long
'    Dim currentRowRange As Range
'
'    ' Create dictionaries for old and new IDs
'    Set oldIDs = CreateObject("Scripting.Dictionary")
'    Set newIDs = CreateObject("Scripting.Dictionary")
'
'    ' Capture the old IDs from the existing table in "Faturamento"
'    Dim tblRow As Range
'    For Each tblRow In tbl.DataBodyRange.Rows
'        idKey = tblRow.Cells(1, 1).Value
'        If Not oldIDs.Exists(idKey) Then
'            oldIDs.Add idKey, True
'        End If
'    Next tblRow
'
'    ' Determine the range of data in the table on "Consulta"
'    With wsConsulta.ListObjects(1).DataBodyRange
'        rowCount = .Rows.Count
'        colCount = .Columns.Count
'        Set dataRange = .Resize(rowCount, colCount)
'    End With
'
'    ' Capture the new IDs from the "Consulta" data range
'    For r = 1 To dataRange.Rows.Count
'        idKey = dataRange.Cells(r, 1).Value
'        If Not newIDs.Exists(idKey) Then
'            newIDs.Add idKey, True
'        End If
'    Next r
'
'    ' Create or clear the backup worksheet "BackupIDs"
'    On Error Resume Next
'    Set wsBackup = ThisWorkbook.Sheets("BackupIDs")
'    On Error GoTo ErrorHandler
'    If wsBackup Is Nothing Then
'        Set wsBackup = ThisWorkbook.Sheets.Add(After:=wsFaturamento)
'        wsBackup.Name = "BackupIDs"
'    Else
'        wsBackup.Cells.Clear
'    End If
'
'    ' Copy the headers from tbl to wsBackup
'    headerCount = tbl.HeaderRowRange.Columns.Count
'    wsBackup.Range(wsBackup.Cells(1, 1), wsBackup.Cells(1, headerCount)).Value = tbl.HeaderRowRange.Value
'
'    ' Loop through each row in the target table
'    Dim backupRow As Long: backupRow = 2
'    For Each currentRowRange In tbl.DataBodyRange.Rows
'        idKey = currentRowRange.Cells(1, 1).Value ' Assuming the ID is in the first column
'        If Not newIDs.Exists(idKey) Then
'            ' Copy the entire row (all columns) to the backup worksheet
'            wsBackup.Range(wsBackup.Cells(backupRow, 1), wsBackup.Cells(backupRow, headerCount)).Value = currentRowRange.Value
'            backupRow = backupRow + 1
'        End If
'    Next currentRowRange
'
'Debug.Print "Time to save backupIDs: " & Timer - temp
'temp = Timer

ErrorSection = "ResizeClearedTable"

    ' Determine the range of data in the table on "Consulta"
    With wsConsulta.ListObjects(1).DataBodyRange
        rowCount = .Rows.Count
        colCount = .Columns.Count
        Set dataRange = .Resize(rowCount, colCount)
    End With
    
    ' Clear and ensure the table range matches the new data size
    With tbl
    
        ' Clear filters
        If .AutoFilter.FilterMode Then .AutoFilter.ShowAllData
        
        .DataBodyRange.ClearContents
        
        targetColCount = .ListColumns.Count ' Get the total number of columns in the target table
        
        ' Resize the table to match the data size
        .Resize .HeaderRowRange.Resize(rowCount + 1, targetColCount)
        Set targetRange = .DataBodyRange.Resize(rowCount, colCount)
    End With
    
ErrorSection = "FillInData"
    
    ' Copy data into the existing table
    targetRange.Value = dataRange.Value

Debug.Print "Time to update table: " & Timer - temp
temp = Timer

    ' Load the DataBodyRange into an array
    With tbl.DataBodyRange
        targetArray = .Value
        targetRowCount = UBound(targetArray, 1)
    End With

    ' --- Columns must be ajusted in case of layout changes ---
    ' Restore the saved values for columns Y to AE
    For i = 1 To targetRowCount
    
ErrorSection = "RestoreSavedData-" & i

        ID = targetArray(i, IDColumn)
        If savedValues.Exists(ID) Then
            ' Assign saved values for the row
            Dim savedArr() As Variant
            savedArr = savedValues(ID)
            Dim j As Long
            For j = 0 To manualDataEntrySize - 1
                targetArray(i, manualDataEntryStart + j) = savedArr(1, j + 1)
            Next j
        ElseIf targetArray(i, statusColumn) = "RECONHECIDO" Then
            Dim k As Long
            For k = 0 To manualDataEntrySize - 3
                targetArray(i, manualDataEntryStart + k) = ""
            Next k
            targetArray(i, manualDataEntryStart + manualDataEntrySize - 2) = "RECONHECIDO"
        Else
            Dim l As Long
            For l = 0 To manualDataEntrySize - 1
                targetArray(i, manualDataEntryStart + l) = ""
            Next l
        End If
        
        If savedPMNames.Exists(ID) And UCase(targetArray(i, PMColumn)) = "NÃO ATRIBUÍDO" And UCase(savedPMNames(ID)) <> "" Then
            targetArray(i, PMColumn) = savedPMNames(ID)
        End If
    Next i

ErrorSection = "RestoreSavedData"
    
    ' Write the updated array back to the worksheet
    With tbl.DataBodyRange
        .Value = targetArray
    End With
    
Debug.Print "Time to restore values: " & Timer - temp
temp = Timer

ErrorSection = "FormatSheet"

    ' --- Columns must be ajusted in case of layout changes ---
    ' Restore columns position
    Columns("W:AB").Cut
    Columns("Q").Insert Shift:=xlToRight
    
Debug.Print "Time to add new week: " & Timer - temp
temp = Timer
    
    ' --- Columns must be ajusted in case of layout changes ---
    ' Ajust column widths
    Columns("C:AB").AutoFit
    Columns("I").ColumnWidth = 5 ' Item Doc. Vendas
    Columns("L").ColumnWidth = 5 ' Incoterms
    Columns("M:R").ColumnWidth = 12 ' Datas
    Columns("U").ColumnWidth = 35 ' Observação
    ' Columns("V").ColumnWidth = 10 ' Situação
    Columns("W").ColumnWidth = 20 ' PM

ErrorSection = "TreatData"

    ' Treate data
     With tbl.DataBodyRange
        ' Load the data into an array for faster processing
        targetArray = .Value
        targetRowCount = .Rows.Count
        ReDim FontColors(1 To targetRowCount)
        ReDim FillColors(1 To targetRowCount)
        
        ' Loop through each row to update column 16 based on column 13
        Dim dt As Date
        For i = 1 To targetRowCount
        
ErrorSection = "TreatData-" & i
    
            ' Fill in the date for the BI Dashboard
            
            ' --- Columns must be ajusted in case of layout changes ---
            ' If Status = RECONHECIDO, then use the BI Month and Year
            If targetArray(i, statusColumn) = "RECONHECIDO" Then
                ' Data de Rec. Receita column
                targetArray(i, 17) = DateSerial(targetArray(i, 3), Month(DateValue("1 " & targetArray(i, 4) & " 2000")) + 1, 0)
            ' If date in Dados Adicionais B is greater than the writen date, update the date
            ElseIf targetArray(i, 14) > targetArray(i, 17) Then
                ' Data de Rec. Receita column = Data Dados Ad. B column
                targetArray(i, 17) = targetArray(i, 14)
            End If
            ' Else (if date Dados Adivionais B is older) leave the manual filled date as it is
        Next i
        
        ' Write the updated data back to the worksheet (updating column 17 if needed)
        .Value = targetArray
        
        ' Apply formula to column 17 of the table
        tbl.ListColumns(16).DataBodyRange.Formula2 = "=IF(IFERROR(INDEX(PREVISÃO_ESTOQUE[Previsão de Estoque],MATCH(LEFT([@PEP],LEN([@PEP])-1),LEFT(PREVISÃO_ESTOQUE[[#Data],[PEP]],LEN(PREVISÃO_ESTOQUE[[#Data],[PEP]])-1),0)),"""")=0,"""",IFERROR(INDEX(PREVISÃO_ESTOQUE[Previsão de Estoque],MATCH(LEFT([@PEP],LEN([@PEP])-1),LEFT(PREVISÃO_ESTOQUE[[#Data],[PEP]],LEN(PREVISÃO_ESTOQUE[[#Data],[PEP]])-1),0)),""""))"
        
        ' Apply formula to column 27 and 28 of the table
        tbl.ListColumns(26).DataBodyRange.Formula = "=IF([@[Data de Rec.Receita ]]="""",[@[Mês BI]],UPPER(TEXT([@[Data de Rec.Receita ]],""mmmm"")))"
        tbl.ListColumns(27).DataBodyRange.Formula = "=IF([@[Data de Rec.Receita ]]="""",[@[Ano BI]],UPPER(TEXT([@[Data de Rec.Receita ]],""aaaa"")))"

        ' Highlight
        'For i = 1 To targetRowCount
        '    If targetArray(i, 5) <> "RECONHECIDO" And (targetArray(i, 14) = "RISCO" Or targetArray(i, 17) <= Now Or targetArray(i, 14) <> targetArray(i, 17)) Then
        '        FontColors(i) = RGB(255, 0, 0) ' Red font
        '        FillColors(i) = RGB(255, 200, 200) ' Light red fill
        '    Else
        '        FontColors(i) = RGB(0, 0, 0) ' Default font color (black)
        '        FillColors(i) = xlNone ' Clear fill color
        '    End If
        'Next i
    
        ' Write the results back to the worksheet
        'For i = 1 To targetRowCount
        '    .Rows(i).Font.Color = FontColors(i)
        '    .Rows(i).Interior.Color = FillColors(i)
        'Next i
    End With
    
Debug.Print "Time to highlight rows: " & Timer - temp
temp = Timer

    ' Set wsAtualizar = ThisWorkbook.Sheets("Atualizar Datas")
    
    ' Place the formulas in the specified cells (using double quotes for embedded quotes)
    'wsAtualizar.Range("A2").Formula = "=FILTER(Faturamento!$G:$G,Faturamento!$V:$V=""NÃO"")"
    'wsAtualizar.Range("B2").Formula = "=FILTER(Faturamento!$H:$H,Faturamento!$V:$V=""NÃO"")"
    'wsAtualizar.Range("C2").Formula = "=LET(x, FILTRO(Faturamento!$Q:$Q,Faturamento!$V:$V=""NÃO"", """"), SE((x="""")+(x=0), """", x))"
    'wsAtualizar.Range("E2").Formula = "=LET(x, FILTRO(Faturamento!$R:$R,Faturamento!$V:$V=""NÃO"", """"), SE((x="""")+(x=0), """", x))"
    'wsAtualizar.Range("F2").Formula = "=LET(x, FILTRO(Faturamento!$S:$S,Faturamento!$V:$V=""NÃO"", """"), SE((x="""")+(x=0), """", x))"
    
    'Application.Calculate
    
    ' Use Autofill to fill column D with "IGNORAR" from D2 down to the last data row
    'If wsAtualizar.Cells(wsAtualizar.Rows.Count, "A").End(xlUp).Row >= 2 Then
    '    wsAtualizar.Range("D2").Value = "IGNORAR"
    '    wsAtualizar.Range("D2").AutoFill Destination:=wsAtualizar.Range("D2:D" & wsAtualizar.Cells(wsAtualizar.Rows.Count, "A").End(xlUp).Row), Type:=xlFillDefault
    'Else
    '    wsAtualizar.Range("D2").Value = ""
    '    wsAtualizar.Range("D2").AutoFill Destination:=wsAtualizar.Range("D2:D" & wsAtualizar.Cells(wsAtualizar.Rows.Count, "A").End(xlUp).Row), Type:=xlFillDefault
    'End If
    
    ' Clear values in columns G, H, and I (from row 2 downward) while keeping row 1 intact
    'wsAtualizar.Range("A2:I" & wsAtualizar.Cells(wsAtualizar.Rows.Count, "G").End(xlUp).Row).ClearContents
    
CleanExit:
    ' Ensure that all optimizations are turned back on
    OptimizeCodeExecution False
    
    MsgBox "Planilha atualizada!", vbInformation
    
    Exit Sub

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in ExampleSub"
    
    ' Optionally, you can log the error details to a file or a logging system here
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
    
End Sub

Sub CopyDashboardTableToHistorico(Optional ShowOnMacroList = False)
    
    Dim srcWS As Worksheet, destWS As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    Dim numRows As Long
    Dim insertRows As Long

    ' Set the source and destination worksheets
    Set srcWS = ThisWorkbook.Worksheets("DASHBOARD")
    Set destWS = ThisWorkbook.Worksheets("Histórico Dashboard")

    ' Find the last used row in DASHBOARD
    If Application.WorksheetFunction.CountA(srcWS.Cells) = 0 Then
        MsgBox "The DASHBOARD sheet is empty."
        Exit Sub
    End If
    lastRow = srcWS.Cells.Find(What:="*", LookIn:=xlFormulas, _
                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    ' Find the last used column in DASHBOARD
    lastCol = srcWS.Cells.Find(What:="*", LookIn:=xlFormulas, _
                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' Define the data range to be copied
    Set dataRange = srcWS.Range(srcWS.Cells(1, 1), srcWS.Cells(lastRow, lastCol))
    numRows = dataRange.Rows.Count

    ' Calculate the total number of rows to insert (1 for timestamp + data rows)
    insertRows = numRows + 2

    ' Insert rows at the top of the destination sheet so existing rows move down
    destWS.Rows("1:" & insertRows).Insert Shift:=xlDown

    ' Add the timestamp in the new top row (cell A1)
    destWS.Range("A1").Value = "Dados de: " & Format(Now, "dd/mm/yyyy hh:mm:ss")

    ' Copy the data from the source and paste starting at row 2 of the destination
    dataRange.Copy Destination:=destWS.Range("A2")

    ' Optional: Clear the clipboard to remove the copy selection
    Application.CutCopyMode = False
    
    ' Ajust column widths
    destWS.Columns("A:AAA").AutoFit
    
End Sub

Function OptimizeCodeExecution(enable As Boolean)
    With Application
        If enable Then
            ' Disable settings for optimization
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
        Else
            ' Re-enable settings after optimization
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
        End If
    End With
End Function

