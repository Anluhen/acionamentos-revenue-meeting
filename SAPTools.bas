Attribute VB_Name = "SAPTools"
Option Explicit

Public SapGuiAuto As Object
Public SAPApplication As Object
Public Connection As Object
Public Session As Object

'----------------------------------------------
' Main Routine: Atualizar
'----------------------------------------------
Sub AtualizarDataOV(Optional ShowOnMacroList = False)

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler
    
    ' Turn off optimizations for performance
    OptimizeCodeExecution True
    
ErrorSection = "Initialization"
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim orderNumber As String
    Dim itemCode As Variant
    Dim dataSeparacao As Date, dataRemessa As Date, dataSM As Date
    Dim msg As String
    Dim loopCount As Long
    Const MAX_LOOP As Long = 25 ' Maximum iterations for loops to avoid infinite loops
    Dim currentItem As String, NextItem As String
    Dim j As Long
    Dim strData As String, dia As String, mes As String, ano As String
    Dim response As VbMsgBoxResult
    Dim item1 As String, item2 As String
    Dim ColumnWarningsStart As Long
    Dim popupText As String

ErrorSection = "SAPSetup"

    '--- Setup SAP and check if it is running ---
    Do While Not SetupSAPScripting
        ' Ask the user to initiate SAP or cancel
        response = MsgBox("SAP não está acessível. Inicie o SAP e clique em OK para tentar novamente, ou Cancelar para sair.", vbOKCancel + vbExclamation, "Aguardando SAP")
    
        If response = vbCancel Then
            MsgBox "Execução terminada pelo usuário.", vbInformation
            GoTo CleanExit  ' Exit the function or sub
        End If
    Loop
    
ErrorSection = "WorksheetSetup"
    
    '--- Setup worksheet ---
    Set ws = ThisWorkbook.Worksheets("Atualizar Datas")
    lastRow = ws.Range("A1").End(xlDown).Row
    
    ColumnWarningsStart = 7
    
    '--- Process each sales order ---
    For i = 2 To lastRow
ErrorSection = "OVProcessing-" & i
    
        If InStr(1, ws.Range("D" & i).Value, "DATA OV") > 0 Then
ErrorSection = "DataOVProcessing-" & i
            orderNumber = ws.Range("A" & i).Value
            itemCode = ws.Range("B" & i).Value
            currentItem = ""
            
            '--- Clear workbook values ---
            ws.Cells(i, ColumnWarningsStart).Value = ""
            ws.Cells(i, ColumnWarningsStart + 1).Value = ""
            ws.Cells(i, ColumnWarningsStart + 2).Value = ""
            
            ' Update status bar with progress info
            Application.StatusBar = "Executando... Ordem de Venda " & orderNumber & " (" & i - 1 & " de " & lastRow - 1 & ")"
            
            '--- Open transaction VA02 ---
            If Not SetSAPText(Session, "wnd[0]/tbar[0]/okcd", "/nVA02") Then GoTo NextIteration
            SAPSendVKey Session, 0
            
            '--- Enter the Sales Order number ---
            If Not SetSAPText(Session, "wnd[0]/usr/ctxtVBAK-VBELN", orderNumber) Then GoTo NextIteration
            SAPSendVKey Session, 0
            
            '--- Select first item and open it for editing ---
            On Error GoTo IterationError
            Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                "subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").SetFocus
            SAPPress Session, "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                "subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_PEIN"
            
ajusteDataOV:
        
            currentItem = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/txtVBAP-POSNR").Text
        
            Do While itemCode <> currentItem
                SAPPress Session, "wnd[0]/tbar[1]/btn[19]"
                
                
                currentItem = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/txtVBAP-POSNR").Text
            Loop
            
        
            dataSM = 0
            
            '--- Open the date adjustment screen ---
            SAPPress Session, "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4500/btnP0_EID2"
            
            '--- Check if the "Data Remessa" field is changeable ---
            loopCount = 0
            Do While Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_SHEDLINE/tabpT\02/" & _
                    "ssubSUBSCREEN_BODY:SAPMV45A:4552/ctxtRV45A-ETDAT").changeable = False
                ws.Cells(i, ColumnWarningsStart).Value = "ERRO"
                ws.Cells(i, ColumnWarningsStart + 1).Value = "Remessa Criada ou Item Recusado"
                SAPPress Session, "wnd[0]/tbar[0]/btn[3]"  ' Go back
                item1 = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/txtVBAP-POSNR").Text
                SAPPress Session, "wnd[0]/tbar[1]/btn[19]"  ' Move forward
                item2 = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/txtVBAP-POSNR").Text
                If item1 = item2 Then
                    GoTo NextIteration
                Else
                    GoTo ajusteDataOV
                End If
            Loop
            
            '--- Set separation and remittance dates based on the workbook ---
            dataSeparacao = ws.Range("C" & i).Value
            dataRemessa = dataSeparacao
            If Not SetSAPText(Session, "wnd[0]/usr/tabsTAXI_TABSTRIP_SHEDLINE/tabpT\02/" & _
                    "ssubSUBSCREEN_BODY:SAPMV45A:4552/ctxtRV45A-ETDAT", Format(dataRemessa, "dd.mm.yyyy")) Then GoTo NextIteration
            SAPSendVKey Session, 0
            
            If Session.findbyid("wnd[0]/sbar").Text <> "" And InStr(1, Session.findbyid("wnd[0]/sbar").Text, "Não existem part.vencida") = 0 Then
                GoTo IterationError
            End If
            
            
            '--- Adjust the remittance date until it matches the desired separation date ---
            loopCount = 0
            Do Until dataSM = dataSeparacao
                ' Check for any popup messages
                popupText = ""
                On Error Resume Next
                popupText = Session.ActiveWindow.popupDialogText
                On Error GoTo 0
                
                If popupText <> "" Then
                    ' Handle production order popup
                    If InStr(1, popupText, "já existe uma ordem de produção", vbTextCompare) > 0 Then
                        ws.Cells(i, ColumnWarningsStart + 1).Value = popupText
                        SAPSendVKey Session, 0
                    End If
                    ' Handle credit block message (cancel operation)
                    If popupText = "Não é permitido aceitar ordem/fornec.: crédito cliente bloqueado" Then
                        ws.Cells(i, ColumnWarningsStart).Value = "ERRO"
                        ws.Cells(i, ColumnWarningsStart + 1).Value = popupText
                        SAPSendVKey Session, 0
                        If InStr(1, popupText, "Verificação dinâmica de crédito: limite de crédito foi excedido", vbTextCompare) > 0 Then
                            SAPSendVKey Session, 0
                        End If
                        SAPPress Session, "wnd[0]/tbar[0]/btn[12]"
                        SAPPress Session, "wnd[1]/usr/btnSPOP-OPTION1"
                        GoTo NextIteration
                    End If
                    ' Additional credit-related messages
                    If InStr(1, popupText, "Verificação dinâmica de crédito: limite de crédito foi excedido", vbTextCompare) > 0 Then
                        SAPSendVKey Session, 0
                    End If
                    If InStr(1, popupText, "Verificação 1 do usuário para crédito sem êxito", vbTextCompare) > 0 Then
                        SAPSendVKey Session, 0
                    End If
                    If InStr(1, popupText, "Verificação crédito: partida pendente mais antiga está em atraso", vbTextCompare) > 0 Then
                        SAPSendVKey Session, 0
                    End If
                End If
                
                On Error GoTo IterationError
                ' If first time, read the current SAP material preparation date (Data SM)
                strData = Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_SHEDLINE/tabpT\02/" & _
                    "ssubSUBSCREEN_BODY:SAPMV45A:4552/ctxtVBEP-MBDAT").Text
                dia = Format(Val(Mid(strData, 1, 2)), "00")
                mes = Format(Val(Mid(strData, 4, 2)), "00")
                ano = Format(Val(Mid(strData, 7, 4)), "0000")
                strData = dia & "/" & mes & "/" & ano
                dataSM = CDate(strData)
                
                ' If the date has already been read, try to update the remittance date
                If dataSM < dataSeparacao Then
                    dataRemessa = dataRemessa + WorksheetFunction.RoundUp((dataSeparacao - dataSM) / 2, 0)
                    If Not SetSAPText(Session, "wnd[0]/usr/tabsTAXI_TABSTRIP_SHEDLINE/tabpT\02/" & _
                            "ssubSUBSCREEN_BODY:SAPMV45A:4552/ctxtRV45A-ETDAT", Format(dataRemessa, "dd.mm.yyyy")) Then Exit Do
                    SAPSendVKey Session, 0
                Else
                    ' When the SAP date is now equal to or later than the desired date, exit the loop.
                    Exit Do
                End If
                
                loopCount = loopCount + 1
                If loopCount > MAX_LOOP Then
                    ws.Cells(i, ColumnWarningsStart).Value = "ERRO"
                    ws.Cells(i, ColumnWarningsStart + 1).Value = "Exceeded max iterations adjusting date"
                    Exit Do
                End If
            Loop
        
            '--- Save changes ---
            On Error Resume Next
            SAPPress Session, "wnd[0]/tbar[0]/btn[11]" ' Save
            SAPSendVKey Session, 0
            msg = Session.findbyid("wnd[0]/sbar/pane[0]").Text
            If msg = "OV c/ Fornec.Completo. As datas de todas as linhas devem ser iguais." Then
                SAPPress Session, "wnd[0]/tbar[0]/btn[3]"
                SAPPress Session, "wnd[0]/tbar[0]/btn[3]"
                Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                    "subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-ETDAT[7,0]").Text = "25.03.2020"
                Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                    "subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-ETDAT[7,1]").Text = "25.03.2020"
                SAPPress Session, "wnd[0]/tbar[0]/btn[11]"
                SAPPress Session, "wnd[1]/tbar[0]/btn[0]"
                j = 0
                Do
                    j = j + 1
                Loop Until Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                    "subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0," & j - 1 & "]").Text = currentItem
            End If
            
            '--- Update workbook with results ---
            ws.Cells(i, ColumnWarningsStart).Value = dataRemessa
            ws.Cells(i, ColumnWarningsStart + 1).Value = "CONCLUÍDO - " & Now()
            ws.Cells(i, ColumnWarningsStart + 2).Value = Session.findbyid("wnd[0]/sbar").Text
            
        ElseIf InStr(1, ws.Range("D" & i).Value, "DATA DADOS AD. B") > 0 Then
ErrorSection = "DataDadosAdBProcessing-" & i
            On Error GoTo IterationError
            
            orderNumber = ws.Range("A" & i).Value
            itemCode = ws.Range("B" & i).Value
            currentItem = ""
            
            '--- Clear workbook values ---
            ws.Cells(i, ColumnWarningsStart).Value = ""
            ws.Cells(i, ColumnWarningsStart + 1).Value = ""
            ws.Cells(i, ColumnWarningsStart + 2).Value = ""
            
            ' Update status bar with progress info
            Application.StatusBar = "Executando... Ordem de Venda " & orderNumber & " (" & i - 1 & " de " & lastRow - 1 & ")"
            
            '--- Open transaction VA02 ---
            If Not SetSAPText(Session, "wnd[0]/tbar[0]/okcd", "/nVA02") Then GoTo NextIteration
            SAPSendVKey Session, 0
            
            '--- Enter the Sales Order number ---
            If Not SetSAPText(Session, "wnd[0]/usr/ctxtVBAK-VBELN", orderNumber) Then GoTo NextIteration
            SAPSendVKey Session, 0
            
            '--- Avoid Info Pop-up ---
            On Error Resume Next
            If Session.findbyid("wnd[1]").Text = "Informação" Then
                Session.findbyid("wnd[1]/tbar[0]/btn[0]").Press
            End If
            On Error GoTo IterationError
            
            '--- Select first item and open it for editing ---
            On Error GoTo IterationError
            Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                "subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,0]").SetFocus
            SAPPress Session, "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                "subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_PEIN"
            
ajusteDataDadosB:

            Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
        
            currentItem = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/txtVBAP-POSNR").Text
        
            Do While itemCode <> currentItem
                SAPPress Session, "wnd[0]/tbar[1]/btn[19]"
                
                currentItem = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/txtVBAP-POSNR").Text
            Loop
            
            '--- Set separation and remittance dates based on the workbook ---
            dataSeparacao = ws.Range("C" & i).Value
            dataRemessa = dataSeparacao
            
            Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cntlCT_ALV_8459/shellcont/shell").modifyCell 0, "DT_EXPEC_FAT", Format(dataSeparacao, "dd.mm.yyyy")
            Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cntlCT_ALV_8459/shellcont/shell").modifyCell 0, "DS_VALOR", "PRODUÇÃO"
            Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/cntlCT_ALV_8459/shellcont/shell").modifyCell 0, "DT_EXPEC_REC", Format(dataSeparacao, "dd.mm.yyyy")
            
            SAPSendVKey Session, 0
            
            If Session.findbyid("wnd[0]/sbar").Text <> "" And InStr(1, Session.findbyid("wnd[0]/sbar").Text, "Não existem part.vencida") = 0 Then
                GoTo IterationError
            End If
        
            '--- Save changes ---
            On Error Resume Next
            SAPPress Session, "wnd[0]/tbar[0]/btn[11]" ' Save
            ' SAPSendVKey Session, 0
            msg = Session.findbyid("wnd[0]/sbar/pane[0]").Text
            If msg = "OV c/ Fornec.Completo. As datas de todas as linhas devem ser iguais." Then
                SAPPress Session, "wnd[0]/tbar[0]/btn[3]"
                SAPPress Session, "wnd[0]/tbar[0]/btn[3]"
                Stop
                'Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                    "subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-ETDAT[7,0]").Text = "25.03.2020"
                'Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                    "subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-ETDAT[7,1]").Text = "25.03.2020"
                SAPPress Session, "wnd[0]/tbar[0]/btn[11]"
                SAPPress Session, "wnd[1]/tbar[0]/btn[0]"
                j = 0
                Do
                    j = j + 1
                Loop Until Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/" & _
                    "subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0," & j - 1 & "]").Text = currentItem
            End If
            
            '--- Update workbook with results ---
            ws.Cells(i, ColumnWarningsStart).Value = dataSeparacao
            ws.Cells(i, ColumnWarningsStart + 1).Value = "CONCLUÍDO - " & Now()
            ws.Cells(i, ColumnWarningsStart + 2).Value = Session.findbyid("wnd[0]/sbar").Text
        Else
ErrorSection = "IgnoredProcessing-" & i
            '--- Update workbook with results ---
            ws.Cells(i, ColumnWarningsStart).Value = "IGNORADO"
            ws.Cells(i, ColumnWarningsStart + 1).Value = "IGNORADO - " & Now()
            ws.Cells(i, ColumnWarningsStart + 2).Value = ""
        End If
        
NextIteration:

        Application.StatusBar = ""
        
    Next i
    
ErrorSection = "Ending"

    Application.StatusBar = False
    
    'Call GerarClaim
    
    EndSAPScripting
    
    GoTo CleanExit

'----------------------------------------------
' Error Handling
'----------------------------------------------
IterationError:
    ws.Range("G" & i) = "ERRO"
    ws.Range("H" & i) = "Erro ao processar o item: " & NextItem & " - " & Session.findbyid("wnd[0]/sbar").Text
    Resume NextIteration
    
CleanExit:
    ' Ensure that all optimizations are turned back on
    OptimizeCodeExecution False
    
    Exit Sub

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in ExampleSub"
    
    ' Optionally, you can log the error details to a file or a logging system here
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Sub

Sub GerarClaim(Optional ShowOnMacroList = False)

    Stop

    Dim linha As Long
    Dim SapGuiAuto As Object
    Dim Application As Object
    Dim Connection As Object
    Dim Session As Object
    Dim item As String
    
    'Conexão com o Objeto SAP
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Application = SapGuiAuto.GetScriptingEngine
    Set Connection = Application.Children(0)
    Set Session = Connection.Children(0)
    
    On Error Resume Next
    
    'Verificar se a tela inicial do SAP ou a CLM1 está ativada através do SAPLogon
    'Do
        'If (session.info.transaction <> "SESSION_MANAGER") And (session.info.transaction <> "CLM1") Then
            'Resp = MsgBox("Deixe aberta apenas a janela principal do SAP ou a transação CLM1, ativada através do SAPLogon. Feche todas as demais janelas do SAP!" & vbNewLine & vbNewLine & "Click em OK quando desejar continuar, ou CANCELAR para encerrar o programa.", vbOKCancel + vbExclamation, "Mensagem do Programa")
            'If Resp = 2 Then
                'MsgBox "Script cancelado!", vbExclamation, "Mensagem do Programa"
                'Exit Sub
            'End If
        'End If
    'Loop Until (session.info.transaction = "SESSION_MANAGER") Or (session.info.transaction = "CLM1")

    'Se janela ativa é a janela principal do SAP, entrar na transação CLM1
    'If session.info.transaction = "SESSION_MANAGER" Then
        Session.findbyid("wnd[0]").maximize
        Session.findbyid("wnd[0]/tbar[0]/okcd").Text = "/nclm1"
        Session.findbyid("wnd[0]").sendVKey 0
    'End If
    
    'Contar o número de linhas
    ThisWorkbook.Activate
    Worksheets("Script").Activate
    linha = Range("A1").End(xlDown).Row
    
    'Loop de preenchimento e alteração
    For i = 2 To linha
  
        Session.findbyid("wnd[0]").maximize
        Session.findbyid("wnd[0]/tbar[0]/okcd").Text = "/nclm1"
        Session.findbyid("wnd[0]").sendVKey 0
        Session.findbyid("wnd[0]/usr/cmbRIWO00-QMART").Key = "ZZ"
        Session.findbyid("wnd[0]").sendVKey 0
        Session.findbyid("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7715/cntlTEXT/shellcont/shell").Text = "Favor alterar data de preparação da OV  " + Plan1.Range("A" & (i)).Value + " conforme segue:" + vbCr + "" + vbCr + "DE:" + Planilha1.Range("G" & (i)).Value + "   PARA:" + Plan1.Range("C" & (i)).Value + vbCr + "" + vbCr + "" + vbCr + ""
        Session.findbyid("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7715/cntlTEXT/shellcont/shell").setSelectionIndexes 126, 126
        Session.findbyid("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7715/txtRIWO00-HEADKTXT").Text = "Alteração de data do gerador" & " " & Plan1.Range("A" & (i)).Value
        Session.findbyid("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPMCLAIM:7800/ctxtCLAIM-URGRP").SetFocus
        Session.findbyid("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPMCLAIM:7800/ctxtCLAIM-URGRP").caretPosition = 0
        Session.findbyid("wnd[0]").sendVKey 4
        Session.findbyid("wnd[1]/usr/cntlTREE_CONTROL_AREA/shellcont/shell").expandNode "         73"
        Session.findbyid("wnd[1]/usr/cntlTREE_CONTROL_AREA/shellcont/shell").topNode = "          1"
        Session.findbyid("wnd[1]/usr/cntlTREE_CONTROL_AREA/shellcont/shell").selectItem "         74", "3"
        Session.findbyid("wnd[1]/usr/cntlTREE_CONTROL_AREA/shellcont/shell").ensureVisibleHorizontalItem "         74", "3"
        Session.findbyid("wnd[1]/usr/cntlTREE_CONTROL_AREA/shellcont/shell").doubleClickItem "         74", "3"
        Session.findbyid("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7740/subBELEG:SAPMCLAIM:3071/ctxtVIQMEL-LS_KDAUF").Text = Plan1.Range("A" & (i)).Value
        Session.findbyid("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7740/subBELEG:SAPMCLAIM:3071/ctxtVIQMEL-LS_KDPOS").Text = Plan1.Range("B" & (i)).Value
        Session.findbyid("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7740/subBELEG:SAPMCLAIM:3071/ctxtVIQMEL-LS_KDPOS").SetFocus
        Session.findbyid("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7740/subBELEG:SAPMCLAIM:3071/ctxtVIQMEL-LS_KDPOS").caretPosition = 2
        Session.findbyid("wnd[0]").sendVKey 0
        Session.findbyid("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7715/cntlTEXT/shellcont/shell").Text = "Favor alterar data de preparação da OV" & " " & Plan1.Range("A" & (i)).Value & "                            " & "DE:" & " " & Planilha1.Range("G" & (i)).Value & "                   " & "PARA:" & " " & Plan1.Range("C" & (i)).Value
        Session.findbyid("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7715/cntlTEXT/shellcont/shell").setSelectionIndexes 148, 148
        Session.findbyid("wnd[0]/shellcont/shell").selectItem "0010", "Column01"
        Session.findbyid("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem "0010", "Column01"
        Session.findbyid("wnd[0]/shellcont/shell").clickLink "0010", "Column01"
        Session.findbyid("wnd[1]/usr/cntlCONTAINER/shellcont/shell[0]").expandNode "HWAU"
        Session.findbyid("wnd[1]/usr/cntlCONTAINER/shellcont/shell[0]").expandNode "623SOPCPWAU"
        Session.findbyid("wnd[1]/usr/cntlCONTAINER/shellcont/shell[0]").selectedNode = "000000003603"
        Session.findbyid("wnd[1]/usr/cntlCONTAINER/shellcont/shell[0]").topNode = "617USPVENWA"
        Session.findbyid("wnd[1]/usr/cntlCONTAINER/shellcont/shell[0]").doubleClickNode "000000003603"
        Session.findbyid("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").modifyCell 0, "RESP", "kamila"
        Session.findbyid("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").setCurrentCell -1, ""
        Session.findbyid("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").firstVisibleColumn = "SUB_GRUPO"
        'session.findById("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").selectColumn "GRUPO"
        'session.findById("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").selectColumn "SUB_GRUPO"
        'session.findById("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").selectColumn "DS_ATIVIDADE"
        'session.findById("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").selectColumn "SEQ"
        'session.findById("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").selectColumn "DT_LIMITE"
        'session.findById("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").selectColumn "HR_LIMITE"
        'session.findById("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").selectColumn "TP_RESP"
        'session.findById("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").selectColumn "RESP" '
        Session.findbyid("wnd[1]/usr/cntlCONTAINER2/shellcont/shell").selectedRows = "0"
        Session.findbyid("wnd[1]/tbar[0]/btn[44]").Press
        Session.findbyid("wnd[0]/tbar[1]/btn[14]").Press
        Session.findbyid("wnd[0]/tbar[0]/btn[11]").Press
        Session.findbyid("wnd[0]/sbar").DoubleClick
        Range("F" & i) = Session.findbyid("wnd[0]/sbar").Text
       
    Next
   
End Sub

'----------------------------------------------
' Helper Functions for SAP GUI operations
'----------------------------------------------

' Safely sets the Text property of a SAP control.
Function SetSAPText(Session As Object, controlId As String, textValue As String) As Boolean
    On Error GoTo ErrHandler
    Session.findbyid(controlId).Text = textValue
    SetSAPText = True
    Exit Function
ErrHandler:
    SetSAPText = False
End Function

' Safely simulates a button press on a SAP control.
Function SAPPress(Session As Object, controlId As String) As Boolean
    On Error GoTo ErrHandler
    Session.findbyid(controlId).Press
    SAPPress = True
    Exit Function
ErrHandler:
    SAPPress = False
End Function

' Safely sends a VKey command (for example, simulating Enter) to SAP.
Function SAPSendVKey(Session As Object, keyCode As Integer) As Boolean
    On Error GoTo ErrHandler
    Session.findbyid("wnd[0]").sendVKey keyCode
    SAPSendVKey = True
    Exit Function
ErrHandler:
    SAPSendVKey = False
End Function

Function SetupSAPScripting() As Boolean
    
    Dim isHomePage As Boolean
    
    ' Create the SAP GUI scripting engine object
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    On Error GoTo 0
    
    If Not IsObject(SapGuiAuto) Or SapGuiAuto Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    On Error Resume Next
    Set SAPApplication = SapGuiAuto.GetScriptingEngine
    On Error GoTo 0
    
    If Not IsObject(SAPApplication) Or SAPApplication Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    ' Get the first connection and session
    Set Connection = SAPApplication.Children(0)
    Set Session = Connection.Children(0)
    
    If Connection Is Nothing Or Session Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    SetupSAPScripting = True
    
End Function

Function EndSAPScripting()
    ' Clean up
    Set Session = Nothing
    Set Connection = Nothing
    Set SAPApplication = Nothing
    Set SapGuiAuto = Nothing
End Function

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
