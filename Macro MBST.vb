'MACRO DESENVOLVIDA PARA REALIZAR ESTORNOS DE BAIXAS DE MATERIAIS
'
Private Sub MBST_Click()

    'ALOCANDO VARIÁVEIS
    Dim DOCUMENTO As String
    L = 3
    DOCUMENTO = Sheets("MBST").Range("B" & L).Value
    
    Sheets("MBST").Range("B" & L).Select

    'INICIALIZAR SAP

    If Not IsObject(SAPApplication) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SAPApplication = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
        Set Connection = SAPApplication.Children(0)
    End If
    If Not IsObject(session) Then
        Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject SAPApplication, "on"
    End If

    'EXECUTAR A TRANSAÇÃO

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMBST"
    session.findById("wnd[0]").sendVKey 0

    'REALIZAR UM LOOP ATÉ ATINGIR UMA CELULA VAZIA

    Do While ActiveCell <> ""

        session.findById("wnd[0]/usr/ctxtRM07M-MBLNR").Text = DOCUMENTO
        session.findById("wnd[0]/usr/ctxtRM07M-MBLNR").caretPosition = 10
        session.findById("wnd[0]/tbar[1]/btn[6]").press
        session.findById("wnd[0]/usr/txtMSEG-SGTXT").Text = "MATERIAL ESTORNADO"
        session.findById("wnd[0]/usr/txtMSEG-SGTXT").SetFocus
        session.findById("wnd[0]/usr/txtMSEG-SGTXT").caretPosition = 18
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press

        'ALOCANDO AS NOVOS VALORES E DESCENDO AS LINHAS
        L = L + 1
        DOCUMENTO = Workbooks("ESTORNO DE DOCUMENTOS.XLSM").Sheets("Plan1").Range("B" & L).Value
        Workbooks("ESTORNO DE DOCUMENTOS.XLSM").Sheets("Plan1").Range("B" & L).Select

    Loop

    session.findById("wnd[0]/tbar[0]/btn[15]").press

End Sub