'MACRO DESENVOLVIDA PARA REALIZAR BAIXA DE MATERIAIS DE PROJETOS
'ESSES MATERIAIS POSSUEM UM CÓDIGO CHAMADO "PEP"
'PARA OBTER O CÓDIGO PEP DOS MATERIAIS FOI ELABORADA UMA FUNÇÃO QUE UTILIZA A TRANSAÇÃO LS24

'FUNÇÃO LS24
Private Sub FIND_PEP_Click()

    L = 5
    MATERIAL = Sheets("MB1A").Range("B" & L).Value
    CENTRO = Sheets("MB1A").Range("F" & L).Value

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

    Sheets("MB1A").Range("B" & L).Select

    Do While ActiveCell <> ""

        If Sheets("MB1A").Range("H" & L).Value <> "" Then
            GoTo PRÓX
        End If

        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NLS24"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtRL01S-LGNUM").Text = "FB1"
        session.findById("wnd[0]/usr/ctxtRL01S-MATNR").Text = MATERIAL
        session.findById("wnd[0]/usr/ctxtRL01S-WERKS").Text = CENTRO
        session.findById("wnd[0]/usr/txtRL01S-LGORT").Text = "AM01"
        session.findById("wnd[0]/usr/ctxtRL01S-SOBKZ").Text = "Q"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/lbl[46,8]").SetFocus
        session.findById("wnd[0]").sendVKey 2
        Sheets("MB1A").Range("H" & L).Value = session.findById("wnd[0]/usr/txtRL01S-LSONR").Text

        PRÓX:
        L = L + 1
        MATERIAL = Sheets("MB1A").Range("B" & L).Value
        CENTRO = Sheets("MB1A").Range("F" & L).Value
        Sheets("MB1A").Range("B" & L).Select

    Loop

    session.findById("wnd[0]/tbar[0]/okcd").Text = "/N"
    session.findById("wnd[0]").sendVKey 0
    MsgBox ("Processo concluído! Confira o PEP antes de realizar a baixa.")

End Sub

'###############################################################################

'FUNÇÃO MB1A
Private Sub EXECUTAR_MB1A_Click()


    'INICIALIZAÇÃO DAS VARIÁVEIS
    L = 5
    MATERIAL = Sheets("MB1A").Range("B" & L).Value
    QUANTIDADE = Sheets("MB1A").Range("C" & L).Value
    UNIDADE = Sheets("MB1A").Range("D" & L).Value
    LOTE = Sheets("MB1A").Range("E" & L).Value
    CENTRO = Sheets("MB1A").Range("F" & L).Value
    SOLICITANTE = Sheets("MB1A").Range("G" & L).Value
    ELEMENTO_PEP = Sheets("MB1A").Range("H" & L).Value

    'INICIALIZAÇÃO DO SAP
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

    session.findById("wnd[0]").maximize

    Sheets("MB1A").Range("B" & L).Select

    'LOOP PRINCIPAL DE TRANSAÇÃO
    Do While ActiveCell <> ""

        On Error GoTo -1
        On Error GoTo 0
        On Error GoTo PROX

        If Sheets("MB1A").Range("I" & L).Value <> "OK" Then

            session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB1A"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/txtMKPF-BKTXT").Text = SOLICITANTE
            session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "221"
            session.findById("wnd[0]/usr/ctxtRM07M-SOBKZ").Text = "Q"
            session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = CENTRO
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/subBLOCK1:SAPMM07M:2424/ctxtMSEGK-MAT_PSPNR").Text = ELEMENTO_PEP
            session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").Text = MATERIAL
            session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").Text = QUANTIDADE
            session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-ERFME[0,44]").Text = UNIDADE
            session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-LGORT[0,48]").Text = "AM01"
            session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").Text = LOTE
            session.findById("wnd[0]/tbar[0]/btn[11]").press

            'SEGUNDA TELA - LT06

            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/ctxtT334T-LGTY0").Text = "902"
            session.findById("wnd[0]/tbar[1]/btn[7]").press
            session.findById("wnd[0]/usr/tabsFUNC_TABSTRIP/tabpAQVB/ssubD0106_S:SAPML03T:1061/tblSAPML03TD1061/txtRL03T-SELMG[0,0]").Text = QUANTIDADE
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[0]/btn[11]").press

            Sheets("MB1A").Range("I" & L).Value = "OK"

        End If


        '-----------------------

        'PRÓXIMA LINHA
        PROX:
        L = L + 1
        MATERIAL = Sheets("MB1A").Range("B" & L).Value
        QUANTIDADE = Sheets("MB1A").Range("C" & L).Value
        UNIDADE = Sheets("MB1A").Range("D" & L).Value
        LOTE = Sheets("MB1A").Range("E" & L).Value
        CENTRO = Sheets("MB1A").Range("F" & L).Value
        SOLICITANTE = Sheets("MB1A").Range("G" & L).Value
        ELEMENTO_PEP = Sheets("MB1A").Range("H" & L).Value

        Sheets("MB1A").Range("B" & L).Select

    Loop

    MsgBox ("Processo concluído!")

End Sub