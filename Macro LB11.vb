Private Sub LB11_Click()

'MACRO DESENVOLVIDA PARA REALIZAR LIMPEZA DE POSIÇÕES TEMPORÁRIAS DENTRO DO ESTOQUE.
'DEVIDO AO GRANDE VOLUME DE MOVIMENTAÇÕES, ALGUMAS ORDENS DE TRANSIÇÕES GERAM SUJEIRA NO ESTOQUE.
'ESSA SUJEIRA GERADA POR BAIXAS E MOVIMENTAÇÃO DE MATERIAIS PODE COMPROMETER AS OPERAÇÕES NO SISTEMA.

    'VARIÁVEIS DE SUPORTE PRINCIPAIS
    L = 4 'Nº DA LINHA INICIAL
    M = 6 'VARIÁVEL DE POSIÇÃO DO SCROLL
    SCR = 6 'VARIÁVEL DE POSIÇÃO ALTERNADA DO SCROLL

    'INICIAR SAP 
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

    'DECLARANDO VARIÁVEIS DA PRIMEIRA LINHA
    MATERIAL = Sheets("LB11").Range("B" & L).Value
    LOTE = Sheets("LB11").Range("C" & L).Value
    TIPO = Sheets("LB11").Range("D" & L).Value
    DEPOSITO = Sheets("LB11").Range("E" & L).Value
    TEMPORARIO = Sheets("LB11").Range("F" & L).Value
    REAL = Sheets("LB11").Range("G" & L).Value

    session.findById("wnd[0]").maximize

    Inicio:
    Do While MATERIAL <> ""

        'INCIAR A TRANSAÇÃO LB11
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NLB11"
        session.findById("wnd[0]").sendVKey 0

        'INSERIR DADOS INCIAIS
        On Error GoTo -1
        On Error GoTo 0
        session.findById("wnd[0]/usr/ctxtRL02B-LGNUM").Text = "FB1"
        session.findById("wnd[0]/usr/ctxtRL02B-MATNR").Text = MATERIAL
        session.findById("wnd[0]/usr/ctxtRL02B-WERKS").Text = "2100"
        session.findById("wnd[0]/usr/txtRL02B-LGORT").Text = "AM08"
        session.findById("wnd[0]").sendVKey 0
        On Error GoTo Reset

        session.findById("wnd[0]").sendVKey 35
        N = session.findById("wnd[1]/usr/lbl[20,3]").Text
        N = N + 6
        session.findById("wnd[1]").Close
        SCR = session.findById("wnd[0]/usr").verticalScrollbar.Position
        session.findById("wnd[0]/usr/lbl[67,4]").SetFocus
        session.findById("wnd[0]").sendVKey 2
        session.findById("wnd[0]/tbar[1]/btn[41]").press

        Do While M <= N

            'MARCAR A PRIMEIRA CAIXINHA
            session.findById("wnd[0]").maximize
            On Error GoTo -1
            On Error GoTo 0
            On Error GoTo LT10
            session.findById("wnd[0]/usr/chk[1," & M & "]").Selected = True

            'EXIBIR OTVISÍVEL
            session.findById("wnd[0]/tbar[1]/btn[44]").press

            'INFORMAR TIPO DA POSIÇÃO'
            session.findById("wnd[0]/usr/ctxtT334T-LGTY0").Text = "S"

            'RETIRAR DE DEPÓSITO VISÍVEL
            session.findById("wnd[0]/tbar[1]/btn[5]").press

            'INSERIR O DEPÓSITO PARA TRANSFERIR
            session.findById("wnd[0]/usr/txtLTAP-VLPLA").Text = "MERCADO"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]").sendVKey 0

            'SALVAR
            session.findById("wnd[0]/tbar[0]/btn[11]").press

            'DESCER A LINHA DAS CAIXINHAS
            M = M + 1
            SCR = SCR + 1

            'RESETAR A BARRA
            If SCR = 25 Then
                SCR = 6
                session.findById("wnd[0]/tbar[0]/okcd").Text = "/NLB11"
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/usr/ctxtRL02B-LGNUM").Text = "FB1"
                session.findById("wnd[0]/usr/ctxtRL02B-MATNR").Text = MATERIAL
                session.findById("wnd[0]/usr/ctxtRL02B-WERKS").Text = "2100"
                session.findById("wnd[0]/usr/txtRL02B-LGORT").Text = "AM08"
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]").sendVKey 35
                N = session.findById("wnd[1]/usr/lbl[20,3]").Text
                N = N + 6
                session.findById("wnd[1]").Close
                M = 6
                session.findById("wnd[0]/usr/lbl[67,4]").SetFocus
                session.findById("wnd[0]").sendVKey 2
                session.findById("wnd[0]/tbar[1]/btn[41]").press

            End If

        Loop

        L = L + 1
        MATERIAL = Sheets("LB11").Range("B" & L).Value

    Loop

    MsgBox "Processo finalizado!"

    Exit Sub

    LT10:
    Err.Clear
    Err.Number = 0
    If session.findById("wnd[0]/usr/txtLTAP-VLQNR").Text = "" Then
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NLT10"
        session.findById("wnd[0]").sendVKey 0

            'INSERIR DADOS NA LT10
        session.findById("wnd[0]/usr/ctxtS1_LGNUM").Text = "FB1"
        session.findById("wnd[0]/usr/ctxtS1_LGTYP-LOW").Text = "*"
        session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = MATERIAL
        session.findById("wnd[0]/tbar[1]/btn[8]").press

            'PROCURAR POR DOCUMENTOS NO DEPÓSITO 911
        session.findById("wnd[0]").sendVKey 35
        NL = session.findById("wnd[1]/usr/lbl[26,3]").Text
        
        If NL = 0 Then GoTo Reset

        NL = NL + 5
        session.findById("wnd[1]").Close

            'MARCAR TUDO
        session.findById("wnd[0]/tbar[1]/btn[45]").press
        D = 6
        X = 0

            'DESMARCAR TODOS QUE NÃO SÃO 911]
        Do While D <= NL
            PS = session.findById("wnd[0]/usr/lbl[5," & D & "]").Text
            If PS <> "911" Then
                session.findById("wnd[0]/usr/lbl[2," & D & "]").SetFocus
                session.findById("wnd[0]").sendVKey 2
                X = X + 1
            End If
            D = D + 1
        Loop

        If X = NL - 5 Then GoTo Reset

            'REALIZAR A BAIXA NO LT10
            session.findById("wnd[0]/tbar[1]/btn[48]").press
            session.findById("wnd[1]/usr/ctxtLAGP-LGTYP").Text = "S"
            session.findById("wnd[1]/usr/ctxtLAGP-LGPLA").Text = "MERCADO"
            session.findById("wnd[1]/tbar[0]/btn[0]").press

        End If

        M = 6
        SCR = 6
        GoTo Inicio

        Reset:
        M = 6
        SCR = 6
        L = L + 1
        MATERIAL = Sheets("LB11").Range("B" & L).Value
        If MATERIAL <> "" Then
            GoTo Inicio
        Else
            MsgBox "Processo concluído"
            Exit Sub
        End If
    End If

End Sub