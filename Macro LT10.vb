'DE MANEIRA MUITO MAIS SIMPLES, REALIZA A MESMA FUNÇÃO DA LT06. 
'PORÉM AO INVES DE UMA QUANTIDADE DETERMINADA, A LT10 CONCLUI A BAIXA/TRANSFERÊNCIA DA QUANTIDADE TOTAL DO ITEM NO ESTOQUE.
'APESAR DE SER MAIS RÁPIDA E COM POUCA ENTRADA DE DADOS, DEVE SER USADA COM CAUTELA.

Private Sub LT10_Click()

    'DECLARANDO VARIÁVEIS
    Dim NI As String
    Dim POSICAO_DESTINO As String
    Dim POSICAO_ANTIGA As String
    Dim POSICAO_VERIFICAR As String

    L = 4
    NI = Sheets("LT10").Range("B" & L).Value
    POSICAO_DESTINO = Sheets("LT10").Range("C" & L).Value
    POSICAO_ANTIGA = Sheets("LT10").Range("D" & L).Value

    'INCIANDO SAP
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

    Sheets("LT10").Range("B" & L).Select

    'INCIAR LOOP DA TRANSAÇÃO LT10
    Do While ActiveCell <> ""
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NLT10"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtS1_LGNUM").Text = "FL2"
        session.findById("wnd[0]/usr/ctxtS1_LGTYP-LOW").Text = "*"
        session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = NI
        session.findById("wnd[0]/tbar[1]/btn[8]").press

        N = 6
        session.findById("wnd[0]").sendVKey 35
        M = session.findById("wnd[1]/usr/lbl[26,3]").Text
        M = M + 5
        session.findById("wnd[1]").Close
        session.findById("wnd[0]").sendVKey 45

        'PERCORRER A LISTA DE ESTOQUES
        Do While N <= M
            POSICAO_VERIFICAR = session.findById("wnd[0]/usr/lbl[9," & N & "]").Text
            If session.findById("wnd[0]/usr/lbl[9," & N & "]").Text <> POSICAO_ANTIGA Then
                session.findById("wnd[0]/usr/lbl[2," & N & "]").SetFocus
                session.findById("wnd[0]").sendVKey 2
            End If
            N = N + 1
        Loop

        On Error GoTo -1
        On Error GoTo 0
        On Error GoTo PRÓX

        session.findById("wnd[0]/tbar[1]/btn[48]").press
        session.findById("wnd[1]/usr/ctxtLAGP-LGTYP").Text = "AC"
        session.findById("wnd[1]/usr/ctxtLAGP-LGPLA").Text = POSICAO_DESTINO
        session.findById("wnd[1]").sendVKey 0

        'PRÓXIMA LINHA
        PRÓX:
        L = L + 1
        NI = Sheets("LT10").Range("B" & L).Value
        POSICAO_DESTINO = Sheets("LT10").Range("C" & L).Value
        POSICAO_ANTIGA = Sheets("LT10").Range("D" & L).Value
        Sheets("LT10").Range("B" & L).Select

    Loop

    MsgBox ("Processo concluído!")

End Sub
