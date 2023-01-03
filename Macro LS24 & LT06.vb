'ESSA MACRO É CRIADA PARA COMPLETAR BAIXAS NÃO FINALIZADAS DOS MATERIAIS EM ESTOQUE
'PARA ISSO FOI CRIADO DUAS FUNÇÕES PRINCIPAIS: UMA FUNÇÃO PARA BUSCAR O DOCUMENTO PENDENTE (LS24)
'E OUTRA FUNÇÃO PARA FINALIZAR A BAIXA ATRAVÉS DO NÚMERO DO DOCUMENTO ENCONTRADO (LT06)

'FUNÇÃO LS24
Private Sub LS24_Click()
    
    'INICIALIZAR VARIÁVEIS
    L = 5
    Dim DOCUMENTO As String
    Dim POSICAO As String
    Dim VERIFICAR As String
    Dim CARACT_VERIFICADA() As String

    DEPOSITO = Sheets("LT06").Range("B" & L).Value
    MATERIAL = Sheets("LT06").Range("C" & L).Value
    QUANTIDADE = Sheets("LT06").Range("D" & L).Value
    POSICAO = Sheets("LT06").Range("E" & L).Text

    '   INICIAR SAP
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

    Sheets("LT06").Range("B" & L).Select
    session.findById("wnd[0]").maximize

    ' OBTER NÚMEROS DOS DOCUMENTOS POR MEIO DA TRANSAÇÃO LS24

    Do While ActiveCell <> ""

        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NLS24"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtRL01S-LGNUM").Text = "FB1"
        session.findById("wnd[0]/usr/ctxtRL01S-MATNR").Text = MATERIAL
        session.findById("wnd[0]/usr/ctxtRL01S-WERKS").Text = "2100"
        session.findById("wnd[0]/usr/txtRL01S-LGORT").Text = DEPOSITO
        session.findById("wnd[0]/usr/ctxtRL01S-LGPLA").Text = POSICAO
        session.findById("wnd[0]/usr/txtRL01S-CHARG").caretPosition = 0
        session.findById("wnd[0]").sendVKey 0
        On Error GoTo -1
        On Error GoTo 0
        On Error GoTo PROX

        M = 8
        X = 0

        Do While X = 0

            On Error GoTo -1
            On Error GoTo 0
            On Error GoTo PROX
            VERIFICAR = session.findById("wnd[0]/usr/lbl[54," & M & "]").Text

            '---------- FUNÇÃO PARA REMOVER OS ZEROS E AJEITAR A NOTA FISCAL OBTIDA NO SAP --------------
            CARACT_VERIFICADA = Split(StrConv(VERIFICAR, vbUnicode), Chr$(0))
            ReDim Preserve CARACT_VERIFICADA(UBound(CARACT_VERIFICADA) - 1)

            'CONTAR NÚMERO DE ESPAÇOS À ESQUERDA
            ESPACOS = 0
            For Each C In CARACT_VERIFICADA
                If C = " " Then
                    ESPACOS = ESPACOS + 1
                Else
                    Exit For
                End If
            Next C

            'REMOVENDO ZEROS NA ESQUERDA
            VERIFICAR = Right(VERIFICAR, Len(VERIFICAR) - ESPACOS)

            If QUANTIDADE = VERIFICAR Then
                DOCUMENTO = session.findById("wnd[0]/usr/lbl[92," & M & "]").Text
                ANO = session.findById("wnd[0]/usr/lbl[103," & M & "]").Text
                ANO = Right(ANO, 4)
                Sheets("LT06").Range("F" & L).Value = DOCUMENTO
                Sheets("LT06").Range("G" & L).Value = ANO
                X = 1
            End If

            M = M + 1

        Loop

        'PULAR PARA A PRÓXIMA LINHA
        PROX:
        L = L + 1
        DEPOSITO = Sheets("LT06").Range("B" & L).Value
        MATERIAL = Sheets("LT06").Range("C" & L).Value
        QUANTIDADE = Sheets("LT06").Range("D" & L).Value
        POSICAO = Sheets("LT06").Range("E" & L).Text

        Sheets("LT06").Range("B" & L).Select

    Loop

    MsgBox ("Processo concluído!")

End Sub

'########################################################################'

'FUNÇÃO LT06
Private Sub LT06_Click()

    'INCIALIZANDO VARIÁVEIS
    L = 5
    Dim POSICAO As String
    Dim TIPO As String
    POSICAO = Sheets("LT06").Range("I" & L).Text
    TIPO = Sheets("LT06").Range("H" & L).Text
    DOCUMENTO = Sheets("LT06").Range("F" & L).Value
    ANO = Sheets("LT06").Range("G" & L).Value

    '   INICIAR SAP
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

    Sheets("LT06").Range("B" & L).Select

    'INICIAR LOOP DA TRANSAÇÃO
    Do While ActiveCell <> ""
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NLT06"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/txtRL02B-MBLNR").Text = DOCUMENTO
        session.findById("wnd[0]/usr/txtRL02B-MJAHR").Text = ANO

        If DOCUMENTO = "" Then
            GoTo PROXI
        End If

        session.findById("wnd[0]").sendVKey 0

        'PROCESSAR LT06 -------

        On Error GoTo -1
        On Error GoTo 0
        On Error GoTo PROXI
        session.findById("wnd[0]/usr/ctxtT334T-LGTY0").Text = TIPO
        session.findById("wnd[0]/tbar[1]/btn[5]").press

        On Error GoTo -1
        On Error GoTo 0
        On Error GoTo PROXI
        session.findById("wnd[0]/usr/ctxtLTAP-VLTYP").Text = TIPO
        session.findById("wnd[0]/usr/txtLTAP-VLPLA").Text = POSICAO
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        Sheets("LT06").Range("J" & L).Value = "OK"

        '----------------------

        'IR PARA A PRÓXIMA LINHA
        PROXI:
        L = L + 1
        POSICAO = Sheets("LT06").Range("I" & L).Text
        TIPO = Sheets("LT06").Range("H" & L).Text
        DOCUMENTO = Sheets("LT06").Range("F" & L).Value
        ANO = Sheets("LT06").Range("G" & L).Value
        Sheets("LT06").Range("B" & L).Select

    Loop

    MsgBox ("Processo concluído!")

End Sub