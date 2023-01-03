'A ZMMCONT TEM COMO PRINCIPAL FUNÇÃO REALIZAR A CONFIRMAÇÃO LOGÍSTICA DE MATERIAIS RECEBIDOS.
'UTILIZANDO O NÚMERO DO RELATÓRIO DE RECEBIMENTO É POSSÍVEL REALIZAR A CONFIRMAÇÃO DE DIVERSOS ITENS.

Private Sub ZMMCONT_Click()

    'DECLARANDO AS VARIÁVEIS INICIAIS
    L = 4
    RR = Sheets("ZMMCONT").Range("D" & L).Value
    DEPOSITO = Sheets("ZMMCONT").Range("E" & L).Value
    CABECALHO = Sheets("ZMMCONT").Range("F" & L).Value
    Dim DESCRICAO As String

    'INCIALIZANDO SAP
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
    Sheets("ZMMCONT").Range("D" & L).Select

    'INICIAR TRANSAÇÃO ZMMCONT
    Do While ActiveCell <> ""
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZMMCONT"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/txtP_LGORT").Text = DEPOSITO
        session.findById("wnd[0]/usr/ctxtS_VBELN-LOW").Text = RR
        session.findById("wnd[0]/tbar[1]/btn[8]").press

        'PULAR PARA A PRÓXIMA LINHA NO CASO DE ERRO
        On Error GoTo -1
        On Error GoTo 0
        On Error GoTo PRÓX

        'MARCAR AS CAIXAS DE SELEÇÃO E CONFIRMAR LANÇAMENTO
        session.findById("wnd[0]/tbar[1]/btn[32]").press
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        session.findById("wnd[1]/usr/btnBUTTON_1").press

        'CRIAR TABELA A PARTIR DA SHELL
        Set TABELA = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        Dim COLUNAS As Object
        Set COLUNAS = TABELA.ColumnOrder
        DESCRICAO = TABELA.GetCellValue(1, COLUNAS(0))
        Sheets("ZMMCONT").Range("G" & L).Value = DESCRICAO

        'VERIFICAR SE É COMPRA DIRETA E INSERIR CABEÇALHO
        If CABECALHO <> "" Then
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB02"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/ctxtRM07M-MBLNR").Text = Right(DESCRICAO, 10)
            session.findById("wnd[0]/tbar[1]/btn[16]").press
            session.findById("wnd[0]/usr/txtMKPF-BKTXT").Text = CABECALHO
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            session.findById("wnd[0]/tbar[0]/btn[15]").press
        End If

        'PRÓXIMA LINHA
        PRÓX:
        L = L + 1
        RR = Sheets("ZMMCONT").Range("D" & L).Value
        DEPOSITO = Sheets("ZMMCONT").Range("E" & L).Value
        CABECALHO = Sheets("ZMMCONT").Range("F" & L).Value
        Sheets("ZMMCONT").Range("D" & L).Select

    Loop

    MsgBox ("Processo concluído!")

End Sub