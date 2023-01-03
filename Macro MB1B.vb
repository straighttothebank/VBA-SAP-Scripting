'A MB1B FOI DESENVOLVIDA PARA REALIZAR TRANSFERÊNCIA DE MATERIAIS

Private Sub MB1B_Click()

    Dim DESCRICAO As String
    Dim RECEBEDOR As String

    DESCRICAO = Sheets("MB1B").Range("C3").Text
    RECEBEDOR = Sheets("MB1B").Range("C4").Text

    'ALOCACAO DOS DADOS QUE SERAO PREENCHIDOS NA PRIMEIRA TELA DA TRANSAÇÃO'
    L = 8 'NUMERO DA LINHA QUE COMEÇA OS DADOS'
    NI = Sheets("MB1B").Range("B" & L).Value
    QUANTIDADE = Sheets("MB1B").Range("C" & L).Value
    SAIDA = Sheets("MB1B").Range("D" & L).Value
    DESTINO = Sheets("MB1B").Range("E" & L).Value

    'ABRIR O SAP'
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
        WScript.ConnectObject Application, "on"
    End If

    Sheets("MB1B").Range("B" & L).Select

    'FUNÇÃO PARA RODAR A MACRO ATÉ ATINGIR UMA CÉLULA VAZIA'
    Do While NI <> ""

        'EXECUTAR A TRANSAÇÃO E ATRIBUIR VALORES'
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "MB1B"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").Text = "311"
        session.findById("wnd[0]/usr/ctxtRM07M-WERKS").Text = "2100"
        session.findById("wnd[0]/usr/txtMKPF-BKTXT").Text = DESCRICAO
        session.findById("wnd[0]/usr/ctxtRM07M-LGORT").Text = SAIDA
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/txtMSEGK-WEMPF").Text = RECEBEDOR
        session.findById("wnd[0]/usr/ctxtMSEGK-UMLGO").Text = DESTINO
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").Text = NI
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").Text = QUANTIDADE
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press

        'FUNÇÃO PARA PULAR A LINHA E PROSSEGUIR COM O PRÓXIMO MATERIAL:
        L = L + 1

        'ATRIBUIR NOVOS DADOS DA LINHA ATUAL'

        'ALOCAÇÃO DO NÚMERO DE IDENTIFICAÇÃO'
        NI = Sheets("MB1B").Range("B" & L).Value
        QUANTIDADE = Sheets("MB1B").Range("C" & L).Value
        SAIDA = Sheets("MB1B").Range("D" & L).Value
        DESTINO = Sheets("MB1B").Range("E" & L).Value

    'FUNÇÃO PARA INICIAR TUDO NOVAMENTE'
    Loop

    Msgbox("Processo Concluído!")

End Sub
