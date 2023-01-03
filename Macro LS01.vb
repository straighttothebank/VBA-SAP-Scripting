Private Sub LS01_Click()

'MACRO DESENVOLVIDA PARA CRIAR DIVERSAS POSIÇÕES DE ESTOQUE DENTRO DO SISTEMA SAP

    'INICIALIZAÇÃO DE VARIÁVEIS
    L = 4
    NUMERO_DEP = Sheets("LS01").Range("B" & L).Value
    TIPO_DEP = Sheets("LS01").Range("C" & L).Value
    POSICAO_DEP = Sheets("LS01").Range("D" & L).Value

    'INICIALIZAR O SAP
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

    'SELECIONAR PRIMEIRA CÉLULA
    Sheets("LS01").Range("B4").Select
    session.findById("wnd[0]").maximize

    'ABRIR TRANSAÇÃO E INICIAR LOOP PARA CRIAÇÃO DAS POSIÇÕES
    Do While ActiveCell <> ""
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NLS01"
        session.findById("wnd[0]").sendVKey 0

        On Error GoTo -1
        On Error GoTo 0
        On Error GoTo PRÓX
        session.findById("wnd[0]/usr/ctxtLAGP-LGNUM").Text = NUMERO_DEP
        session.findById("wnd[0]/usr/ctxtLAGP-LGTYP").Text = TIPO_DEP
        session.findById("wnd[0]/usr/ctxtLAGP-LGPLA").Text = POSICAO_DEP
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtLAGP-LGBER").Text = "001"
        session.findById("wnd[0]/tbar[0]/btn[11]").press

        'SEGUIR COM A PRÓXIMA LINHA
        PRÓX:
        L = L + 1
        NUMERO_DEP = Sheets("LS01").Range("B" & L).Value
        TIPO_DEP = Sheets("LS01").Range("C" & L).Value
        POSICAO_DEP = Sheets("LS01").Range("D" & L).Value
        Sheets("LS01").Range("B" & L).Select

    Loop

    MsgBox ("Processo concluído!")

End Sub
