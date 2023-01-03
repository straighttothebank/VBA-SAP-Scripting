'MACRO DESENVOLVIDA PARA ALTERAÇÃO DE DESCRIÇÃO NOS RELATÓRIOS DE RECEBIMENTOS DOS MATERIAIS

Private Sub MB02_Click()

    'ALOCACAO DOS DADOS QUE SERAO PREENCHIDOS NO SAP'
    L = 4 'NUMERO DA LINHA QUE COMEÇA OS DADOS'
    RR = Sheets("MB02").Range("B" & L).Value
    CABECALHO = Sheets("MB02").Range("C" & L).Value

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

    Sheets("MB02").Range("B" & L).Select

    Do While ActiveCell <> "" 'O SCRIPT VAI RODAR ATE CHEGAR NUMA CELULA VAZIA

    'EXECUTAR A TRANSACAO
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "MB02"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRM07M-MBLNR").Text = RR
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[0]/usr/txtMKPF-BKTXT").Text = PALETE
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[0]/btn[15]").press

    'PINTAR CÉLULAS
    Sheets("MB02").Range("B" & L).Select
        Range("B4:C4").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With

    Sheets("MB02").Range("C" & L).Select
        Range("B4:C4").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With


    'PULAR A LINHA DA PLANILHA E INICIAR COM OUTRO RR
    L = L + 1
    RR = Sheets("MB02").Range("B" & L).Value
    PALETE = Sheets("MB02").Range("C" & L).Value
    Sheets("MB02").Range("B" & L).Select

    Loop
 
End Sub