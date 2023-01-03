'A MACRO ME23N & ZMMSRR01 FOI DESENVOLVIDA PARA ATENDER AS NECESSIDADES DE IDENTIFICAR OS MATERIAIS RECEBIDOS VIA RELATORIO DE RECEBIMENTO (RR).
'INICIALMENTE UTILIZA-SE A TRANSAÇÃO ME23N PARA OBTER AUTOMATICAMENTE AS RRs REFERENTES A UM MATERIAL DE DETERMINADA NOTA FISCAL.
'ASSIM QUE OBTIDO A RR, INICIA A IMPRESSÃO DO RELATÓRIO ATRAVÉS DA MACRO ZMMSRR01.

'FUNÇÃO ME23N
Private Sub ME23N_Click()

    'DIMENSIONANDO AS VARIAVEIS

    Dim PEDIDO As String
    Dim NOTAFISCAL As String
    Dim NI As String
    Dim DATA As String
    Dim VERIFICADA As String
    Dim CARACT_VERIFICADA() As String
    Dim ZEROS As Integer
    Dim ITEM As String
    Dim COPY() As String
    Dim d_VERIFICADA As String
    Dim P_MATERIAL As String


    L = 5 'LINHA EM QUE COMEÇA OS DADOS NO EXCEL

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
        WScript.ConnectObject Application, "on"
    End If

    'SELECIONAR A PRIMEIRA CÉLULA DA PLANILHA
    Sheets("Macro ME23N & ZMMSRR01").Range("B" & L).Select

    'ENTRAR COM ME23N
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NME23N"
    session.findById("wnd[0]").sendVKey 0

    'RODAR O PRIMEIRO WHILE PARA OS DADOS DA PLANILHA DO EXCEL' 'WHILE DA PLANILHA'
    Do While ActiveCell <> ""

        Sheets("Macro ME23N & ZMMSRR01").Range("B" & L).Select
        PEDIDO = Sheets("Macro ME23N & ZMMSRR01").Range("B" & L).Value
        NOTAFISCAL = Sheets("Macro ME23N & ZMMSRR01").Range("C" & L).Value

        On Error GoTo -1
        On Error GoTo 0
        On Error GoTo PROX_LINHA


        'INSERIR PEDIDO
        session.findById("wnd[0]/tbar[1]/btn[17]").press
        session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").Text = PEDIDO
        session.findById("wnd[1]/tbar[0]/btn[0]").press

        P = 0 'POSIÇÃO DO MATERIAL NA LISTA
        P_MATERIAL = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1," & P & "]").Text

        Do While (P_MATERIAL <> "") And (Sheets("Macro ME23N & ZMMSRR01").Range("E" & L).Value = "")

            VERIFICADA = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEGUI:1332/subSUB0:SAPLEINB:0300/tblSAPLEINBTC_0300/txtEKES-XBLNR[5,0]").Text

            If VERIFICADA <> "" Then
                RECEBIMENTO = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEGUI:1332/subSUB0:SAPLEINB:0300/tblSAPLEINBTC_0300/ctxtEKES-VBELN[7,0]").Text
            Else
                VERIFICADA = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEGUI:1332/subSUB0:SAPLEINB:0300/tblSAPLEINBTC_0300/txtEKES-XBLNR[5,1]").Text
                RECEBIMENTO = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEGUI:1332/subSUB0:SAPLEINB:0300/tblSAPLEINBTC_0300/ctxtEKES-VBELN[7,1]").Text
            If VERIFICADA = "" Then GoTo PROX_MATERIAL
            End If


            '---------- FUNÇÃO PARA REMOVER OS ZEROS E AJEITAR A NOTA FISCAL OBTIDA NO SAP ----------
            CARACT_VERIFICADA = Split(StrConv(VERIFICADA, vbUnicode), Chr$(0))
            ReDim Preserve CARACT_VERIFICADA(UBound(CARACT_VERIFICADA) - 1)

            'CONTAR NÚMERO DE ZEROS À ESQUERDA
            ZEROS = 0
            For Each C In CARACT_VERIFICADA
                If C = "0" Then
                    ZEROS = ZEROS + 1
                Else
                    Exit For
                End If
            Next C

            'REMOVENDO ZEROS NA ESQUERDA
            VERIFICADA = Right(VERIFICADA, Len(VERIFICADA) - ZEROS)

            d_VERIFICADA = Right(VERIFICADA, 2)
            d_VERIFICADA = Left(d_VERIFICADA, 1)
            If d_VERIFICADA = "-" Then
                'REMOVER NÚMERO DO ITEM NA NOTA FISCAL
                VERIFICADA = Left(VERIFICADA, Len(VERIFICADA) - 2)
            End If

            '------------ COMPARANDO NOTA FISCAIS E ALOCANDO VALORES NA PLANILHA ---------------

            If NOTAFISCAL = VERIFICADA Then

                'ALOCANDO RECEBIMENTO NA PLANILHA
                Sheets("Macro ME23N & ZMMSRR01").Range("E" & L).Select
                ActiveCell.Value = RECEBIMENTO

                'ALOCANDO TIPO NA PLANILHA ------------

                'SELECIONAR ABA DE CLASSIFICAÇÃO CONTÁBIL
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13").Select

                On Error GoTo -1
                On Error GoTo 0
                On Error GoTo CONTINUAR

                If session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-AUFNR").Text <> "" Then
                    Sheets("Macro ME23N & ZMMSRR01").Range("F" & L).Select
                    ActiveCell.Value = "COMPRA DIRETA"
                End If

                CONTINUAR:
                'VOLTAR PARA A ABA DE CONFIRMAÇÃO
                session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT17").Select

                Sheets("Macro ME23N & ZMMSRR01").Range("F" & L).Select
                If ActiveCell = "" Then
                    ActiveCell.Value = "ESTOQUE LIVRE"
                End If

            End If

            PROX_MATERIAL:
            'PULAR A LINHA NA LISTA DE MATERIAIS DO PEDIDO
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press

            P = P + 1

            If P = 15 Then
                GoTo PROX_LINHA
            End If

            P_MATERIAL = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1," & P & "]").Text

        Loop

        PROX_LINHA:
        'PULAR A LINHA DA PLANILHA
        L = L + 1
        Sheets("Macro ME23N & ZMMSRR01").Range("B" & L).Select

    Loop

    session.findById("wnd[0]/tbar[0]/btn[3]").press

    MsgBox ("Processo Concluído")

    End Sub

'##############################

'FUNÇÃO ZMMSRR01
Private Sub ZMMSRR01_Click()

    'DEFININDO AS VARIÁVEIS INICIAIS
    L = 5
    L_ANTERIOR = 4
    RR = Sheets("Macro ME23N & ZMMSRR01").Range("E" & L).Value
    PALETE = Sheets("Macro ME23N & ZMMSRR01").Range("D" & L).Text
    PALETE_ANTERIOR = Sheets("Macro ME23N & ZMMSRR01").Range("D" & L_ANTERIOR).Text
    LANCAMENTO = "01.06.2022"
    HOJE = Format(Date, "DD.MM.YYYY")

    'INCIALIZANDO O SAP
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

    '------- INICIAR TRANSAÇÃO ------------'

    session.findById("wnd[0]").maximize

    Sheets("Macro ME23N & ZMMSRR01").Range("E" & L).Select

    Do While ActiveCell <> ""

        'IMPRIMIR CAPA E VERIFICAR SE O LOCAL (PALETE) ATUAL É DIFERENTE DO ANTERIOR

        If PALETE <> PALETE_ANTERIOR Then
            Sheets("Capa de separação").Select
            Sheets("Capa de separação").Range("C8").Value = PALETE
            ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
            Sheets("Macro ME23N & ZMMSRR01").Select
            Sheets("Macro ME23N & ZMMSRR01").Range("E" & L).Select
        End If

        'IMPRIMIR RR
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NZMMSRR01"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtS_VBELN-LOW").Text = RR
        session.findById("wnd[0]/usr/ctxtP_BUDAT-LOW").Text = LANCAMENTO
        session.findById("wnd[0]/usr/ctxtP_BUDAT-HIGH").Text = HOJE
        session.findById("wnd[0]/usr/ctxtP_TDDEST").Text = "M105"
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        On Error GoTo -1
        On Error GoTo 0
        On Error GoTo PROX
        session.findById("wnd[1]/usr/btnBUTTON_1").press

        'PULAR A LINHA
        PROX:
        L = L + 1
        L_ANTERIOR = L_ANTERIOR + 1
        RR = Sheets("Macro ME23N & ZMMSRR01").Range("E" & L).Value
        PALETE = Sheets("Macro ME23N & ZMMSRR01").Range("D" & L).Text
        PALETE_ANTERIOR = Sheets("Macro ME23N & ZMMSRR01").Range("D" & L_ANTERIOR).Text

        Sheets("Macro ME23N & ZMMSRR01").Select
        Sheets("Macro ME23N & ZMMSRR01").Range("E" & L).Select

    Loop

    MsgBox ("Processo concluído!")

End Sub