Private Sub Workbook_Open()
    
    '====================================================
    Set Planilha_Import = ThisWorkbook
    Set Ws = Planilha_Import.Sheets("COLA")
    '====================================================


    '====================================================
    Set Multiplobtn = Ws.Shapes("MULTIPLO")
    Set Individualbtn = Ws.Shapes("INDIVIDUAL")
    Set Totalbtn = Ws.Shapes("TOTAL")
    Set Limparbtn = Ws.Shapes("LIMPAR")
    Set Tutorialbtn = Ws.Shapes("TUTORIAL")
    Set Gerarbtn = Ws.Shapes("GERAR")
    Set Admbtn = Ws.Shapes("ADMINISTRADOR")
    Set TextBox = Ws.Shapes("TEXT_BOX")
    '====================================================


    '====================================================
    Multiplobtn.Fill.Transparency = 0
    Individualbtn.Fill.Transparency = 0
    Totalbtn.Fill.Transparency = 0
    Limparbtn.Fill.Transparency = 0
    Gerarbtn.Fill.Transparency = 0
    Tutorialbtn.Fill.Transparency = 0
    Admbtn.Fill.Transparency = 0
    TextBox.TextFrame.Characters.Text = ""
    '====================================================

    Numerador = 1
    
    '====================================================
    Ws.Range("C1").Value = "Dt.Venc"
    Ws.Range("D1").Value = "Descricao"
    Ws.Range("E1").Value = "Núm NF"
    Ws.Range("F1").Value = "Val.Doc"
    Ws.Range("G1").Value = "Acrés"
    Ws.Range("H1").Value = "Juros"
    Ws.Range("I1").Value = "Desc"
    Ws.Range("J1").Value = "V.Liquido"
    Ws.Range("K1").Value = "C/C Pagto"
    Ws.Range("L1").Value = "Lj"
    Ws.Range("M1").Value = "Nº"
    Ws.Range("M2").Value = Numerador
    '====================================================

    Importar = 0
    
    '====================================================
    Total = False
    TotalMsg = False
    AdmState = False
    '====================================================
    
    
    '====================================================
    'Esse trecho é utilizado para exibir a menssagem de tutorial e deve ser
    'sugerido apenas quando o usuário abre a planilha. Utilizar o botão reset
    'não fará a menssagem ser mostrada novamente.
    
    If TutorialMsg = False Then 'Por padrão, variaveis booleanas iniciam em false
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("Bem-Vindo! Gostaria de ver o tutorial?", vbYesNo + vbQuestion, "Confirmação")

        If resposta = vbYes Then
            Tutorial
        End If
    TutorialMsg = True 'Após esse trehco, a variável permanece em True
    End If
    '====================================================
    
    
    '====================================================
    PreSet 'Function
    Clean 'Function
    Bloquear 'Function
    '====================================================
    
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    '====================================================
    'Essa Sub é utilizada para monitorar a planilha de cadastro e bloquear
    'novamente após o usuário sem permissão de administrador cadastrar
    'novos fornecedores.
    '====================================================
    If AdmState = False Then
        'MsgBox "Celula: " & Target.Address & " planilha: " & Sh.Name
        If Sh.Name = "CADASTRO" Then
            Desbloquear
            Planilha_Import.Sheets("CADASTRO").Range(Target.Address).Locked = True
            Bloquear
            
        End If
    End If
End Sub

Sub Reset()
    '====================================================
    'Apenas chama a função de abertura do arquivo, também serve como
    'reset da ferramenta, redefinindo os padrões e concertando possíveis
    'erros de uso.
    '====================================================
    Dim resposta As VbMsgBoxResult
    Dim texto As String

    resposta = MsgBox("Manter Anotações?", vbYesNo + vbQuestion, "Confirmação")
    If resposta = vbYes Then
        Set Planilha_Import = ThisWorkbook
        Set Ws = Planilha_Import.Sheets("COLA")
        Set TextBox = Ws.Shapes("TEXT_BOX")
        texto = TextBox.TextFrame.Characters.Text
    Else
        texto = ""
    End If
    
    Call Workbook_Open
    
    TextBox.TextFrame.Characters.Text = texto
End Sub
