Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)

Option Explicit

    'Declaração de variáveis globais e objetos globais.

    'Para nome de declarações deste tipo, utilizo o padrão:

    '   *Primeira letra maiuscula;
    '   *Nome composto separado por underline;
    
    '================================================
    'Instanciamento de Objetos:
    
    '===Arquivos .Xlsx===
    Public Planilha_Import As Workbook 'Arquivo atual da planilha de importação
    Public Planilha_Cont As Workbook 'Arquivo gerado na pasta do Sistema Contimatic
    
    '===Planilhas===
    Public Ws As Worksheet 'Muda de acordo com a necessidade de cada função
    'Ws será constantemente alterada pelas funções
    
    '===Buttons===
    'Cada botão tem o nome contido na forma acrecido de "btn"
    'Nem todo botão no padding necessariamene tem um váriavel declarada
    
    Public Multiplobtn As Shape
    Public Individualbtn As Shape
    Public Totalbtn As Shape
    Public Limparbtn As Shape
    Public Tutorialbtn As Shape 'Este é o único que foge da regra pois ficaria extenso demais
    Public Gerarbtn As Shape
    'Public Maisbtn As Shape 'Não é mais necessário
    'Public Menosbtn As Shape 'Não é mais necessário
    Public Admbtn As Shape
    Public TextBox As Shape
    '================================================
    'Declaração de variáveis:
    
    Public Senha As String 'Armazena a senha para desbloquear as planilas
    Public AdmKey As String 'Armazena a senha para liberar o modo Adm
    
    Public Numerador As Long 'Armazena o número do próximo lançamento
    
    Public Importar As Integer 'Armazena a forma de importação
    'INDIVIDUAL x MULTIPLO
    
    Public Total As Boolean 'Armazena a condição para realizar a automação completa
    Public TotalMsg As Boolean 'Armazena a condição para a necessidade do aviso de importação completa
    Public TutorialMsg As Boolean 'Armazena a condição para a necessidade do aviso de tutorial
    Public AdmState As Boolean 'Armazena o status de Adm
    '================================================
Sub Gerar()

    '================================================
    ' Esta é a Sub principal do programa, ela é o trigger para  todos os
    'ajustes da planilha financeira e define qual tipo de importação
    'deve ser realizada: INDIVIDUAL ou MULTIPLA e define se será
    'realizada a importação completamente automatizada ou manual.
    '================================================
    
    ' Setting de configurações visuais
    
    Gerarbtn.Fill.Transparency = 0.8
    ActiveWindow.Zoom = 90
    '================================================
    
    Dim linhas As Long 'Armazena o numero de linhas da planilha financeira
    linhas = Planilha_Import.Sheets("COLA").Cells(Rows.Count, 3).End(xlUp).Row
    
    Dim cadastrar As Boolean 'Armazena a condição de necessidade de cadastro
    
    If linhas > 1 Then 'linhas
    'É necessário que haja mais de uma linha na planilha para realizar a importação
    
        If Importar <> 0 Then 'Importar
        'É necessário que o usuário selecione o tipo de importação
        
            Desbloquear 'Function
            'Sleep (50)
            Limpar_Linhas 'Function
            Trocar 'Function
            Ajuste 'Function
            Rural 'Function
            Up 'Function
            'Inverter       'Function desativada por falta de uso
            Carregar 'Function
            Add 'Function
            cadastrar = Cadastro 'Function
            
            If cadastrar = False Then 'Cadastro
                ' Se não for necessário cadastrar fornecedores a planilha pode seguir para a importação completa
                
                If Total = True Then 'Total
                    'Se o usuário selecionar "IMPORTAÇÃO COMPLETA" a função executa
                    Importacao_Completa 'Function
                End If 'Total
                
            Else 'Cadastro
                'Se for necessário cadastrar fornecedores a importação não pode proceder e o usuário é alertado
                MsgBox ("Existem fornecedores para serem cadastrados! Favor verificar antes de prosseguir.")
                
            End If 'End Cadastro
            
        Else 'Importar
            'Caso o usuário não selecione o layout de importação nenhum função é ativada
            MsgBox "Selecione o Layout de Importação!"
        
        End If 'Impotar
    
    Else 'linhas
        'Menssagem em caso não haver informações na planilha
        MsgBox "Planilha vazia!"
    End If ' linhas
    
    
    '====================================================
    'Setting de configurações visuais, bloqueio e confirmação de sucesso
    
    Gerarbtn.Fill.Transparency = 0
    
    If AdmState = False Then
        Bloquear 'Function
    End If

    MsgBox ":)"

End Sub
Sub Adm()

    '====================================================
    'Sub Adm serve para habilitar o usuário como Administrador
    'Usuário Administrador tem acesso a alterar dados da planilha toda
    'Usuário sem Adm só pode alterar dados limitados da planilha
    '====================================================
    
    If AdmState = False Then ' State
        'Se o usuário ainda não for administrador a função pede a senha
        Dim resposta As String 'Armazena a senha fornecida pelo usuário
        resposta = InputBox("Insira a senha do Administrador")
        
        If Not IsEmpty(resposta) Then ' resposta
            'Se a senha fornecida não for vazia, segue:
            
            If resposta = AdmKey Then ' AdmKey
                'Se a senha estiver correta o status de Adm é habilitado e o button fica transparente
                AdmState = True
                Admbtn.Fill.Transparency = 0.8
                Desbloquear
                
            Else ' AdmKey
                'Se a senha estiver incorreta o usuário segue sem o status de Adm
                AdmState = False
                Admbtn.Fill.Transparency = 0
                Bloquear
                Exit Sub
            End If
        
        ElseIf IsEmpty(resposta) Then ' resposta
            'Se a resposta for vazia, o comportamento é o mesmo se a senha estivesse errada
            AdmState = False
            Admbtn.Fill.Transparency = 0
            Bloquear
            Exit Sub
        End If
    
    ElseIf AdmState = True Then ' State
        'Se o usuário ja for administrador ele perde o status
        AdmState = False
        Admbtn.Fill.Transparency = 0
        Bloquear
    End If
    
End Sub
Sub Numero()

    '====================================================
    'Sub Numero serve para o usuário alterar o valor do próximo lançamento
    '====================================================
    
    Dim resposta As String 'Armazena a resposta do usuário
    
    Set Ws = Planilha_Import.Sheets("COLA")
    'Instancia Ws na planilha COLA
    
    Do 'Loop para colher a resposta do usuáro
        resposta = InputBox("Insira o número do lançamento:")
         
        If resposta = "" Then
            'Se a resposta for vazia a função se encerra sem nenhuma alteração
            Exit Sub
        End If
        
        If resposta < 1 Then
            'Caso o usuário insira valor negativo ou nulo, o mesmo passa a valer '1'
            resposta = 1
        End If
        
        If IsNumeric(resposta) Then
            'Se o valor for numérico a função procede
            'Caso o usuário insira cadas decimais, essas são desconsideradas
            Numerador = Round(resposta, 0)
            Ws.Range("M2").Value = Numerador
            Desbloquear
            Ws.Columns(13).AutoFit
            Bloquear
            If AdmState = False Then
            End If
            Exit Do
        Else
            MsgBox "Insira um valor numérico!"
        End If
    Loop
    
End Sub

Sub Mais()
    '====================================================
    'Incrementa '1' ao valor do próximo lançamento
    '====================================================

    Numerador = Numerador + 1
    Planilha_Import.Sheets("COLA").Range("M2").Value = Numerador
    
End Sub

Sub Menos()

    '====================================================
    'Decrementa '1' ao valor do próximo lançamento
    '====================================================

    Numerador = Numerador - 1
    
    If Numerador < 1 Then
        'Caso o usuário decremente o valor além de '1', o mesmo passa a valer '1'
        Numerador = 1
    End If
    
    Planilha_Import.Sheets("COLA").Range("M2").Value = Numerador
    
End Sub

Sub Import_Total()

    Total = Not Total
    
    If Total = True Then
        Totalbtn.Fill.Transparency = 0.8
            
        If TotalMsg = False Then
            MsgBox "ATENÇÃO!!!" & vbCrLf & vbCrLf & "A importação completa é sujeita a erros de processo decorrentes de mudanças no sistema."
        End If
    End If
    
    TotalMsg = True
        
    If Total = False Then
        Totalbtn.Fill.Transparency = 0
    End If

End Sub

Sub Multiplo()

    Importar = 2
    
    Multiplobtn.Fill.Transparency = 0.8
    Individualbtn.Fill.Transparency = 0

End Sub

Sub Individual()

    Importar = 1
    
    Multiplobtn.Fill.Transparency = 0
    Individualbtn.Fill.Transparency = 0.8

End Sub

Public Sub Tutorial()
    
    ThisWorkbook.FollowHyperlink Address:="https://kpetrus7.github.io/tutorial-import/", NewWindow:=True
     
End Sub

Sub Clean()
    
    Limparbtn.Fill.Transparency = 0.8
    
    Desbloquear
    
    Set Ws = Planilha_Import.Sheets("COLA")
    Ws.Activate

    Dim i As Long
    i = (Ws.Cells(Rows.Count, 3).End(xlUp).Row) + 500
    
    Ws.Range("C2", "L" & i).ClearContents
    
    Set Ws = Planilha_Import.Sheets("CONCILIACAO")
    Ws.Range("A2", "K" & i).ClearContents
    
    
    Set Ws = Planilha_Import.Sheets("IMPORT")
    Ws.Range("C2", "i" & i).ClearContents
    Ws.Range("M2", "S" & i).ClearContents
    
    If AdmState = False Then
        Bloquear
    End If
        
    Sleep 500
    Limparbtn.Fill.Transparency = 0
    
End Sub

Function Carregar()

    Dim WsNext As Worksheet

    Dim lastRow As Long
    Dim lastCol As Long
    Dim rng As Range
    Dim cell As Range
    
    Set Ws = Planilha_Import.Sheets("COLA")
    Set WsNext = Planilha_Import.Sheets("CONCILIACAO")
    
    If Not IsEmpty(Ws.Range("C2").Value) Then
    
        lastRow = Ws.Cells(Ws.Rows.Count, 3).End(xlUp).Row
        lastCol = Ws.Cells(2, Ws.Columns.Count).End(xlToLeft).Column
    
        Ws.Range(Ws.Cells(2, 3), Ws.Cells(lastRow, lastCol)).Copy
        WsNext.Range("A2").PasteSpecial Paste:=xlPasteValues
    
        Set rng = WsNext.Range("A2:A" & lastRow)
        For Each cell In rng
            cell.Value = CDate(cell.Value)
        Next cell
    
        Application.CutCopyMode = False
    End If
    
End Function
Function Limpar_Linhas()

    Dim ultimaLinha As Long
    Dim i As Long
    
    Set Ws = Planilha_Import.Sheets("COLA")
    
    ultimaLinha = Ws.Cells(Rows.Count, 3).End(xlUp).Row + 2
    'Debug.Print ultimaLinha
    
    For i = ultimaLinha To 2 Step -1
        
        If IsEmpty(Cells(i, 6)) Or Not IsNumeric(Cells(i, 6)) Or IsEmpty(Cells(i, 3)) Then
            Ws.Rows(i).Delete
        End If
    Next i
    
End Function

Function Trocar()

    Dim G As Range
    Dim H As Range
    Dim i As Range
    
    Set G = Sheets("COLA").Range("G1")
    Set H = Sheets("COLA").Range("H1")
    Set i = Sheets("COLA").Range("I1")
    
    If G.Value <> "Acrés" And H.Value = "Acrés" Then
        Trocar_G_H
    End If
            
    If G.Value <> "Acrés" And i.Value = "Acrés" Then
        Trocar_G_I
    End If
    
    If H.Value <> "Juros" And i.Value = "Juros" Then
        Trocar_H_I
    End If

End Function
Function Ajuste()

    Dim i As Long
    Dim cell As Range
    
    Set Ws = Planilha_Import.Sheets("COLA")
    
    i = Ws.Cells(Ws.Rows.Count, 3).End(xlUp).Row
    
    For Each cell In Ws.Range("E1:L" & i)
        If IsNumeric(cell.Value) Then
            cell.Value = CDec(cell.Value)
        End If
    Next cell

    Ws.Cells.EntireColumn.AutoFit
    Ws.Columns("C:C").NumberFormat = "m/d/yyyy"
    Ws.Columns("F:J").Style = "Currency"
    
End Function

Function Rural()

    Dim Funrural()
    
    Dim a As Long   'Contador loop For para matriz funrural x descrição
    Dim b As Long   'Contador loop For para matriz funrural x descrição
    Dim i As Long   'Número de elementos matriz funrural
    Dim n As Long   'Contador função For que mapeia tabela funrural
    Dim r As Long   'Número de linhas tabela de importação
    Dim s As Long   'Contador função For que atribui valor aos elementos da matriz Funrural
    
    Dim temp As Double 'Armazena o valor para ser substituido no "Val.Doc"
    
    i = Sheets("CADASTRO").Cells(Rows.Count, "L").End(xlUp).Row

    ReDim Funrural(1 To i)
    
    For n = 3 To i
        Funrural(n) = Sheets("CADASTRO").Range("L" & n).Value
    Next n
    
    r = Sheets("COLA").Cells(Rows.Count, "D").End(xlUp).Row
    
    Dim Fornecedor()
    ReDim Fornecedor(1 To r)
    
    For s = 2 To r
        Fornecedor(s) = Sheets("COLA").Range("D" & s).Value
    Next s
    
    For a = 2 To r
        For b = 3 To i

                If Fornecedor(a) = Funrural(b) Then

                    Range("I" & a).Value = 0
                    temp = Range("J" & a).Value
                    Range("F" & a).Value = temp
                    temp = 0
                    
            End If
        Next b
    Next a
     
End Function
Function Up()
    
    Dim i As Long
    i = Cells(Rows.Count, 3).End(xlUp).Row - 1
    
     If i >= 2 Then
    
        With Range("C2:L" & i + 1)
            .Sort Key1:=Range("I2"), Order1:=xlDescending, _
                  Key2:=Range("H2"), Order2:=xlDescending, _
                  Key3:=Range("G2"), Order3:=xlDescending, Header:=xlNo
        End With
    End If
                   
End Function

Function Trocar_H_I()

    Dim ultimaLinha As Long
    Dim i As Long
    Dim temp As Variant

    ultimaLinha = Cells(Rows.Count, "i").End(xlUp).Row

    For i = 1 To ultimaLinha
       
        temp = Cells(i, "i").Value
 
        Cells(i, "i").Value = Cells(i, "H").Value
        
        Cells(i, "H").Value = temp
    Next i
    
End Function
Function Trocar_G_H()

    Dim ultimaLinha As Long
    Dim i As Long
    Dim temp As Variant

    ultimaLinha = Cells(Rows.Count, "G").End(xlUp).Row

    For i = 1 To ultimaLinha
       
        temp = Cells(i, "G").Value
 
        Cells(i, "G").Value = Cells(i, "H").Value

        Cells(i, "H").Value = temp
    Next i
    
End Function
Function Trocar_G_I()

    Dim ultimaLinha As Long
    Dim i As Long
    Dim temp As Variant

    ultimaLinha = Cells(Rows.Count, "G").End(xlUp).Row

    For i = 1 To ultimaLinha
       
        temp = Cells(i, "G").Value
 
        Cells(i, "G").Value = Cells(i, "i").Value
        
        Cells(i, "i").Value = temp
    Next i
    
End Function

Function Add()
    
    Set Ws = Planilha_Import.Sheets("IMPORT")
    
    Dim rng As Range
    Dim cell As Range
    
    Dim i As Long
    
    Dim T As String
    
    i = (Planilha_Import.Sheets("CONCILIACAO").Cells(Rows.Count, "A").End(xlUp).Row)
    
    If Importar = 2 Then 'Multiplo
    
        Ws.Range("C2").Value = Numerador
        Ws.Range("D2").Formula = "=MAX(CONCILIACAO!A:A)"
        Ws.Range("E2").Formula = "=VLOOKUP(CONCILIACAO!B2,FORNECEDORES,2,0)"
        Ws.Range("F2").Value = "T"
        Ws.Range("G2").Value = "=CONCILIACAO!D2"
        Ws.Range("H2").Formula = "=VLOOKUP(CONCILIACAO!I2,PAGTO,3,0)"
        Ws.Range("I2").Formula = "=VLOOKUP(CONCILIACAO!I2,PAGTO,4,0)&IF(CONCILIACAO!C2=0,""000001"",IF(CONCILIACAO!C2<10,""00000""&CONCILIACAO!C2,IF(CONCILIACAO!C2<100,""0000""&CONCILIACAO!C2,IF(CONCILIACAO!C2<1000,""000""&CONCILIACAO!C2,IF(CONCILIACAO!C2<10000,""00""&CONCILIACAO!C2,IF(CONCILIACAO!C2<100000,""0""&CONCILIACAO!C2,CONCILIACAO!C2))))))&VLOOKUP(CONCILIACAO!I2,PAGTO,5,0)&IF(OR(E2=20959,E2=383),"" ""&CONCILIACAO!B2,"""")"
        
        Ws.Range("M2:M4").Value = Numerador
        Ws.Range("N2:N4").Formula = "=MAX(CONCILIACAO!A:A)"
        Ws.Range("O2:P2").Value = "M"
        Ws.Range("Q2").Formula = "=SUM(G:G)+SUM(CONCILIACAO!F:F)+SUM(CONCILIACAO!E:E)"
        
        Ws.Range("O3").Value = "T"
        Ws.Range("P3").Formula = "=VLOOKUP(CONCILIACAO!I2,PAGTO,2,0)"
        Ws.Range("Q3").Formula = "=Q2-SUM(CONCILIACAO!G:G)"
        Ws.Range("R3").Formula = "=VLOOKUP(CONCILIACAO!I2,PAGTO,6,0)"
        
        Ws.Range("O4").Formula = "=IF(AND(CONCILIACAO!F2=0,CONCILIACAO!G2=0,CONCILIACAO!E2>0),8000,IF(AND(CONCILIACAO!G2=0,CONCILIACAO!E2=0,CONCILIACAO!F2>0),8040,""T""))"
        Ws.Range("P4").Formula = "=IF(AND(CONCILIACAO!E2=0,CONCILIACAO!F2=0,CONCILIACAO!G2>0),3310,""T"")"
        Ws.Range("Q4").Formula = "=SUM(CONCILIACAO!E2:G2)"
        Ws.Range("R4").Formula = "=IF(AND(CONCILIACAO!F2=0,CONCILIACAO!E2=0,CONCILIACAO!G2>0),628,IF(AND(CONCILIACAO!F2=0,CONCILIACAO!G2=0,CONCILIACAO!E2>0),98,IF(AND(CONCILIACAO!E2=0,CONCILIACAO!G2=0,CONCILIACAO!F2>0),627,"""")))"
        Ws.Range("S4").Formula = "=IF(AND(CONCILIACAO!E2=0,CONCILIACAO!F2=0,CONCILIACAO!G2>0),CONCILIACAO!C2&"" LOJA ""&CONCILIACAO!J2,IF(AND(CONCILIACAO!G2=0,CONCILIACAO!E2=0,CONCILIACAO!F2>0),CONCILIACAO!C2,""""))"
        
        If i > 2 Then
            Set rng = Ws.Range("C2:" & "i" & i)
            rng.FillDown

            Set rng = Ws.Range("M4:" & "S" & i + 2)
            rng.FillDown
        End If
        
        For i = i + 5 To 3 Step -1
            If Ws.Range("Q" & i).Value = 0 Then
                Ws.Range("M" & i & ":S" & i).ClearContents
           End If
        Next i
        
    End If
    
    If Importar = 1 Then 'Individual
    
        Ws.Range("M2:S50000").ClearContents
    
        Ws.Range("C2").Formula = "=ROW(C2) - 2 + CONCILIACAO!$K$2"
        Ws.Range("D2").Value = "=CONCILIACAO!A2"
        Ws.Range("E2").Formula = "=VLOOKUP(CONCILIACAO!B2,FORNECEDORES,2,0)"
        Ws.Range("F2").Formula = "=VLOOKUP(CONCILIACAO!I2,PAGTO,2,0)"
        Ws.Range("G2").Value = "=CONCILIACAO!D2"
        Ws.Range("H2").Formula = "=VLOOKUP(CONCILIACAO!I2,PAGTO,3,0)"
        Ws.Range("I2").Formula = "=VLOOKUP(CONCILIACAO!I2,PAGTO,4,0)&IF(CONCILIACAO!C2=0,""000001"",IF(CONCILIACAO!C2<10,""00000""&CONCILIACAO!C2,IF(CONCILIACAO!C2<100,""0000""&CONCILIACAO!C2,IF(CONCILIACAO!C2<1000,""000""&CONCILIACAO!C2,IF(CONCILIACAO!C2<10000,""00""&CONCILIACAO!C2,IF(CONCILIACAO!C2<100000,""0""&CONCILIACAO!C2,CONCILIACAO!C2))))))&VLOOKUP(CONCILIACAO!I2,PAGTO,5,0)&IF(OR(E2=20959,E2=383),"" ""&CONCILIACAO!B2,"""")"
        
        If i > 2 Then
           Set rng = Ws.Range("C2:" & "i" & i)
            rng.FillDown
        End If
        
        MsgBox ("Atenção!!!" & vbCrLf & vbCrLf & "Verifique a existência de eventos como Juros, Descontos ou Tarifas")
        
    End If
    
    Dim ref As Long
    ref = Planilha_Import.Sheets("IMPORT").Cells(Rows.Count, "C").End(xlUp).Row
    
    Numerador = (Planilha_Import.Sheets("IMPORT").Range("C" & ref).Value) + 1
    
    Planilha_Import.Sheets("COLA").Range("M2").Value = Numerador
    
    Planilha_Import.Sheets("IMPORT").Activate
    Planilha_Import.Sheets("IMPORT").Range("A1").Select
    ActiveWindow.Zoom = 100
    
    
End Function

Function Inverter()

    Set Ws = Planilha_Import.Sheets("COLA")

    Dim k As Double
    Dim j As Long
    Dim i As Long
    i = Ws.Cells(Ws.Rows.Count, 3).End(xlUp).Row
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Inverterer valores?", vbYesNo + vbQuestion, "Confirmação")
    
    If resposta = vbYes Then
    
        For j = 2 To i
        
            k = Ws.Cells(j, "F").Value
            Ws.Cells(j, "F").Value = -k
            
            k = Ws.Cells(j, "J").Value
            Ws.Cells(j, "J").Value = -k
             
        Next j
    End If
        
End Function

Function Cadastro()
    
    Dim i As Long
    i = Planilha_Import.Sheets("IMPORT").Cells(Rows.Count, "E").End(xlUp).Row
    
    Dim cell As Range
    Dim rng As Range
    Set rng = Planilha_Import.Sheets("IMPORT").Range("E2:E" & i)
    
    Cadastro = False
    
    For Each cell In rng
        
        Desbloquear
        
        If IsError(cell.Value) Then
            
            Dim k As Long
            Dim j As Long
            j = cell.Row
            
            Planilha_Import.Sheets("CONCILIACAO").Range("B" & j).Copy
            
            k = (Planilha_Import.Sheets("CADASTRO").Cells(Rows.Count, "B").End(xlUp).Row) + 1
            
            Planilha_Import.Sheets("CADASTRO").Range("B" & k).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            
            Planilha_Import.Sheets("CADASTRO").Unprotect Password:=Senha
            Planilha_Import.Sheets("CADASTRO").Range("C" & k).Locked = False
            Planilha_Import.Sheets("CADASTRO").Protect Password:=Senha
            Planilha_Import.Sheets("CADASTRO").Activate
            Planilha_Import.Sheets("CADASTRO").Range("C" & k).Select
            
            Cadastro = True
            
            Menos 'Sub
            
        End If
    
    Next cell
    
End Function

Function Importacao_Completa()

    Dim caminho As String
    caminho = Planilha_Import.Sheets("CADASTRO").Range("S3").Value
    
    Dim arquivo As String
    arquivo = "3_2025_Lctos.xlsx"
    
    If Dir(caminho) = "" Then 'DIR
    
        Set Planilha_Cont = Workbooks.Add
        Planilha_Cont.SaveAs Filename:=caminho
        
    Else 'DIR
    
        Dim WbTeste As Workbook
        
        On Error Resume Next
        Set WbTeste = Workbooks(arquivo)
        On Error GoTo 0
        
        If WbTeste Is Nothing Then
            Set Planilha_Cont = Workbooks.Open(Filename:=caminho)
        End If
        
    End If 'DIR
        
    Planilha_Cont.Worksheets(1).Range("A1").Value = "Lançamento"
    Planilha_Cont.Worksheets(1).Range("B1").Value = "Data"
    Planilha_Cont.Worksheets(1).Range("C1").Value = "Débito"
    Planilha_Cont.Worksheets(1).Range("D1").Value = "Crédito"
    Planilha_Cont.Worksheets(1).Range("E1").Value = "Valor"
    Planilha_Cont.Worksheets(1).Range("F1").Value = "Histórico Padrão"
    Planilha_Cont.Worksheets(1).Range("G1").Value = "Complemento"
    Planilha_Cont.Worksheets(1).Range("H1").Value = "CCDB"
    Planilha_Cont.Worksheets(1).Range("I1").Value = "CCCR"
    Planilha_Cont.Worksheets(1).Range("J1").Value = "CNPJ"
    
    Delete_Ws Planilha_Cont
    Transferir_Valores Planilha_Import, Planilha_Cont
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Limpar planilha?", vbYesNo + vbQuestion, "Confirmação")

    If resposta = vbYes Then
        Clean
    End If
    
End Function

Function Delete_Ws(Wb As Workbook)
    
    Dim Nome As Variant
    Dim WsDelete As Variant
    WsDelete = Array("1 ATIVO", "2 PASSIVO", "3 RECEITA", "4 DESPESA", "5 CUSTO", "6 RESULTADO")
    
    Dim WsTeste As Worksheet

    For Each Nome In WsDelete
        On Error Resume Next
        
        Set WsTeste = Wb.Worksheets(Nome)
        
        If Not WsTeste Is Nothing Then
            Application.DisplayAlerts = False
            WsTeste.Delete
            Application.DisplayAlerts = True
        End If
        
        Set WsTeste = Nothing
        On Error GoTo 0
    Next Nome
    
End Function

Function Transferir_Valores(origem As Workbook, Destino As Workbook)
    On Error Resume Next
    
    Dim Credito As Long
    Dim Debito As Long
    Dim contabil As Long
    
    Credito = origem.Worksheets(3).Cells(Rows.Count, 13).End(xlUp).Row
    Debito = origem.Worksheets(3).Cells(Rows.Count, 3).End(xlUp).Row
    contabil = Destino.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row

    If Importar = 2 Then
        origem.Worksheets(3).Range("M2:S" & Credito).Copy
        Destino.Worksheets(1).Range("A" & contabil + 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
    End If

    contabil = Destino.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    origem.Worksheets(3).Range("C2:i" & Debito).Copy
    Destino.Worksheets(1).Range("A" & contabil + 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    On Error GoTo 0
    
End Function

Function PreSet()

    Senha = "#a3V$1$hm_a3V$1$hm#"
    AdmKey = "123"

End Function

Function Bloquear()
    
    Set Ws = ActiveSheet
    
    Planilha_Import.Sheets("COLA").Protect Password:=Senha
    Planilha_Import.Sheets("CONCILIACAO").Protect Password:=Senha
    Planilha_Import.Sheets("IMPORT").Protect Password:=Senha
    Planilha_Import.Sheets("CADASTRO").Protect Password:=Senha
    
    Ws.Activate

End Function

Function Desbloquear()

    Set Ws = ActiveSheet
        
    Planilha_Import.Sheets("COLA").Unprotect Password:=Senha
    Planilha_Import.Sheets("CONCILIACAO").Unprotect Password:=Senha
    Planilha_Import.Sheets("IMPORT").Unprotect Password:=Senha
    Planilha_Import.Sheets("CADASTRO").Unprotect Password:=Senha
    
    Ws.Activate

End Function
