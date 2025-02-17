VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TrabPrec 
   Caption         =   "PreCalc"
   ClientHeight    =   8475.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14700
   OleObjectBlob   =   "TrabPrec.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "TrabPrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    PreferenciaBox.AddItem "Preferência Até 2020"
    PreferenciaBox.AddItem "Preferência SINDIFAZ"
    PreferenciaBox.AddItem "Preferência ORÇ.2021"
    PreferenciaBox.AddItem "Preferência ORÇ.2022"
    PreferenciaBox.AddItem "Preferência ORÇ.2023"
    PreferenciaBox.AddItem "Preferência ORÇ.2024"
    AcordoBox.AddItem "Acordo Até 2020"
    AcordoBox.AddItem "Acordo SINDIFAZ"
    AcordoBox.AddItem "Acordo Só Laguz"
    AcordoBox.AddItem "Acordo ORÇ.2021"
    AcordoBox.AddItem "Acordo ORÇ.2022"
    AcordoBox.AddItem "Acordo ORÇ.2023"
    AcordoBox.AddItem "Acordo ORÇ.2024"
    SemModeloBox.AddItem "CRIAR MODELO"
    OrdemBox.AddItem "Ordem Até 2020"
    OrdemBox.AddItem "Ordem ORÇ.2021"
    OrdemBox.AddItem "Ordem ORÇ.2022"
    OrdemBox.AddItem "Ordem ORÇ.2023"
    OrdemBox.AddItem "Ordem ORÇ.2024"
    PreferenciaBox1.AddItem "Preferência ORÇ.2022"
    PreferenciaBox1.AddItem "Preferência ORÇ.2023"
    PreferenciaBox1.AddItem "Preferência ORÇ.2024"
End Sub
Private Sub IrParaPlanilha_Click()
    Dim caminhoArquivo As String
    Dim pastaBase As String

    ' Obtém o caminho da pasta onde a planilha principal está localizada
    pastaBase = ThisWorkbook.Path & "\planilhas\"
    
    caminhoArquivo = ""
    

    ' Verifica a seleção de cada ComboBox e atribui o caminho do arquivo correspondente
    If PreferenciaBox.Value <> "" Then
        Select Case PreferenciaBox.Value
            Case "Preferência Até 2020"
                caminhoArquivo = pastaBase & "Preferencia2020.xlsm"
            Case "Preferência SINDIFAZ"
                caminhoArquivo = pastaBase & "PreferenciaSindifaz.xlsm"
            Case "Preferência ORÇ.2021"
                caminhoArquivo = pastaBase & "Preferencia2021.xlsm"
            Case "Preferência ORÇ.2022"
                caminhoArquivo = pastaBase & "Preferencia2022.xlsm"
            Case "Preferência ORÇ.2023"
                caminhoArquivo = pastaBase & "Preferencia2023.xlsm"
            Case "Preferência ORÇ.2024"
                caminhoArquivo = pastaBase & "Preferencia2024.xlsm"
        End Select
    End If

    ' Verificando a AcordoBox
    If AcordoBox.Value <> "" Then
        Select Case AcordoBox.Value
            Case "Acordo Até 2020"
                caminhoArquivo = pastaBase & "Acordo2023.xlsm"
            Case "Acordo SINDIFAZ"
                caminhoArquivo = pastaBase & "AcordoSindifaz.xlsm"
            Case "Acordo Só Laguz"
                caminhoArquivo = pastaBase & "AcordoSoLaguz.xlsm"
            Case "Acordo ORÇ.2021"
                caminhoArquivo = pastaBase & "AcordoOrc2021.xlsm"
            Case "Acordo ORÇ.2022"
                caminhoArquivo = pastaBase & "AcordoOrc2022.xlsm"
            Case "Acordo ORÇ.2023"
                caminhoArquivo = pastaBase & "AcordoOrc2023.xlsm"
            Case "Acordo ORÇ.2024"
                caminhoArquivo = pastaBase & "AcordoOrc2024.xlsm"
        End Select
    End If

    ' Verificando a OrdemBox
    If OrdemBox.Value <> "" Then
        Select Case OrdemBox.Value
            Case "Ordem Até 2020"
                caminhoArquivo = pastaBase & "Ordem2020.xlsm"
            Case "Ordem ORÇ.2021"
                caminhoArquivo = pastaBase & "Ordem2021.xlsm"
            Case "Ordem ORÇ.2022"
                caminhoArquivo = pastaBase & "Ordem2022.xlsm"
            Case "Ordem ORÇ.2023"
                caminhoArquivo = pastaBase & "Ordem2023.xlsm"
            Case "Ordem ORÇ.2024"
                caminhoArquivo = pastaBase & "Ordem2024.xlsm"
        End Select
    End If

    ' Verificando a PreferenciaBox1
    If PreferenciaBox1.Value <> "" Then
        Select Case PreferenciaBox1.Value
            Case "Preferência ORÇ.2022"
                caminhoArquivo = pastaBase & "Prefe2022.xlsm"
            Case "Preferência ORÇ.2023"
                caminhoArquivo = pastaBase & "Prefe2023.xlsm"
            Case "Preferência ORÇ.2024"
                caminhoArquivo = pastaBase & "Prefe2024.xlsm"
        End Select
    End If

    ' Se um arquivo foi selecionado, abrir o arquivo
    If caminhoArquivo <> "" Then
        ' Verifica se o arquivo existe antes de tentar abri-lo
        If Dir(caminhoArquivo) <> "" Then
            Workbooks.Open caminhoArquivo
        Else
            MsgBox "O arquivo não foi encontrado: " & caminhoArquivo, vbExclamation
        End If
    Else
        MsgBox "Por favor, selecione uma opção para acessar a planilha!", vbExclamation
    End If
End Sub

Private Sub NPrecatorio_Change()
    Dim numero As String
    numero = VBA.Replace(Me.NPrecatorio.Text, "-", "")
    numero = VBA.Replace(numero, ".", "")
    
    ' Garante que são apenas números e limita a 20 caracteres (sem os separadores)
    numero = Left(numero, 21)
    
    If Len(numero) >= 21 Then
        Me.NPrecatorio.Text = Left(numero, 7) & "-" & Mid(numero, 8, 2) & "." & Mid(numero, 10, 4) & "." & Mid(numero, 14, 1) & "." & Mid(numero, 15, 2) & "." & Mid(numero, 17, 4)
        Me.NPrecatorio.SelStart = Len(Me.NPrecatorio.Text) ' Mantém o cursor no final
    End If
End Sub
Private Sub NPrecatorio_AfterUpdate()
    Dim wsBanco As Worksheet
    Dim wsMenu As Worksheet
    Dim rng As Range
    Dim cel As Range
    Dim linha As Long
    Dim precatorio As String
    Dim protocolo As String
    Dim Exequente As String
    Dim Executado As String
    Dim Orcamento As String
    Dim DataHomologado As String
    
    ' Definir as planilhas
    Set wsBanco = ThisWorkbook.Sheets("BANCO")
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    
    ' Capturar o número digitado no TextBox
    precatorio = Me.NPrecatorio.Text
    
    ' Buscar a célula correspondente ao número digitado (Coluna E da planilha BANCO)
    Set rng = wsBanco.Range("E1:E7745")
    Set cel = rng.Find(what:=precatorio, lookat:=xlWhole)
    
    ' Se encontrou, preenche os campos e a planilha Menu
    If Not cel Is Nothing Then
        linha = cel.Row
        protocolo = wsBanco.Cells(linha, 1).Value
        Exequente = wsBanco.Cells(linha, 3).Value
        Executado = wsBanco.Cells(linha, 4).Value
        Orcamento = wsBanco.Cells(linha, 2).Value
        DataHomologado = wsBanco.Cells(linha, 6).Value
        
        ' Preencher os TextBox do formulário
        Me.DataProtocolo.Text = protocolo
        Me.Exequente.Text = Exequente
        Me.Executado.Text = Executado
        Me.Orcamento.Text = Orcamento
        Me.DataHomologado.Text = DataHomologado
        
        ' Atualizar as células na planilha Menu
        wsMenu.Range("NumeroPrecatorio").Value = precatorio
        wsMenu.Range("DataProtocolo").Value = protocolo
        wsMenu.Range("NExequente").Value = Exequente
        wsMenu.Range("NExecutado").Value = Executado
        wsMenu.Range("Porcamento").Value = Orcamento
        wsMenu.Range("DataHomologado").Value = DataHomologado
    Else
        ' Mensagens padrão se não encontrar o precatório
        Me.DataProtocolo.Text = "Digite a Data do Protocolo"
        Me.Exequente.Text = "Digite Nome do Exequente"
        Me.Executado.Text = "Digite Nome do Executado"
        Me.Orcamento.Text = "Digite o orçamento"
        Me.DataHomologado.Text = "Digite a Data do Cálculo Homologado"
        
        ' Limpar as células da planilha Menu
        wsMenu.Range("NumeroPrecatorio").ClearContents
        wsMenu.Range("DataProtocolo").ClearContents
        wsMenu.Range("NExequente").ClearContents
        wsMenu.Range("NExecutado").ClearContents
        wsMenu.Range("Porcamento").ClearContents
        wsMenu.Range("DataHomologado").ClearContents
    End If
End Sub
Private Sub DataProtocolo_Change()
    ThisWorkbook.Sheets("Menu").Range("DataProtocolo").Value = Me.DataProtocolo.Text
End Sub
Private Sub Exequente_Change()
    ThisWorkbook.Sheets("Menu").Range("NExequente").Value = Me.Exequente.Text
End Sub
Private Sub Executado_Change()
    ThisWorkbook.Sheets("Menu").Range("NExecutado").Value = Me.Executado.Text
End Sub
Private Sub Orcamento_Change()
    ThisWorkbook.Sheets("Menu").Range("Porcamento").Value = Me.Orcamento.Text
End Sub
Private Sub DataHomologado_Change()
    ThisWorkbook.Sheets("Menu").Range("DataHomologado").Value = Me.DataHomologado.Text
End Sub
Private Sub PrincipalHomologado_Change()
    ThisWorkbook.Sheets("Menu").Range("PrincipalHomologado").Value = Me.PrincipalHomologado.Text
    AtualizarTotalHomologado
End Sub
Private Sub JurosHomologado_Change()
    ThisWorkbook.Sheets("Menu").Range("JurosHomologado").Value = Me.JurosHomologado.Text
    AtualizarTotalHomologado
End Sub
Private Sub AtualizarTotalHomologado()
    Dim principal As Double
    Dim juros As Double
    
    ' Verifica se os valores digitados são números válidos
    If IsNumeric(Me.PrincipalHomologado.Text) Then
        principal = CDbl(Me.PrincipalHomologado.Text)
    Else
        principal = 0
    End If
    
    If IsNumeric(Me.JurosHomologado.Text) Then
        juros = CDbl(Me.JurosHomologado.Text)
    Else
        juros = 0
    End If
    
    ' Calcula o total e preenche o TextBox
    Me.TotalHomologado.Text = principal + juros
End Sub


