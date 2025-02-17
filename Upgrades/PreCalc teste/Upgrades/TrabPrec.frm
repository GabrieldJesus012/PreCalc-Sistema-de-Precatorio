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
    PreferenciaBox.AddItem "Prefer�ncia At� 2020"
    PreferenciaBox.AddItem "Prefer�ncia SINDIFAZ"
    PreferenciaBox.AddItem "Prefer�ncia OR�.2021"
    PreferenciaBox.AddItem "Prefer�ncia OR�.2022"
    PreferenciaBox.AddItem "Prefer�ncia OR�.2023"
    PreferenciaBox.AddItem "Prefer�ncia OR�.2024"
    AcordoBox.AddItem "Acordo At� 2020"
    AcordoBox.AddItem "Acordo SINDIFAZ"
    AcordoBox.AddItem "Acordo S� Laguz"
    AcordoBox.AddItem "Acordo OR�.2021"
    AcordoBox.AddItem "Acordo OR�.2022"
    AcordoBox.AddItem "Acordo OR�.2023"
    AcordoBox.AddItem "Acordo OR�.2024"
    SemModeloBox.AddItem "CRIAR MODELO"
    OrdemBox.AddItem "Ordem At� 2020"
    OrdemBox.AddItem "Ordem OR�.2021"
    OrdemBox.AddItem "Ordem OR�.2022"
    OrdemBox.AddItem "Ordem OR�.2023"
    OrdemBox.AddItem "Ordem OR�.2024"
    PreferenciaBox1.AddItem "Prefer�ncia OR�.2022"
    PreferenciaBox1.AddItem "Prefer�ncia OR�.2023"
    PreferenciaBox1.AddItem "Prefer�ncia OR�.2024"
End Sub
Private Sub IrParaPlanilha_Click()
    Dim caminhoArquivo As String
    Dim pastaBase As String

    ' Obt�m o caminho da pasta onde a planilha principal est� localizada
    pastaBase = ThisWorkbook.Path & "\planilhas\"
    
    caminhoArquivo = ""
    

    ' Verifica a sele��o de cada ComboBox e atribui o caminho do arquivo correspondente
    If PreferenciaBox.Value <> "" Then
        Select Case PreferenciaBox.Value
            Case "Prefer�ncia At� 2020"
                caminhoArquivo = pastaBase & "Preferencia2020.xlsm"
            Case "Prefer�ncia SINDIFAZ"
                caminhoArquivo = pastaBase & "PreferenciaSindifaz.xlsm"
            Case "Prefer�ncia OR�.2021"
                caminhoArquivo = pastaBase & "Preferencia2021.xlsm"
            Case "Prefer�ncia OR�.2022"
                caminhoArquivo = pastaBase & "Preferencia2022.xlsm"
            Case "Prefer�ncia OR�.2023"
                caminhoArquivo = pastaBase & "Preferencia2023.xlsm"
            Case "Prefer�ncia OR�.2024"
                caminhoArquivo = pastaBase & "Preferencia2024.xlsm"
        End Select
    End If

    ' Verificando a AcordoBox
    If AcordoBox.Value <> "" Then
        Select Case AcordoBox.Value
            Case "Acordo At� 2020"
                caminhoArquivo = pastaBase & "Acordo2023.xlsm"
            Case "Acordo SINDIFAZ"
                caminhoArquivo = pastaBase & "AcordoSindifaz.xlsm"
            Case "Acordo S� Laguz"
                caminhoArquivo = pastaBase & "AcordoSoLaguz.xlsm"
            Case "Acordo OR�.2021"
                caminhoArquivo = pastaBase & "AcordoOrc2021.xlsm"
            Case "Acordo OR�.2022"
                caminhoArquivo = pastaBase & "AcordoOrc2022.xlsm"
            Case "Acordo OR�.2023"
                caminhoArquivo = pastaBase & "AcordoOrc2023.xlsm"
            Case "Acordo OR�.2024"
                caminhoArquivo = pastaBase & "AcordoOrc2024.xlsm"
        End Select
    End If

    ' Verificando a OrdemBox
    If OrdemBox.Value <> "" Then
        Select Case OrdemBox.Value
            Case "Ordem At� 2020"
                caminhoArquivo = pastaBase & "Ordem2020.xlsm"
            Case "Ordem OR�.2021"
                caminhoArquivo = pastaBase & "Ordem2021.xlsm"
            Case "Ordem OR�.2022"
                caminhoArquivo = pastaBase & "Ordem2022.xlsm"
            Case "Ordem OR�.2023"
                caminhoArquivo = pastaBase & "Ordem2023.xlsm"
            Case "Ordem OR�.2024"
                caminhoArquivo = pastaBase & "Ordem2024.xlsm"
        End Select
    End If

    ' Verificando a PreferenciaBox1
    If PreferenciaBox1.Value <> "" Then
        Select Case PreferenciaBox1.Value
            Case "Prefer�ncia OR�.2022"
                caminhoArquivo = pastaBase & "Prefe2022.xlsm"
            Case "Prefer�ncia OR�.2023"
                caminhoArquivo = pastaBase & "Prefe2023.xlsm"
            Case "Prefer�ncia OR�.2024"
                caminhoArquivo = pastaBase & "Prefe2024.xlsm"
        End Select
    End If

    ' Se um arquivo foi selecionado, abrir o arquivo
    If caminhoArquivo <> "" Then
        ' Verifica se o arquivo existe antes de tentar abri-lo
        If Dir(caminhoArquivo) <> "" Then
            Workbooks.Open caminhoArquivo
        Else
            MsgBox "O arquivo n�o foi encontrado: " & caminhoArquivo, vbExclamation
        End If
    Else
        MsgBox "Por favor, selecione uma op��o para acessar a planilha!", vbExclamation
    End If
End Sub

Private Sub NPrecatorio_Change()
    Dim numero As String
    numero = VBA.Replace(Me.NPrecatorio.Text, "-", "")
    numero = VBA.Replace(numero, ".", "")
    
    ' Garante que s�o apenas n�meros e limita a 20 caracteres (sem os separadores)
    numero = Left(numero, 21)
    
    If Len(numero) >= 21 Then
        Me.NPrecatorio.Text = Left(numero, 7) & "-" & Mid(numero, 8, 2) & "." & Mid(numero, 10, 4) & "." & Mid(numero, 14, 1) & "." & Mid(numero, 15, 2) & "." & Mid(numero, 17, 4)
        Me.NPrecatorio.SelStart = Len(Me.NPrecatorio.Text) ' Mant�m o cursor no final
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
    
    ' Capturar o n�mero digitado no TextBox
    precatorio = Me.NPrecatorio.Text
    
    ' Buscar a c�lula correspondente ao n�mero digitado (Coluna E da planilha BANCO)
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
        
        ' Preencher os TextBox do formul�rio
        Me.DataProtocolo.Text = protocolo
        Me.Exequente.Text = Exequente
        Me.Executado.Text = Executado
        Me.Orcamento.Text = Orcamento
        Me.DataHomologado.Text = DataHomologado
        
        ' Atualizar as c�lulas na planilha Menu
        wsMenu.Range("NumeroPrecatorio").Value = precatorio
        wsMenu.Range("DataProtocolo").Value = protocolo
        wsMenu.Range("NExequente").Value = Exequente
        wsMenu.Range("NExecutado").Value = Executado
        wsMenu.Range("Porcamento").Value = Orcamento
        wsMenu.Range("DataHomologado").Value = DataHomologado
    Else
        ' Mensagens padr�o se n�o encontrar o precat�rio
        Me.DataProtocolo.Text = "Digite a Data do Protocolo"
        Me.Exequente.Text = "Digite Nome do Exequente"
        Me.Executado.Text = "Digite Nome do Executado"
        Me.Orcamento.Text = "Digite o or�amento"
        Me.DataHomologado.Text = "Digite a Data do C�lculo Homologado"
        
        ' Limpar as c�lulas da planilha Menu
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
    
    ' Verifica se os valores digitados s�o n�meros v�lidos
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


