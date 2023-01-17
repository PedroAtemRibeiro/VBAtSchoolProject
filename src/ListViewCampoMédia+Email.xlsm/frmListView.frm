VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListView 
   Caption         =   "Modelo ListView"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11835
   OleObjectBlob   =   "frmListView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const colRegistro As Integer = 1
Const colID As Integer = 2
Const colNome As Integer = 3
Const colNota1 As Integer = 4
Const colNota2 As Integer = 5
Const colNota3 As Integer = 6
Private idSelecionado As Long
Dim graficoNumero As Integer

Private Sub btnAlterar_Click()

If txtRegistro <> "" And txtID <> "" And txtNome <> "" And txtNota1 <> "" And txtNota2 <> "" And txtNota3 <> "" Then

     txtInstrucoes = "Alterar"
    
    btnSalvar.Enabled = True
    cmdNovo.Enabled = False
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    
    Call Habilitar
    
Else
    
    MsgBox "Por favor, Selecione a linha que deseja alterar"

End If
    

End Sub

Private Sub AtualizarInformacoes(ByVal id As Long, ByVal idSelecionado As Long)

    With Sheet1
    
        .Cells(idSelecionado, colRegistro).Value = txtRegistro.Value + 1
        .Cells(idSelecionado, colID).Value = txtID.Value
        .Cells(idSelecionado, colNome).Value = txtNome.Value
        .Cells(idSelecionado, colNota1).Value = txtNota1.Value
        .Cells(idSelecionado, colNota2).Value = txtNota2.Value
        .Cells(idSelecionado, colNota3).Value = txtNota3.Value
    
    End With

End Sub

Private Sub AtualizaListView()

    ListViewAluno.ListItems.Clear
    
    Dim linhaAtual As Integer
    Dim i
    Dim ultimaLinha
    
    ultimaLinha = Sheet1.Cells(Sheet1.Cells.Rows.Count, colRegistro).End(xlUp).Row 'Vai encontrar a ultima da minha planilha
    For linhaAtual = 2 To ultimaLinha
        
        Set i = ListViewAluno.ListItems.Add(Text:=Format(Sheet1.Cells(linhaAtual, colRegistro).Value + 1, 0))
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colID).Value
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colNome).Value
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colNota1).Value
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colNota2).Value
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colNota3).Value
        
        On Error Resume Next
        i.ListSubItems.Add Text:=Format((Sheet1.Cells(linhaAtual, colNota1) + Sheet1.Cells(linhaAtual, colNota2) + Sheet1.Cells(linhaAtual, colNota3)) / 3, "#,#0.0")
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colNome).Value & "@gmail.com"
    
    Next linhaAtual

End Sub

Sub pintaLinhasAbaixoMedia()

    Dim linha As Integer
    Dim coluna As Integer
   
    For linha = 1 To ListViewAluno.ListItems.Count
    
        For coluna = 1 To 7
        
            If ListViewAluno.ListItems.Item(linha).SubItems(6) < 6 Then
             
            ListViewAluno.ListItems.Item(linha).ForeColor = RGB(255, 51, 51)
            ListViewAluno.ListItems.Item(linha).ListSubItems(coluna).ForeColor = RGB(255, 51, 51)
            ListViewAluno.ListItems.Item(linha).ListSubItems(coluna).Bold = True
            
            End If
    
        Next coluna
    
    Next linha

End Sub

Private Sub btnCancelar_Click()

    
    txtInstrucoes = "Cancelar"
    
    btnSalvar.Enabled = False
    cmdNovo.Enabled = True
    btnAlterar.Enabled = True
    btnExcluir.Enabled = False
    
    Call Habilitar
    
End Sub

Private Sub btnExcluir_Click()

    idSelecionado = txtRegistro.Value + 1
    
    
    Dim confirmar As VbMsgBoxResult
    confirmar = MsgBox("Deseja mesmo excluir o registro " & txtRegistro & "?", vbYesNo, "Confirmar")
    
    
    If confirmar = vbYes Then
        
        Sheet1.Range(Sheet1.Cells(idSelecionado, colRegistro), _
        Sheet1.Cells(idSelecionado, colRegistro)).EntireRow.Delete
        
        Sheets("Alunos").Select
        
        [A2] = "1"
        [A3] = "=R[-1]C+1"
        Range("A3").AutoFill Destination:=Range("A3:A" & Range("b" & Rows.Count).End(xlUp).Row)
        
              
        Call AtualizarInformacoes(CLng(txtRegistro.Value), idSelecionado)
        
        Call AtualizaListView
        
        Call CalculaListView
        Call pintaLinhasAbaixoMedia
        
        btnSalvar.Enabled = False
        btnExcluir.Enabled = False
        
        Call desabilitar
    
            
            
    Else
    
    
    End If
    
    
            
        

End Sub

Private Sub btnSalvar_Click()
    If txtID <> "" And txtNome <> "" And txtNota1 <> "" And txtNota2 <> "" And txtNota3 <> "" Then


    If txtInstrucoes = "Alterar" Then
        
        idSelecionado = txtRegistro.Value + 1
        
        ' Clng -2.147.483.648 a 2.147.483.647
        Call AtualizarInformacoes(CLng(txtRegistro.Value), idSelecionado)
        
        Call AtualizaListView
        
        Call CalculaListView
        Call pintaLinhasAbaixoMedia
    
    
    ElseIf txtInstrucoes = "Novo" Then
    
        Dim QtdRegistro As Integer
        
        QtdRegistro = Format(ListViewAluno.ListItems.Count, 0) + 1
        
        Range("A1048576").End(xlUp).Offset(l, 0) = QtdRegistro
         
        txtRegistro = QtdRegistro
        
    
    
        idSelecionado = txtRegistro.Value + 1
        
        ' Clng -2.147.483.648 a 2.147.483.647
        Call AtualizarInformacoes(CLng(txtRegistro.Value), idSelecionado)
        
        Call AtualizaListView
        
        Call CalculaListView
        Call pintaLinhasAbaixoMedia
        
        btnSalvar.Enabled = False
        btnExcluir.Enabled = False
        
        Call desabilitar
    
    
    
        End If
    
    
    Else
    
    
        Call desabilitar
End If
    
    
End Sub
Sub desabilitar()
    txtRegistro.Enabled = False
    txtID.Enabled = False
    txtNome.Enabled = False
    txtNota1.Enabled = False
    txtNota2.Enabled = False
    txtNota3.Enabled = False
    
    
    txtRegistro.BackColor = &H8000000F
    txtID.BackColor = &H8000000F
    txtNome.BackColor = &H8000000F
    txtNota1.BackColor = &H8000000F
    txtNota2.BackColor = &H8000000F
    txtNota3.BackColor = &H8000000F


End Sub
Sub Habilitar()
    txtRegistro.Enabled = False
    txtID.Enabled = False
    txtNome.Enabled = True
    txtNota1.Enabled = True
    txtNota2.Enabled = True
    txtNota3.Enabled = True
    
    
    txtRegistro.BackColor = &H8000000F
    txtID.BackColor = &H8000000F
    txtNome.BackColor = &H80000005
    txtNota1.BackColor = &H80000005
    txtNota2.BackColor = &H80000005
    txtNota3.BackColor = &H80000005


End Sub

Private Sub CheckBox_Selecao_Click()

    Dim i As Integer
    
    If CheckBox_Selecao = True Then
    
        'Marcar tudo
        For i = 1 To ListViewAluno.ListItems.Count
        
            If ListViewAluno.ListItems.Item(i).Checked = False Then
            
                ListViewAluno.ListItems.Item(i).Checked = True
            
            End If
            
            CheckBox_Selecao.Value = True
            CheckBox_Selecao.Caption = "Limpar campos"
        
        Next i
    
    Else
    
        
        'Desmarcar tudo
        For i = 1 To ListViewAluno.ListItems.Count
        
            If ListViewAluno.ListItems.Item(i).Checked = True Then
            
                ListViewAluno.ListItems.Item(i).Checked = False
            
            End If
            
            CheckBox_Selecao.Value = False
            CheckBox_Selecao.Caption = "Selecionar tudo"
        
        Next i
    
    End If
    

End Sub

Private Sub cmdExportar_Click()

    Sheets("Relatorio").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    
    Dim i As Long
    Dim linha As Integer
    linha = 2
    
    For i = 1 To Me.ListViewAluno.ListItems.Count
    
        If Me.ListViewAluno.ListItems.Item(i).Checked Then
        
            [A1] = "Registro"
            [B1] = "ID"
            [C1] = "Nome"
            [D1] = "Nota 1"
            [E1] = "Nota 2"
            [F1] = "Nota 3"
        
            Sheets("Relatorio").Cells(linha, 1) = Me.ListViewAluno.ListItems.Item(i).Text
            Sheets("Relatorio").Cells(linha, 2) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(1).Text
            Sheets("Relatorio").Cells(linha, 3) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(2).Text
            Sheets("Relatorio").Cells(linha, 4) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(3).Text
            Sheets("Relatorio").Cells(linha, 5) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(4).Text
            Sheets("Relatorio").Cells(linha, 6) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(5).Text
            
            linha = linha + 1
        
        End If
    
    Next

End Sub

Private Sub cmdNovo_Click()

    txtInstrucoes = "Novo"
    
    Dim ultimaLinha As Integer
        
    ultimaLinha = ListViewAluno.ListItems.Count
    
    
    
    btnSalvar.Enabled = True
    cmdNovo.Enabled = False
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    
    txtID = ListViewAluno.ListItems.Item(ultimaLinha).ListSubItems(1).Text + 1
    txtRegistro = ListViewAluno.ListItems.Item(ultimaLinha).Text + 1
    
    
    txtID.Enabled = True
    
    
    Call Habilitar
    
    
End Sub

Private Sub ComboBox_Nome_Change()

    Dim Pesquisar As String
    Dim ValorEncontrado, ultimaLinha As Long
    Dim linha As Long
    Dim i
    
    Pesquisar = LCase(ComboBox_Nome.Value)
    
    ListViewAluno.ListItems.Clear
    
    If Pesquisar <> Empty Then
    
        ultimaLinha = Sheet1.Cells(Sheet1.Cells.Rows.Count, colRegistro).End(xlUp).Row
        
        For linha = 2 To ultimaLinha
            
            ValorEncontrado = InStr(1, Sheet1.Cells(linha, 3), Pesquisar, vbTextCompare)
                
                If ValorEncontrado > 0 Then
                
                    Set i = ListViewAluno.ListItems.Add(Text:=Format(Sheet1.Cells(linha, colRegistro).Value, "0"))
                    i.ListSubItems.Add Text:=Sheet1.Cells(linha, colID).Value
                    i.ListSubItems.Add Text:=Sheet1.Cells(linha, colNome).Value
                    i.ListSubItems.Add Text:=Sheet1.Cells(linha, colNota1).Value
                    i.ListSubItems.Add Text:=Sheet1.Cells(linha, colNota2).Value
                    i.ListSubItems.Add Text:=Sheet1.Cells(linha, colNota3).Value
                    
                    
                    On Error Resume Next
                    i.ListSubItems.Add Text:=Format((Sheet1.Cells(linha, colNota1) + Sheet1.Cells(linha, colNota2) + Sheet1.Cells(linha, colNota3)) / 3, "#,#0.0")
                    i.ListSubItems.Add Text:=Sheet1.Cells(linha, colNome).Value & "@gmail.com"
                
                End If
            
        
        Next linha
    
    Else
    
            ultimaLinha = Sheet1.Cells(Sheet1.Cells.Rows.Count, colRegistro).End(xlUp).Row
            
            For linha = 2 To ultimaLinha
                
                    
                        Set i = ListViewAluno.ListItems.Add(Text:=Format(Sheet1.Cells(linha, colRegistro).Value, "0"))
                        i.ListSubItems.Add Text:=Sheet1.Cells(linha, colID).Value
                        i.ListSubItems.Add Text:=Sheet1.Cells(linha, colNome).Value
                        i.ListSubItems.Add Text:=Sheet1.Cells(linha, colNota1).Value
                        i.ListSubItems.Add Text:=Sheet1.Cells(linha, colNota2).Value
                        i.ListSubItems.Add Text:=Sheet1.Cells(linha, colNota3).Value
                    
                        On Error Resume Next
                        i.ListSubItems.Add Text:=Format((Sheet1.Cells(linha, colNota1) + Sheet1.Cells(linha, colNota2) + Sheet1.Cells(linha, colNota3)) / 3, "#,#0.0")
                        i.ListSubItems.Add Text:=Sheet1.Cells(linha, colNome).Value & "@gmail.com"
            
            Next linha
    
        
    
    End If
    
    Call CalculaListView
    
    

End Sub





Private Sub ListViewAluno_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If Not ListViewAluno.ListItems.Count = 0 Then
    
        txtRegistro.Text = ListViewAluno.SelectedItem.Text 'Este aqui é o primeiro item a ser carregado
        txtID.Text = ListViewAluno.SelectedItem.SubItems(1)
        txtNome.Text = ListViewAluno.SelectedItem.SubItems(2)
        txtNota1.Text = ListViewAluno.SelectedItem.SubItems(3)
        txtNota2.Text = ListViewAluno.SelectedItem.SubItems(4)
        txtNota3.Text = ListViewAluno.SelectedItem.SubItems(5)
        
    
    Else
    
        MsgBox "ListView esta vazia ou sem dados,"
    
    End If
    

End Sub

Private Sub UserForm_Initialize()

    With ListViewAluno
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Registro", 60, lvwColumnLeft
        .ColumnHeaders.Add , , "Codigo", 60, lvwColumnLeft
        .ColumnHeaders.Add , , "Aluno", 90, lvwColumnLeft
        .ColumnHeaders.Add , , "Nota 1", 60, lvwColumnLeft
        .ColumnHeaders.Add , , "Nota 2", 60, lvwColumnLeft
        .ColumnHeaders.Add , , "Nota 3", 60, lvwColumnLeft
        .ColumnHeaders.Add , , "Media", 60, lvwColumnLeft
        .ColumnHeaders.Add , , "E-mail", 90, lvwColumnLeft
    End With
    
    ListViewAluno.ListItems.Clear
    
    Dim olaMundo
    Dim linhaAtual As Integer
    Dim i
    Dim ultimaLinha
    
    ultimaLinha = Sheet1.Cells(Sheet1.Cells.Rows.Count, colRegistro).End(xlUp).Row 'Vai encontrar a ultima da minha planilha
    For linhaAtual = 2 To ultimaLinha
        
        Set i = ListViewAluno.ListItems.Add(Text:=Format(Sheet1.Cells(linhaAtual, colRegistro).Value, 0))
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colID).Value
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colNome).Value
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colNota1).Value
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colNota2).Value
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colNota3).Value
        
        On Error Resume Next
        i.ListSubItems.Add Text:=Format((Sheet1.Cells(linhaAtual, colNota1) + Sheet1.Cells(linhaAtual, colNota2) + Sheet1.Cells(linhaAtual, colNota3)) / 3, "#,#0.0")
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colNome).Value & "@gmail.com"
    Next linhaAtual
    
    CheckBox_Selecao.Value = False
    CheckBox_Selecao.Caption = "Selecionar tudo"
    
    Call CalculaListView
    Call pintaLinhasAbaixoMedia
    Call desabilitar
    

End Sub

Sub CalculaListView()

    Dim i_1, i_2, i_3 As Currency
    Dim nota1, nota2, nota3 As Currency
    
    nota1 = 0
    nota2 = 0
    nota3 = 0
    
    For i_1 = 1 To ListViewAluno.ListItems.Count
    
        nota1 = nota1 + Me.ListViewAluno.ListItems(i_1).ListSubItems(3)
    
    Next i_1
    
    txtResumoNota1 = nota1
    
    '-----------------------------------
    
    For i_2 = 1 To ListViewAluno.ListItems.Count
    
        nota2 = nota2 + Me.ListViewAluno.ListItems(i_2).ListSubItems(4)
    
    Next i_2
    
    txtResumoNota2 = nota2
    
    '-----------------------------------
    
    For i_3 = 1 To ListViewAluno.ListItems.Count
    
        nota3 = nota3 + Me.ListViewAluno.ListItems(i_3).ListSubItems(5)
    
    Next i_3
    
    txtResumoNota3 = nota3
    
    '-----------------------------------
    
    On Error Resume Next
    Dim QtdRegistro, media1, media2 As Currency
    QtdRegistro = Format(ListViewAluno.ListItems.Count, 0)
    
    
    media1 = (nota1 + nota2 + nota3) / 3
    media2 = media1 / QtdRegistro
    
    lblQtd.Caption = QtdRegistro
    
    txtResumoMedia = Format(media2, "#,#0.0")
    
    If QtdRegistro > 1 Then
    
                lblSituacao.Caption = "Turma: "
                
                If media2 >= 6 Then
                
                    lblSituacao2.Caption = "Acima da media"
                
                Else
                
                    lblSituacao2.Caption = "Abaixo da media"
                
                End If
    
    Else
    
                lblSituacao.Caption = "Aluno: "
                
                If media2 >= 6 Then
                
                    lblSituacao2.Caption = "Aprovado(a)"
                
                Else
                
                    lblSituacao2.Caption = "Reprovado(a)"
                
                End If
    
    
    End If
    
    Sheets("Grafico").Range("B2") = nota1
    Sheets("Grafico").Range("B3") = nota2
    Sheets("Grafico").Range("B4") = nota3
    
    graficoNumero = 1
    Call atualizaGrafico

End Sub

Private Sub atualizaGrafico()

    Set graficoSelecionado = Sheets("Grafico").ChartObjects(graficoNumero).Chart
    graficoSelecionado.Parent.Width = 400
    graficoSelecionado.Parent.Height = 150
    
    'Vamos salvar o grafico no formato de Gif
    Fname = ThisWorkbook.Path & Application.PathSeparator & "graficoo.gif"
    graficoSelecionado.Export fileName:=Fname, FilterName:="GIF"
    
    'Carrega o grafico para o Userform
    imageUserFormGrafico.Picture = LoadPicture(Fname)

End Sub

















