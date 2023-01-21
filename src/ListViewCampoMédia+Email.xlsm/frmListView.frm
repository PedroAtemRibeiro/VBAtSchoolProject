VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListView 
   Caption         =   "Modelo ListView"
   ClientHeight    =   8700.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15945
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
Const colData As Integer = 7
Const colMateria As Integer = 8
Private idSelecionado As Long
Dim graficoNumero As Integer

Private Sub btnAlterar_Click()

If txtRegistro <> "" And txtID <> "" And txtNome <> "" And txtNota1 <> "" And txtNota2 <> "" And txtNota3 <> "" And txtData <> "" And txtMateria <> "" Then

     txtInstrucoes = "Alterar"
    
    btnSalvar.Enabled = True
    btnAlterar.Enabled = True
    btnExcluir.Enabled = True
    cmdNovo.Enabled = True
    
    Call Habilitar
    
Else
    
    MsgBox "Por favor, Selecione a linha que deseja alterar"

End If
    

End Sub

Private Sub AtualizarInformacoes(ByVal id As Long, ByVal idSelecionado As Long)

    With Sheet1
    
        .Cells(idSelecionado, colRegistro).Value = txtRegistro.Value
        .Cells(idSelecionado, colID).Value = txtID.Value
        .Cells(idSelecionado, colNome).Value = txtNome.Value
        .Cells(idSelecionado, colNota1).Value = txtNota1.Value
        .Cells(idSelecionado, colNota2).Value = txtNota2.Value
        .Cells(idSelecionado, colNota3).Value = txtNota3.Value
        .Cells(idSelecionado, colData).Value = txtData.Value
        .Cells(idSelecionado, colMateria).Value = txtMateria.Value
    
    End With

End Sub

Private Sub AtualizaListView()

    ListViewAluno.ListItems.Clear
    
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
    
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colData).Value
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colMateria).Value
    Next linhaAtual

        Call pintaLinhasAbaixoMedia
        
        
        
        
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
    btnExcluir.Enabled = True
    
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
        Range("A3").AutoFill Destination:=Range("A3:A" & Range("B" & Rows.Count).End(xlUp).Row)
        
        
        
        btnSalvar.Enabled = False
        
        
        
        Call desabilitar
        
        btnSalvar.Enabled = False
        btnExcluir.Enabled = True
            
 
    
    
    End If
    
    Call AtualizaListView
    Call UserForm_Initialize
            
        

End Sub

Private Sub btnSalvar_Click()
    If txtID <> "" And txtNome <> "" And txtNota1 <> "" And txtNota2 <> "" And txtNota3 <> "" And txtData <> "" And txtMateria <> "" Then


    If txtInstrucoes = "Alterar" Then
        
        idSelecionado = txtRegistro.Value + 1
        
       
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
        
        
        
        Call CalculaListView
        
        
        Call desabilitar
        
        
        btnExcluir.Enabled = True
        cmdNovo.Enabled = True
        btnAlterar.Enabled = True
        btnCancelar.Enabled = True
        btnSalvar.Enabled = False
        
        
        
        Sheets("Alunos").Select
        
        [A2] = "1"
        [A3] = "=R[-1]C+1"
        Range("A3").AutoFill Destination:=Range("A3:A" & Range("B" & Rows.Count).End(xlUp).Row)
    
        Call AtualizaListView
    
        End If
    
    
    Else
    
    
        Call desabilitar
End If
    
    Call UserForm_Initialize
    
End Sub
Sub desabilitar()
    txtRegistro.Enabled = False
    txtID.Enabled = False
    txtNome.Enabled = False
    txtNota1.Enabled = False
    txtNota2.Enabled = False
    txtNota3.Enabled = False
    txtData.Enabled = False
    txtMateria.Enabled = False
    
    
    txtRegistro.BackColor = &H8000000F
    txtID.BackColor = &H8000000F
    txtNome.BackColor = &H8000000F
    txtNota1.BackColor = &H8000000F
    txtNota2.BackColor = &H8000000F
    txtNota3.BackColor = &H8000000F
    txtData.BackColor = &H8000000F
    txtMateria.BackColor = &H8000000F


End Sub
Sub Habilitar()
    txtRegistro.Enabled = False
    txtID.Enabled = False
    txtNome.Enabled = True
    txtNota1.Enabled = True
    txtNota2.Enabled = True
    txtNota3.Enabled = True
    txtData.Enabled = True
    txtMateria.Enabled = True
    
    
    txtRegistro.BackColor = &H8000000F
    txtID.BackColor = &H8000000F
    txtNome.BackColor = &H80000005
    txtNota1.BackColor = &H80000005
    txtNota2.BackColor = &H80000005
    txtNota3.BackColor = &H80000005
    txtData.BackColor = &H80000005
    txtMateria.BackColor = &H80000005


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
    
    
    ProgressBar1.Visible = False
    

End Sub

Private Sub cmdExportar_Click()
    
    ProgressBar1.Visible = True
    

    Sheets("Relatorio").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    
    Dim i As Long
    Dim linha As Integer
    linha = 2
    
    ProgressBar1.Max = CInt(ListViewAluno.ListItems.Count)
    
    For i = 1 To Me.ListViewAluno.ListItems.Count
    
        If Me.ListViewAluno.ListItems.Item(i).Checked Then
        
            [A1] = "Registro"
            [B1] = "ID"
            [C1] = "Nome"
            [D1] = "Nota 1"
            [E1] = "Nota 2"
            [F1] = "Nota 3"
            [G1] = "Data"
            [H1] = "Materia"
            
            Sheets("Relatorio").Cells(linha, 1) = Me.ListViewAluno.ListItems.Item(i).Text
            Sheets("Relatorio").Cells(linha, 2) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(1).Text
            Sheets("Relatorio").Cells(linha, 3) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(2).Text
            Sheets("Relatorio").Cells(linha, 4) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(3).Text
            Sheets("Relatorio").Cells(linha, 5) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(4).Text
            Sheets("Relatorio").Cells(linha, 6) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(5).Text
            Sheets("Relatorio").Cells(linha, 7) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(8).Text
            Sheets("Relatorio").Cells(linha, 8) = Me.ListViewAluno.ListItems.Item(i).ListSubItems(9).Text
            
            If linha > ListViewAluno.ListItems.Count Then
                
                    
                    
            Else
                
                ProgressBar1.Value = linha
            
            
            End If
            
            
            linha = linha + 1
        
        End If
    
    Next

End Sub

Private Sub cmdNovo_Click()

    txtInstrucoes = "Novo"
    
    Dim ultimaLinha As Integer
        
    ultimaLinha = ListViewAluno.ListItems.Count
    
    txtNome = ""
    txtNota1 = ""
    txtNota2 = ""
    txtNota3 = ""
    txtData = ""
    txtMateria = ""
    
    
    
    
    btnSalvar.Enabled = True
    cmdNovo.Enabled = False
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    
    
    
    txtID = ListViewAluno.ListItems.Item(ultimaLinha).ListSubItems(1).Text + 1
    txtRegistro = ListViewAluno.ListItems.Item(ultimaLinha).Text + 1
    
    
    txtID.Enabled = True
    
    
    Call Habilitar
    
    
    
End Sub

Sub Filtrar()

    Dim linha As Long
    Dim textoPesquisa As String
    Dim Data As Date, DataInicio As Date, DataFim As Date
        
    If txtDataInicio <> Empty Then
        
        DataInicio = txtDataInicio.Value
        
    End If
    
    If txtDataFim <> Empty Then
        
        DataFim = txtDataFim.Value
        
    End If
    
       If txtDataInicio = Empty And txtDataFim <> Empty Then
        
        DataInicio = DataFim
        
    End If
    

    If txtDataInicio <> Empty And txtDataFim = Empty Then
        
        DataFim = DataInicio
        
    End If
    
    linha = 2
    
    ListViewAluno.ListItems.Clear
    
        With Sheet1
            
            While .Cells(linha, 1).Value <> Empty
                
                textoPesquisa = .Cells(linha, 3).Value
                
                If UCase(Left(textoPesquisa, Len(ComboBox_Nome.Text))) = UCase(ComboBox_Nome.Text) Then
                
                        textoPesquisa = .Cells(linha, 8).Value
                    
                        If UCase(Left(textoPesquisa, Len(combobox_Materia.Text))) = UCase(combobox_Materia.Text) Then
                    
                        
                                If DataInicio <> Empty And DataFim <> Empty Then
                                
                                    Data = .Cells(linha, 7).Value
                                    
                                    If Data >= DataInicio And Data <= DataFim Then
                                    
                                        
                                        Set lista = ListViewAluno.ListItems.Add(Text:=Cells(linha, 1).Value)
                                        lista.ListSubItems.Add Text:=Cells(linha, 2).Value
                                        lista.ListSubItems.Add Text:=Cells(linha, 3).Value
                                        lista.ListSubItems.Add Text:=Cells(linha, 4).Value
                                        lista.ListSubItems.Add Text:=Cells(linha, 5).Value
                                        lista.ListSubItems.Add Text:=Cells(linha, 6).Value
                                        
                                        
                                        On Error Resume Next
                                        lista.ListSubItems.Add Text:=Format((Sheet1.Cells(linhaAtual, colNota1) + Sheet1.Cells(linhaAtual, colNota2) + Sheet1.Cells(linhaAtual, colNota3)) / 3, "#,#0.0")
                                        lista.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colNome).Value & "@gmail.com"
                                        
                                        lista.ListSubItems.Add Text:=Cells(linha, 7).Value
                                        lista.ListSubItems.Add Text:=Cells(linha, 8).Value
                                    
                                    End If
                                    
                                    
                                Else
                                
                                         Set lista = ListViewAluno.ListItems.Add(Text:=Cells(linha, 1).Value)
                                        lista.ListSubItems.Add Text:=Cells(linha, 2).Value
                                        lista.ListSubItems.Add Text:=Cells(linha, 3).Value
                                        lista.ListSubItems.Add Text:=Cells(linha, 4).Value
                                        lista.ListSubItems.Add Text:=Cells(linha, 5).Value
                                        lista.ListSubItems.Add Text:=Cells(linha, 6).Value
                                        
                                        
                                        On Error Resume Next
                                        lista.ListSubItems.Add Text:=Format((Sheet1.Cells(linha, colNota1) + Sheet1.Cells(linha, colNota2) + Sheet1.Cells(linha, colNota3)) / 3, "#,#0.0")
                                        lista.ListSubItems.Add Text:=Sheet1.Cells(linha, colNome).Value & "@gmail.com"
                                        
                                        lista.ListSubItems.Add Text:=Cells(linha, 7).Value
                                        lista.ListSubItems.Add Text:=Cells(linha, 8).Value
                                    
                                End If
                        
                        End If
                
                End If
                
               linha = linha + 1
            
            Wend
            
        End With
        
        
    
    Set lista = Nothing
    
    Exit Sub
    
End Sub

Private Sub combobox_Materia_Change()

        Call CalculaListView
        
        
        Call Filtrar

End Sub

Private Sub ComboBox_Nome_Change()
                
        Call CalculaListView
        
        
        
        Call Filtrar
        
End Sub

Private Sub ListViewAluno_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If Not ListViewAluno.ListItems.Count = 0 Then
    
        txtRegistro.Text = ListViewAluno.SelectedItem.Text 'Este aqui é o primeiro item a ser carregado
        txtID.Text = ListViewAluno.SelectedItem.SubItems(1)
        txtNome.Text = ListViewAluno.SelectedItem.SubItems(2)
        txtNota1.Text = ListViewAluno.SelectedItem.SubItems(3)
        txtNota2.Text = ListViewAluno.SelectedItem.SubItems(4)
        txtNota3.Text = ListViewAluno.SelectedItem.SubItems(5)
        txtData.Text = ListViewAluno.SelectedItem.SubItems(8)
        txtMateria.Text = ListViewAluno.SelectedItem.SubItems(9)
        
    
    Else
    
        MsgBox "ListView esta vazia ou sem dados,"
    
    End If
    

End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)


        If Not Node.Parent Is Nothing Then
        
                combobox_Materia = Node.Parent
                
                ComboBox_Nome = Node.Text
            
        Else
                   combobox_Materia = Node.Key
                   
                   ComboBox_Nome = ""
        
        End If


End Sub

Private Sub txtDataFim_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        
        

                
        Call CalculaListView
        
        
        Call Filtrar

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
        .ColumnHeaders.Add , , "Data", 90, lvwColumnLeft
        .ColumnHeaders.Add , , "Materia", 90, lvwColumnLeft
        
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
        
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colData).Value
        i.ListSubItems.Add Text:=Sheet1.Cells(linhaAtual, colMateria).Value
    Next linhaAtual
    
    CheckBox_Selecao.Value = False
    CheckBox_Selecao.Caption = "Selecionar tudo"
    
    Call CalculaListView
    Call pintaLinhasAbaixoMedia
    Call desabilitar
    
    Dim totalLinha As Long
    Dim linha As Long
    
    
    totalLinha = Sheets("Alunos").UsedRange.Rows.Count
    
    For linha = 2 To totalLinha
    
        combobox_Materia.AddItem Sheets("Alunos").Range("H" & linha).Value
    
    Next linha
    
    
    For lista1 = 0 To totalLinha
        
            For lista2 = totalLinha - 1 To lista1 + 1 Step -1
            
                If combobox_Materia.List(lista2) = combobox_Materia.List(lista1) Then
                
                    combobox_Materia.RemoveItem (lista2)
                    
                End If
            
            Next lista2
        
        
    Next lista1
    
    '------------------------------------------------------------------------------------------------------
    
    linha = 0
    
    
    
    totalLinha = Sheets("Alunos").UsedRange.Rows.Count
    
    For linha = 2 To totalLinha
    
        ComboBox_Nome.AddItem Sheets("Alunos").Range("C" & linha).Value
    
    Next linha
    
    
    For SegundaLista1 = 0 To totalLinha
        
            For SegundaLista2 = totalLinha - 1 To SegundaLista1 + 1 Step -1
            
                If ComboBox_Nome.List(SegundaLista2) = ComboBox_Nome.List(SegundaLista1) Then
                
                    ComboBox_Nome.RemoveItem (SegundaLista2)
                    
                End If
            
            Next SegundaLista2
        
        
    Next SegundaLista1

    
    ordenaItemsComboBox ComboBox_Nome.name
    ordenaItemsComboBox combobox_Materia.name
    
    Dim ultimaLinhaListView As Double
    ultimaLinhaListView = ListViewAluno.ListItems.Count
    
    If ultimaLinhaListView >= "1" Then
    
        ListViewAluno.ListItems(ultimaLinhaListView).Selected = True
        ListViewAluno.ListItems(ultimaLinhaListView).EnsureVisible
    
    End If
    
    Dim planilhaSelecionada As Worksheet
    Set planilhaSelecionada = Sheets("Dados Combobox Dependente")
    
    With planilhaSelecionada
    
        TreeView1.Nodes.Add Key:="Matematica", Text:="Matematica"
        TreeView1.Nodes.Add Key:="Fisica", Text:="Fisica"
        TreeView1.Nodes.Add Key:="Quimica", Text:="Quimica"
        TreeView1.Nodes.Add Key:="Portugues", Text:="Portugues"
    
    End With
    
    Call preencheNosFilhos("Matematica")
    Call preencheNosFilhos("Fisica")
    Call preencheNosFilhos("Quimica")
    Call preencheNosFilhos("Portugues")
    
End Sub
Private Sub preencheNosFilhos(ByVal pai As String)
    
    With Planilha2
    
        Dim contador As Long
        contador = 2
        
        Do Until Sheets("Dados Combobox Dependente").Cells(contador, 1) = "Parar"
        
            If Sheets("Dados Combobox Dependente").Cells(contador, 3) = pai Then
            
                    TreeView1.Nodes.Add pai, tvwChild, pai + CStr(contador), Sheets("Dados Combobox Dependente").Cells(contador, 2)
                
            Else
            
            
            End If
        
            contador = contador + 1
        Loop
    
    End With


End Sub

Sub ordenaItemsComboBox(DadosCombobox)

    On Error GoTo Erro
    
    Dim posicao As Double
    Dim totalLinhas As Double
    Dim texto As String
    Dim dentroVez1, vez1 As Double
    
    posição = 0
    
    
    totalLinhas = Controls(DadosCombobox).ListCount - 1
    
    For vez1 = posicao To totalLinhas
        
        For dentroVez1 = vez1 To totalLinhas
        
            If Controls(DadosCombobox).List(vez1) > Controls(DadosCombobox).List(dentroVez1) Then
                
                texto = Controls(DadosCombobox).List(dentroVez1)
                Controls(DadosCombobox).List(dentroVez1) = Controls(DadosCombobox).List(vez1)
                Controls(DadosCombobox).List(vez1) = texto
            End If
        
        Next dentroVez1
        
    Next vez1
    
Erro:
    Exit Sub
    
        MsgBox "Erro ao Ordenar Registros"
    
    
    




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

















