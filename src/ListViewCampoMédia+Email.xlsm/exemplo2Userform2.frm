VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exemplo2Userform2 
   Caption         =   "Exemplo Combobox 2"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000.001
   OleObjectBlob   =   "exemplo2Userform2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "exemplo2Userform2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public totalLinhas As Integer
Public Alunos As Worksheet
Public AlunosLinha As Integer


Private Sub ComboBox_ID_Change()

    If ComboBox_ID.ListIndex = -1 Then
    
        Exit Sub
    
    End If
    
    AlunosLinha = ComboBox_ID.ListIndex + 2
    Call atualizaCampos
    

End Sub

Sub atualizaCampos()

    With Alunos
    
        txtNome.Value = .Cells(AlunosLinha, 3).Value
        txtNota1.Value = .Cells(AlunosLinha, 4).Value
        txtNota2.Value = .Cells(AlunosLinha, 5).Value
        txtNota3.Value = .Cells(AlunosLinha, 6).Value
    
    End With
    
    Dim media As Integer
    media = (Val(txtNota1) + Val(txtNota2) + Val(txtNota3)) / 3
    
    If media >= 6 Then
    
        lblMedia = "Media (" & media & ") Aprovado(a)"
    
    Else
    
        lblMedia = "Media (" & media & ") Reprovado(a)"
    
    End If

End Sub

Private Sub UserForm_Initialize()

    Set Alunos = Worksheets("Alunos")
    
    'TotalLinhas = Planilha1.UsedRange.Rows.Count
    totalLinhas = Sheet1.UsedRange.Rows.Count
    
    If totalLinhas > 1 Then
    
        With ComboBox_ID
            
            .RowSource = "Alunos!B2:B" & totalLinhas
            
        End With
    
    End If

End Sub
