VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Exemplo 1 Combobox"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub atualizaMeses()

    With ComboBox_Mes
        .AddItem "Janeiro"
        .AddItem "Fevereiro"
        .AddItem "Marco"
        .AddItem "Abril"
        .AddItem "Maio"
        .AddItem "Junho"
        .AddItem "Julho"
        .AddItem "Agosto"
        .AddItem "Setembro"
        .AddItem "Outubro"
        .AddItem "Novembro"
        .AddItem "Dezembro"
    
    End With

End Sub

Private Sub UserForm_Initialize()

   Call atualizaMeses

End Sub
