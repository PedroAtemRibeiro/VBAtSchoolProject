Attribute VB_Name = "M�dulo1"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+1"
    Range("A3").Select
End Sub
