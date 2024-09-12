# macrosExcel
Codigo de Macro para Excel

Cambiar el tipo de numero (cedula) en excel de texto a numero

Sub Poner_numeros()
'
' Macro1 Macro
For i = 1 To 3744
    Dim valor
    valor = ActiveCell.Value
    ActiveCell.FormulaR1C1 = CLngLng(valor)
    ActiveCell.Offset(1, 0).Select
    
Next
End Sub
