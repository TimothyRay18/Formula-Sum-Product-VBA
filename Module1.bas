Attribute VB_Name = "Module1"
Function getMaxRow(col As Integer) As Double
    getMaxRow = ActiveSheet.cells(Rows.Count, col).End(xlUp).row
End Function
Sub SetFormulaCritical()
Attribute SetFormulaCritical.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Dim row As Double
    Dim col As Double
    Dim maxrow1 As String
    Dim maxrow2 As String
    Dim maxrow3 As String
    row = ActiveCell.row
    col = ActiveCell.Column
    maxrow1 = CStr(getMaxRow(1) - row)
    maxrow2 = CStr(getMaxRow(1) - row - 1)
    maxrow3 = CStr(getMaxRow(1) - row - 2)
    
    ActiveCell.Formula2R1C1 = _
        "=SUMPRODUCT((R[5]C:R[" + maxrow1 + "]C=RC[-1])*(SUBTOTAL(103,OFFSET(R[5]C,ROW(R[5]C:R[" + maxrow1 + "]C)-MIN(ROW(R[5]C:R[" + maxrow1 + "]C)),0))))"
    'Range("BC17").Select
    cells(row, (col + 1)).Select
    ActiveCell.Formula2R1C1 = _
        "=SUMPRODUCT((R[5]C:R[" + maxrow1 + "]C=RC[-2])*(SUBTOTAL(103,OFFSET(R[5]C,ROW(R[5]C:R[" + maxrow1 + "]C)-MIN(ROW(R[5]C:R[" + maxrow1 + "]C)),0))))"
    'Range("BB18").Select
    cells((row + 1), col).Select
    ActiveCell.Formula2R1C1 = _
        "=SUMPRODUCT((R[4]C:R[" + maxrow2 + "]C=RC[-1])*(SUBTOTAL(103,OFFSET(R[4]C,ROW(R[4]C:R[" + maxrow2 + "]C)-MIN(ROW(R[4]C:R[" + maxrow2 + "]C)),0))))"
    'Range("BC18").Select
    cells((row + 1), (col + 1)).Select
    ActiveCell.Formula2R1C1 = _
        "=SUMPRODUCT((R[4]C:R[" + maxrow2 + "]C=RC[-2])*(SUBTOTAL(103,OFFSET(R[4]C,ROW(R[4]C:R[" + maxrow2 + "]C)-MIN(ROW(R[4]C:R[" + maxrow2 + "]C)),0))))"
    'Range("BB19").Select
    cells((row + 2), col).Select
    ActiveCell.Formula2R1C1 = _
        "=SUMPRODUCT((R[3]C:R[" + maxrow3 + "]C=RC[-1])*(SUBTOTAL(103,OFFSET(R[3]C,ROW(R[3]C:R[" + maxrow3 + "]C)-MIN(ROW(R[3]C:R[" + maxrow3 + "]C)),0))))"
    'Range("BC19").Select
    cells((row + 2), (col + 1)).Select
    ActiveCell.Formula2R1C1 = _
        "=SUMPRODUCT((R[3]C:R[" + maxrow3 + "]C=RC[-2])*(SUBTOTAL(103,OFFSET(R[3]C,ROW(R[3]C:R[" + maxrow3 + "]C)-MIN(ROW(R[3]C:R[" + maxrow3 + "]C)),0))))"
End Sub
