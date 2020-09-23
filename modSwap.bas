Attribute VB_Name = "modSwap"
Option Explicit
'//////////////////////////////////////////////////
'modSwap
'Author: Seth Guenther
'Created: 7/22/01
'Last modified: 7/23/01
'//////////////////////////////////////////////////

Public Sub Swap(ByRef x As Variant, ByRef y As Variant)
'Swaps the values in the memory locations referred to
'by x and y.
    Dim temp As Variant
    
    temp = x
    x = y
    y = temp
End Sub
