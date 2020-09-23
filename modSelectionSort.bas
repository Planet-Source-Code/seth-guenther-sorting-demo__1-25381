Attribute VB_Name = "modSelectionSort"
Option Explicit
'////////////////////////////////////////////////////////////////
'modSelectionSort
'Author: Seth Guenther
'Created: 7/22/01
'Last modified: 7/23/01
'
'This module contains the Selection Sort algorithm for sorting
'an array.  The algorithm is as follows: beginning at the first
'position in the array, find the smallest value in the array (or
'the largest one, if sorting in descending order.)  Swap that
'value with the value in the first position.  Move to the second
'position in the array, and repeat the above, continuing for the
'entire length of the array.
'////////////////////////////////////////////////////////////////

Public Sub SelectionSort(ByRef t() As Variant, ByVal n As Double)
'This procedure sorts an array t of length n using
'Selection sort.
    Dim i As Double
    
    For i = 0 To n - 1                 'loop through array
        Swap t(i), t(FindMin(t, i, n))  'swap smallest value with
                                        'value at current position
        
        '**************************************
        'Use this statement to sort in descending order
        'Swap t(n - i), t(FindMin(t, 0, n - i))
        '**************************************
    Next i
End Sub

Private Function FindMin(ByRef t() As Variant, ByVal first As Double, ByVal last As Double) As Double
'This procedure finds the index of the smallest
'value in the array t(first...last).
    Dim i, min As Double
    
    min = first                         'assume smallest is at first index
    For i = first + 1 To last           'loop through array, if smaller
        If t(i) < t(min) Then min = i   'value found, change index
    Next i
    
    FindMin = min                       'return index of smallest value

End Function
