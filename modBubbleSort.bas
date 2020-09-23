Attribute VB_Name = "modBubbleSort"
Option Explicit
'///////////////////////////////////////////////////////////////
'modBubbleSort
'Author: Seth Guenther
'Created: 7/22/01
'Last modified: 7/23/01
'
'This module contains the Bubble sort algorithm for sorting
'an array.  The algorithm is as follows: assume the array
'is sorted.  Make one pass through the array.  If any one value
'is greater than the one following (or less than, if sorting
'in descending order) then the array is not sorted.  Swap
'the two values.  Continue this down the entire length of the
'array.  At the end, assume once more the array is sorted,
'and repeat the above process, until all values are sorted.
'///////////////////////////////////////////////////////////////

Public Sub BubbleSort(ByRef t() As Variant, ByVal n As Double)
'This procedure sorts array t of length n using bubble sort.
'Note: by using < instead of > to compare array values
'in the loop below, the procedure will sort the array
'in descending order.
    Dim sorted As Boolean
    Dim i As Double
    
    Do
        sorted = True                'assume array is sorted
        For i = 0 To n - 1           'loop through array values
            If t(i) > t(i + 1) Then  'compare each value with the next one
                Swap t(i), t(i + 1)  'swap if unsorted
                sorted = False       'assumption is wrong; array not sorted
            End If
        Next i
    Loop Until sorted                'repeat process until array is sorted
End Sub
