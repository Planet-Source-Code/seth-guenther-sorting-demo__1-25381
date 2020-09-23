Attribute VB_Name = "modMergeSort"
Option Explicit
'///////////////////////////////////////////////////////////
'modMergeSort
'Author: Seth Guenther
'Created: 7/20/2001
'Last modified: 7/23/2001
'
'This module contains the MergeSort algorithm for sorting
'arrays.  The algorithm is as follows:  recursively divide
'the array in half, each time merging the two halves back
'together in sorted order.  At the most atomic level,
'the algorithm merges two arrays of length 1 together,
'then two arrays of length 2, and so on.  Each time, the
'values within the two halves are already sorted.
'///////////////////////////////////////////////////////////

Public Sub mergeSort(ByRef t() As Variant, ByVal first As Double, ByVal last As Double)
'The mergeSort procedure recursively divides the array in half,
'eventually splitting it up into individual elements.  It then
'calls the merge procedure to recombine the elements in sorted
'order.
    Dim mid As Double
    
    mid = (first + last) \ 2     'Find the midpoint of the array
    
    If first <> last Then
        mergeSort t, first, mid     'sort the left half
        mergeSort t, mid + 1, last  'sort the right half
        merge t, first, mid, last   'merge the two halves back together
    End If
End Sub

Private Sub merge(ByRef t() As Variant, ByVal first As Double, ByVal mid As Double, ByVal last As Double)
'The merge procedure combines the two halves of t - t(first...mid) and
't(mid+1...last) - back into one sorted array.  The values within the two
'halves are already sorted (the beauty of recursion), so all that is left to
'be done is to combine the halves, one element at a time, in sorted order.
    Dim length, i, j, k As Double
    
    length = last - first + 1   'determine the length of the array
    ReDim tempArray(length) As Variant  'create a temporary array to hold the sorted values
    
    i = first       'set up indices, i is used to traverse the left half
    j = mid + 1     'j for the left half
    k = 0           'and k for the entire array
    
    'Loop through each half of t, one element at a time.  If the element in the
    'left half is less than (or comes before) the corresponding element in the
    'right half, then put the left element in the temporary array, and vice versa,
    'each time moving to the next position in the temporary array.  Continue this
    'this process until we reach the end of one of the halves.
    Do While (i <= mid And j <= last)
        If t(i) < t(j) Then    '*** Change to > to sort in descending order ***
            tempArray(k) = t(i)
            i = i + 1
            k = k + 1
        Else
            tempArray(k) = t(j)
            j = j + 1
            k = k + 1
        End If
    Loop
    
    'Now place any leftover values into the temporary array.  No
    'comparison is required because the leftover values are sorted.
    Do While (i <= mid)     'leftover from left half
        tempArray(k) = t(i)
        k = k + 1
        i = i + 1
    Loop
    Do While (j <= last)    'leftover from right half
        tempArray(k) = t(j)
        k = k + 1
        j = j + 1
    Loop
    
    'Now copy the values from the temporary array back into the main array.
    For i = first To last
        t(i) = tempArray(i - first)
    Next i
End Sub
