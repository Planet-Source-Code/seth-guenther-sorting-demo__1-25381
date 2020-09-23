Attribute VB_Name = "modInsertionSort"
Option Explicit
'//////////////////////////////////////////////////////////////
'modInsertionSort
'Author: Seth Guenther
'Created: 7/22/01
'Last modified: 7/23/01
'
'This module contains the Insertion Sort algorithm for sorting
'arrays.  The algorithm works as follows: start with an unsorted
'array t of length n.  Define the subarrays t(0) to be sorted,
'and t(1..n-1) to be unsorted.  Insert the first value in the
'unsorted portion of t into the correct place in the sorted
'portion, sliding all other values down one position.  Now
'the subarray t(0..1) is sorted and t(2..n-1) is unsorted.
'Repeat the above process for the entire length of the array.
'//////////////////////////////////////////////////////////////

Public Sub InsertionSort(ByRef t() As Variant, ByVal n As Double)
'This procedure sorts the array t of length n
'using Insertion Sort.
    Dim unsorted, i As Double
    Dim nextValue As Variant
    
    'Loop through t.  It will always be the case that
    't(0..unsorted-1) is sorted, and t(unsorted..n-1)
    'is unsorted.
    For unsorted = 1 To n
        nextValue = t(unsorted)      'next value to sort
        i = unsorted                 'index of next value
        
        'Shift all values down one position until correct
        'position in t(0..unsorted-1) is reached.
        Do While (t(i - 1) > nextValue)     '*** Change to < to sort in descending order ***
            t(i) = t(i - 1)
            i = i - 1
            If i = 0 Then Exit Do    'special case - value is minimal
        Loop
           
        t(i) = nextValue             'insert new value
    Next unsorted
End Sub
