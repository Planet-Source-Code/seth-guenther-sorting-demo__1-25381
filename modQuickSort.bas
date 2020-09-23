Attribute VB_Name = "modQuickSort"
Option Explicit
'///////////////////////////////////////////////////////////////
'modQuickSort
'Author: Seth Guenther
'Created: 7/22/01
'Last modified: 7/23/01
'
'This module contains the QuickSort algorithm for sorting
'arrays.  The algorithm is as follows: define a pivot value
'such that the array t is partitioned into two subarrays:
't(0..pivot-1) and t(pivot+1..n).  Recursively partition
'the left and right partitions.  To partition the array,
'define two index placeholders: up and down.  Initialize
'up to the first index in the partition, and down to the
'last.  Also initialize a pivot value to be the first value
'in the partition.  Move up to the index of the first value
'greater than the pivot value, and move down to the index of
'the first value less than the pivot value.  Swap the values,
'and repeat the process until up passes down.  Swap the pivot
'value and the value at down, and set the index of the
'pivot value to down.
'///////////////////////////////////////////////////////////////

Public Sub QuickSort(ByRef t() As Variant, ByVal first As Double, ByVal last As Double)
'This procedure recursivly sorts an array
'using QuickSort.
    Dim pivot As Double
    
    'When first=last, array cannot be partitioned further.
    If first < last Then
        pivot = Partition(t, first, last)   'find the pivot by partitioning
        QuickSort t, first, pivot - 1       'sort the left partition
        QuickSort t, pivot + 1, last        'sort the right partition
    End If
End Sub

Public Function Partition(ByRef t() As Variant, ByVal first As Double, ByVal last As Double) As Double
'This function partitions the array t(first..last) into
't(first..pivot-1) whose values are all less than the pivot
'value, and t(pivot+1..last) whose values are all greater than
'the pivot value.  The return value of the function is the index
'of the pivot value.
'Note:  To sort in descending order, switch the signs in the
'       following comparisons:  t(up) <= pivot and
'                               t(down) >= pivot
        
    Dim up, down As Double
    Dim pivot As Variant
    
    pivot = t(first)    'set the pivot value to the first value in the partition
    up = first          'intialize up and down
    down = last
    
    'Loop until up passes down
    Do While (up < down)
        'Increment up until the first value greater than the
        'pivot value is reached
        Do While (t(up) <= pivot) And (up < last)
            up = up + 1
        Loop
        'Decrement down until the first value less than the
        'pivot value is reached
        Do While (t(down) > pivot)
            down = down - 1
        Loop
        'If up has not passed down, swap the two values
        If up < down Then Swap t(up), t(down)
    Loop
    
    'Swap the pivot value with the value at down; the pivot value
    'is now in the center of the partition.
    Swap t(first), t(down)
    'Return down as the index of the new pivot value
    Partition = down
End Function
