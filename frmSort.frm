VERSION 5.00
Begin VB.Form frmSort 
   AutoRedraw      =   -1  'True
   Caption         =   "Sorting Example"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   Icon            =   "frmSort.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Algorithm to use"
      Height          =   2295
      Left            =   2760
      TabIndex        =   15
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton opSelect 
         Caption         =   "Selection Sort"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton opInsert 
         Caption         =   "Insertion Sort"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton opMerge 
         Caption         =   "Merge Sort"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton opQuick 
         Caption         =   "Quick Sort"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton opBubble 
         Caption         =   "Bubble Sort"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2535
      Begin VB.TextBox numElements 
         Height          =   360
         Left            =   1680
         TabIndex        =   0
         Text            =   "50"
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox maxValue 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Text            =   "100"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "How many numbers?"
         Height          =   330
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label Label2 
         Caption         =   "Max value:"
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sorted"
      Height          =   1575
      Left            =   1440
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
      Begin VB.TextBox txtSorted 
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Unsorted"
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
      Begin VB.TextBox txtUnsorted 
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Default         =   -1  'True
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblStats 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   4215
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'////////////////////////////////////////////
'frmSort
'Author: Seth Guenther
'Created: 7/20/01
'Last modified: 7/23/01
'
'This is a simple project to demonstrate the
'use of sorting algorithms.
'////////////////////////////////////////////

Private Sub Form_Load()
'Use system time to seed random number generator
    Randomize Timer
End Sub

Private Sub cmdSort_Click()
'Creates an array of random values and sorts using
'algorithm specified by user.
    Dim arrayLen, i As Double
    Dim startTime, stopTime As Date
    Dim algorithm As String
    
    arrayLen = Val(numElements)     'determine user length
    'ReDim is used to create a dynamic array (one
    'whose size is variable)
    ReDim unsortedArray(arrayLen) As Variant
    ReDim arrayToSort(arrayLen) As Variant
        
    'Fill the array with random integer values (although
    'floating point and ASCII values are also
    'permissible)
    For i = 0 To arrayLen - 1
        arrayToSort(i) = Int(Rnd * Val(maxValue)) + 1
        unsortedArray(i) = arrayToSort(i)
    Next i

    'Sort using specified algorithm
    startTime = Now     'record time before sort
    If opBubble.Value Then
        algorithm = "Bubble"
        BubbleSort arrayToSort, arrayLen - 1
    ElseIf opSelect.Value Then
        algorithm = "Selection"
        SelectionSort arrayToSort, arrayLen - 1
    ElseIf opInsert.Value Then
        algorithm = "Insertion"
        InsertionSort arrayToSort, arrayLen - 1
    ElseIf opMerge.Value Then
        algorithm = "Merge"
        mergeSort arrayToSort, 0, arrayLen - 1
    Else
        algorithm = "Quick"
        QuickSort arrayToSort, 0, arrayLen - 1
    End If
    stopTime = Now      'record time after sort
   
    'Display the time it took to sort array
    lblStats = "Array sorted using " & algorithm & " Sort in " & _
    DateDiff("s", startTime, stopTime) & " seconds"

    txtUnsorted = ""    'clear old values
    txtSorted = ""
    
    'Display sorted/unsorted arrays if user requests
    If MsgBox("Display unsorted and sorted arrays?", vbYesNo, "Question") = vbYes Then
        For i = 0 To arrayLen - 1
            txtUnsorted = txtUnsorted & unsortedArray(i) & vbCrLf
            txtSorted = txtSorted & arrayToSort(i) & vbCrLf
        Next i
    End If
End Sub
   
