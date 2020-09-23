VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort array & remove duplicates v1.1"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1440
      Width           =   4575
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Written by Olof Larsson (kalebeck@hotmail.com)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long 'used to measure the speed, not essential for the code to work

'------------------------------------
'
'   This project is made by Olof Larsson
'   © 2006, kalebeck@hotmail.com
'
'   The sortingalgorithm is written by Philippe Lord // Marton
'    Email:      StromgaldMarton@Hotmail.com
'    ICQ:        12181387
'
'   TriQuickSortString ' Sorts the string array.  // when the distance gets below 5, which speeds things A LOT (over 40%).
'
'   The following code is a demonstration of how to remove duplicates
'   from a string array as quickly as possible. The program first uses
'   the very fast TriQuickSort algorithm to sort the array after it has
'   been dimensioned and populated. Then it uses the remdups sub to remove
'   any possible duplicates from the array. remdups assumes that you don't
'   want any vbNullString in your array, if you would, then just replace
'   vbNullString in the remdups sub with any other character, like Chr$(1),
'   or whatever that floats your boat. This code can be used to remove
'   duplicates from arrays that contains hundreds of thousands of entries,
'   even millions. And it's very fast.
'
'   I hope you enjoy it!
'
'   Enjoy!
'
'------------------------------------

Private Sub Form_Load()


    '-----------------------------
    ' The following code demonstrates on how to create your array and populate it using a file
    ' You can of course use any other normal way to populate the array with strings
    '
    ' This code will load a file with 182,193 entries and remove the duplicates, it will
    ' also measure how fast this process is completed on your computer
    '-----------------------------
    
    
    '-----------------------------
    ' Opens the file muff.txt in the application directory and uses it
    ' to populate the array to demonstrate how the code works
        
    Dim ho() As String, g As Long, tim As Long
    ReDim ho(0) As String
    
    Open App.Path & "\muff.txt" For Input As #1
    Dim a As String, total As Long
    Do Until EOF(1)
        Line Input #1, a
        If g >= UBound(ho) Then
            ReDim Preserve ho(UBound(ho) + 20000) As String
        End If
        total = total + 1
        ho(g) = a
        g = g + 1
    Loop
    Close #1
    ReDim Preserve ho(total) As String
    
    '-----------------------------
 
    g = GetTickCount 'measures the speed of the process
    
    '-----------------------------
    ' This is what does the big job
    
    TriQuickSortString ho 'sorts your string array
    remdups ho 'removes all duplicates
    
    '-----------------------------
    
    Text1.Text = total - UBound(ho) - 1 & " duplicates removed in " & Round((GetTickCount - g) / 1000, 3) & " seconds" & vbCrLf & vbCrLf & "Items left in the array: " & (UBound(ho) + 1) & vbCrLf & "Original size: " & total
    
    '-----------------------------
    ' Prints the contents of your array to the listbox, after the duplicates have been removed
    
    'For g = 0 To UBound(ho)
    '    List1.AddItem Chr$(34) & ho(g) & Chr$(34)
    'Next g
    
    '-----------------------------
    
    '-----------------------------
    ' Prints the contents of your array to the file output.txt in the
    ' application directory, after the duplicates have been removed
    
    Open App.Path & "\output.txt" For Output As #1
    For g = 0 To UBound(ho)
        Print #1, ho(g)
    Next g
    Close #1
    
    '-----------------------------
    
End Sub
