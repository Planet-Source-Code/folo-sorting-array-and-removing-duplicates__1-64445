Attribute VB_Name = "Module1"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)
Public Enum SortOrder
   SortAscending = 0
   SortDescending = 1
End Enum

Public Sub remdups(ByRef arr() As String)

    '-----------------------------
    ' Coded by Olof Larsson
    '-----------------------------
    
    Dim c As Long, coun As Long
    coun = UBound(arr)
    Dim g As Long
    c = 1
    For g = 0 To coun
        If g + c > coun Then
            Exit For
        Else
            If arr(g) <> vbNullString Then
                If arr(g) = arr(g + c) Then
                    arr(g + c) = vbNullString
                    c = c + 1
                    g = g - 1
                Else
                    c = 1
                End If
            End If
        End If
    Next g
    c = 0
    For g = 0 To coun
        If arr(g) = vbNullString Then
        Else
            arr(c) = arr(g)
            c = c + 1
        End If
    Next g
    ReDim Preserve arr(c - 1)
End Sub

Private Sub TriQuickSortString2(ByRef sArray() As String, ByVal iSplit As Long, ByVal iMin As Long, ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim sTemp As String
   
   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   
   If (iMax - iMin) > iSplit Then
      i = (iMax + iMin) / 2
      
      If sArray(iMin) > sArray(i) Then SwapStrings sArray(iMin), sArray(i)
      If sArray(iMin) > sArray(iMax) Then SwapStrings sArray(iMin), sArray(iMax)
      If sArray(i) > sArray(iMax) Then SwapStrings sArray(i), sArray(iMax)
      
      j = iMax - 1
      SwapStrings sArray(i), sArray(j)
      i = iMin
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(j)), 4 ' sTemp = sArray(j)
      
      Do
         Do
            i = i + 1
         Loop While sArray(i) < sTemp
         
         Do
            j = j - 1
         Loop While sArray(j) > sTemp
         
         If j < i Then Exit Do
         SwapStrings sArray(i), sArray(j)
      Loop
      
      SwapStrings sArray(i), sArray(iMax - 1)
      
      TriQuickSortString2 sArray, iSplit, iMin, j
      TriQuickSortString2 sArray, iSplit, i + 1, iMax
   End If
   
   ' clear temp var (sTemp)
   i = 0
   CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
End Sub

Public Sub TriQuickSortString(ByRef sArray() As String, Optional ByVal SortOrder As SortOrder = SortAscending)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim i       As Long
   Dim j       As Long
   Dim sTemp   As String
   
   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   
   ' *NOTE*  the value 4 is VERY important here !!!
   ' DO NOT CHANGE 4 FOR A LOWER VALUE !!!
   TriQuickSortString2 sArray, 4, iLBound, iUBound
   InsertionSortString sArray, iLBound, iUBound
   
   If SortOrder = SortDescending Then ReverseStringArray sArray
End Sub

Private Sub InsertionSortString(ByRef sArray() As String, ByVal iMin As Long, ByVal iMax As Long)
   Dim i     As Long
   Dim j     As Long
   Dim sTemp As String
   
   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   
   For i = iMin + 1 To iMax
      CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(i)), 4 ' sTemp = sArray(i)
      j = i
      
      Do While j > iMin
         If sArray(j - 1) <= sTemp Then Exit Do

         CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sArray(j - 1)), 4 ' sArray(j) = sArray(j - 1)
         j = j - 1
      Loop
      
      CopyMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sTemp), 4
      ' sArray(j) = sTemp
   Next i
   
   ' clear temp var (sTemp)
   i = 0
   CopyMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
End Sub

Public Sub ReverseStringArray(ByRef sArray() As String)
   Dim iLBound As Long
   Dim iUBound As Long

   iLBound = LBound(sArray)
   iUBound = UBound(sArray)
   
   While iLBound < iUBound
      SwapStrings sArray(iLBound), sArray(iUBound)
   
      iLBound = iLBound + 1
      iUBound = iUBound - 1
   Wend
End Sub

Private Sub SwapStrings(ByRef s1 As String, ByRef s2 As String)
   Dim i As Long

   ' StrPtr() returns 0 (NULL) if string is not initialized
   ' But StrPtr() is 5% faster than using CopyMemory, so I used that workaround, which is safe and fast.
   i = StrPtr(s1)
   If i = 0 Then CopyMemory ByVal VarPtr(i), ByVal VarPtr(s1), 4

   CopyMemory ByVal VarPtr(s1), ByVal VarPtr(s2), 4
   CopyMemory ByVal VarPtr(s2), i, 4
End Sub
