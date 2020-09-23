Attribute VB_Name = "modRoutines"
Option Explicit
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

' Sort function created by Mr. Lord & tweaked to handle various dimensioned arrays

Public Sub TriQuickSortLong(ByRef iArray() As Long, iMemberID As Byte)
   Dim iLBound As Long
   Dim iUBound As Long
   Dim I       As Long
   Dim J       As Long
   Dim iTemp   As Long
   
   iLBound = LBound(iArray, 2)
   iUBound = UBound(iArray, 2)
   
   ' *NOTE*  the value 4 is VERY important here !!!
   ' DO NOT CHANGE 4 FOR A LOWER VALUE !!!
   TriQuickSortLong2 iArray, 4, iLBound, iUBound, iMemberID
   InsertionSortLong iArray, iLBound, iUBound, iMemberID
   
End Sub

Private Sub TriQuickSortLong2(ByRef iArray() As Long, ByVal iSplit As Long, ByVal iMin As Long, ByVal iMax As Long, ByVal iMemberID As Byte)
   Dim I     As Long
   Dim J     As Long
   Dim iTemp() As Long
   Dim k As Long

   ReDim iTemp(LBound(iArray, 1) To UBound(iArray, 1))

   If (iMax - iMin) > iSplit Then
      I = (iMax + iMin) / 2
      
      If iArray(1, iMin) > iArray(1, I) Then SwapLongs iArray(), iMin, I
      If iArray(1, iMin) > iArray(1, iMax) Then SwapLongs iArray(), iMin, iMax
      If iArray(1, I) > iArray(1, iMax) Then SwapLongs iArray(), I, iMax
      
      J = iMax - 1
      SwapLongs iArray(), I, J
      
      I = iMin
      For k = LBound(iTemp) To UBound(iTemp)
        iTemp(k) = iArray(k, J)
      Next
        
      
      Do
         Do
            I = I + 1
         Loop While iArray(iMemberID, I) < iTemp(iMemberID)
         
         Do
            J = J - 1
         Loop While iArray(iMemberID, J) > iTemp(iMemberID) And J > 0
' Note: Only logic modification I made was to add the "J > 0" above
' In certain cases, J would fall below zero & routine would crash
'//LaVolpe
         If J < I Then Exit Do
         
         SwapLongs iArray(), I, J
      Loop
      
      SwapLongs iArray(), I, iMax - 1
      
      TriQuickSortLong2 iArray, iSplit, iMin, J, iMemberID
      TriQuickSortLong2 iArray, iSplit, I + 1, iMax, iMemberID
   End If
End Sub

   Private Sub SwapLongs(ByRef iArray, Index1 As Long, Index2 As Long)
   Dim I As Long, J As Long
   For J = LBound(iArray, 1) To UBound(iArray, 1)
    I = iArray(J, Index1)
    iArray(J, Index1) = iArray(J, Index2)
    iArray(J, Index2) = I
   Next
   End Sub

Private Sub InsertionSortLong(ByRef iArray() As Long, ByVal iMin As Long, ByVal iMax As Long, iMemberID As Byte)
   Dim I     As Long
   Dim J     As Long
   Dim iTemp() As Long
   Dim k As Long
   ReDim iTemp(LBound(iArray, 1) To UBound(iArray, 1))
   
   For I = iMin + 1 To iMax
      For k = LBound(iTemp) To UBound(iTemp)
        iTemp(k) = iArray(k, I)
      Next
      J = I
      
      Do While J > iMin
         If iArray(iMemberID, J - 1) <= iTemp(iMemberID) Then Exit Do

         For k = LBound(iTemp) To UBound(iTemp)
            iArray(k, J) = iArray(k, J - 1)
         Next
         
         J = J - 1
      Loop
      
        For k = LBound(iTemp) To UBound(iTemp)
            iArray(k, J) = iTemp(k)
        Next
   Next I
End Sub

Public Function HiWord(LongIn As Long) As Integer
  Call CopyMemory(HiWord, ByVal VarPtr(LongIn) + 2, 2)
End Function

Public Function LoWord(LongIn As Long) As Integer
  Call CopyMemory(LoWord, ByVal VarPtr(LongIn), 2)
End Function

Public Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
  MakeLong = CLng(LoWord)
  Call CopyMemory(ByVal VarPtr(MakeLong) + 2, HiWord, 2)
End Function

Public Sub MoveObject(hwnd As Long)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub
