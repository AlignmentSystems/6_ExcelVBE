Attribute VB_Name = "numberStation"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       28th March 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
Const clngNumerator As Long = 100
Const clngDenominator As Long = 0

Public Sub TestNumbers()
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       28th March 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Dim lngBrokenAnswer As Long

On Error GoTo ErrHandler

Debug.Print "TestNumbers: Testing error handling"

If 1 = 2 Then
    Debug.Print "testing error handling"
    Debug.Print "testing error handling"
End If

lngBrokenAnswer = clngNumerator / clngDenominator

If 2 = 1 Then
    Debug.Print "testing error handling"
    Debug.Print "testing error handling"
End If

Exit Sub
ErrHandler:

Debug.Print "[TestNumbers]Number=" & Err.Number & VBA.vbCrLf & "Description=" & Err.Description
Err.Clear
On Error GoTo 0




End Sub



Public Sub TestNumbers2()
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       28th March 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
Dim lngBrokenAnswer As Long

On Error GoTo ErrHandler

1   Debug.Print "TestNumbers2: Testing error handling"

2   If 1 = 2 Then
3       Debug.Print "testing error handling"
4       Debug.Print "testing error handling"
5   End If

6   lngBrokenAnswer = clngNumerator / clngDenominator

7   If 2 = 1 Then
8       Debug.Print "testing error handling"
9       Debug.Print "testing error handling"
10  End If

Exit Sub
ErrHandler:

Debug.Print "[TestNumbers2]Number=" & Err.Number & VBA.vbCrLf & "Description=" & Err.Description & VBA.vbCrLf & "LineOfCode=" & Erl
Err.Clear
On Error GoTo 0




End Sub
