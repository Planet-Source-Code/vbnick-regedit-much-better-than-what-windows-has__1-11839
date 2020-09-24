Attribute VB_Name = "General_Library"
' General library routines
'
' These routines are a subset of the normal Gen_Lib.BAS,
' solely for demonstration of the Registry routines
'

Option Explicit

'    simple sub to centre any form on screen
'
Public Sub CentreForm(pfrmCentre As Form)

  With pfrmCentre
    .top = ((Screen.Height * 0.95) - .Height) / 2
    .left = (Screen.Width - .Width) / 2
    End With
End Sub



Function Maximum(ByVal a, ByVal B)
  If (a >= B) Then
    Maximum = a
  Else
    Maximum = B
  End If
End Function


Function Minimum(ByVal a, ByVal B)
  If (a <= B) Then
    Minimum = a
  Else
    Minimum = B
  End If
End Function



' Remove all entries from a collection
'
Public Sub EmptyCollection(pCol As Collection)

  Do While pCol.Count > 0
    pCol.Remove 1
  Loop
End Sub



