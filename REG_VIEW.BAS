Attribute VB_Name = "Registry_Viewing"
' Registry viewing
'   Notes:
'
'
'   Bugs:
'
'
Option Explicit










' Loads a sub-tree into the TreeView control pTree
'
'   Algorithm:
' The traditional recursive descent, but flattened into a non-recursive
' implementation (easier to implement with VB).
' Arrays are used to simulate the call stack of a recursive implementation
'
' A state machine is used to control execution of the algorithm, and the
' switch between logging values and logging keys
'
'   Enhancements:
' Show icons to distinguish the key nodes from values
'
'   Notes:
' Values are prefixed with a "v" before being added to the node collection,
' for otherwise there will be a RTE if a key & a value with identical names are encountered
' It would be neater if this were done with a trailing "\" instead
'
Public Sub ViewRegistrySubTree(ByVal pRootKey As Long, pInitialKey$, pTree As TreeView)



Dim hKey As Long
Dim hSubKey As Long

' Fragments of keys
Dim lpSubkeyName As String
Dim lpValueName As String
Dim lpValue As String

' Entire keys
Dim lRootKey$



Dim llRetVal As Long

Dim lpType As Long

Dim lcolDataSources As New Collection
Dim lobjDataSource
Dim nodX As Node

Const lMAX_DESCENT = 7

' Arrays for doing the recursive descent
Dim lDescentLevel%
Dim lLastDescentLevel%
Dim lDescendingIndex(0 To lMAX_DESCENT) As Long
Dim lDescendingKey$(0 To lMAX_DESCENT)
Dim lDescendingKeyLong$(0 To lMAX_DESCENT)
Dim lDescendinghKey(0 To lMAX_DESCENT) As Long
Dim lDescendingScanState%(0 To lMAX_DESCENT)
Dim lScanState%

Dim i%
Dim lWork$


Const lSTATE_OPEN_KEY = 1
Const lSTATE_ITERATE_VALUES = 2
Const lSTATE_ITERATE_KEYS = 3
Const lSTATE_PUSH_KEY = 4
Const lSTATE_POP_KEY = 5
Const lSTATE_FINISHED = 999

Dim lKeyCount As Long

Dim lMaxExpandLevel%


  On Error GoTo ViewRegistrySubTree_ERH
  pTree.Nodes.Clear
  DoEvents
  
  ' At first we'll expand 1,2 & >=4
  ' Later we'll turn this to just 1 & >=4
  lMaxExpandLevel% = 2
  
  lLastDescentLevel% = 0
  lDescendinghKey(lLastDescentLevel%) = pRootKey
  lDescendingKey$(0) = ""
  lDescendingKeyLong$(0) = "\"
  
  lDescentLevel% = lLastDescentLevel% + 1
  lDescendingKey$(lDescentLevel%) = pInitialKey$
  lDescendingKeyLong$(lDescentLevel%) = "\" & lDescendingKey$(lDescentLevel%)
  lDescendingScanState%(lDescentLevel%) = 0
  lDescendingIndex(lDescentLevel%) = 0
  
  lScanState% = lSTATE_OPEN_KEY
  lKeyCount = 0
  
  
  ' Loop
  Do While (llRetVal = ERROR_SUCCESS) And (lScanState% <> lSTATE_FINISHED) And (lDescentLevel% >= 0) And (lDescentLevel% <= lMAX_DESCENT)
  
    Select Case lScanState%
    
    Case lSTATE_OPEN_KEY
      llRetVal = RegOpenKey(lDescendinghKey(lLastDescentLevel%), lDescendingKey$(lDescentLevel%), KEY_READ, lDescendinghKey(lDescentLevel%))
      
      If llRetVal = ERROR_SUCCESS Then
        lKeyCount = lKeyCount + 1
        lDescendingKeyLong$(lDescentLevel%) = IIf((lDescendingKeyLong$(lLastDescentLevel%) = "\"), "", lDescendingKeyLong$(lLastDescentLevel%)) & "\" & lDescendingKey$(lDescentLevel%)
    
        If lDescentLevel% = 1 Then
'          Debug.Print "Adding First Key <>", "<"; lDescendingKeyLong$(lDescentLevel%); ">", "<"; lDescendingKey$(lDescentLevel%); ">"
          Set nodX = pTree.Nodes.Add(, , lDescendingKeyLong$(lDescentLevel%), lDescendingKey$(lDescentLevel%), tvwTextOnly)
        Else
'          Debug.Print "Adding Key <"; lDescendingKeyLong$(lLastDescentLevel%); ">", "<"; lDescendingKeyLong$(lDescentLevel%); ">", "<"; lDescendingKey$(lDescentLevel%); ">"
          Set nodX = pTree.Nodes.Add(lDescendingKeyLong$(lLastDescentLevel%), tvwChild, lDescendingKeyLong$(lDescentLevel%), lDescendingKey$(lDescentLevel%), tvwTextOnly)
        End If
        nodX.Expanded = (lDescentLevel% > 3) Or (lDescentLevel% <= lMaxExpandLevel%)
        If (pTree.Nodes.Count > 300) Then lMaxExpandLevel% = 2
        If (lDescentLevel% < 4) Or ((lKeyCount Mod 32) = 0) Then DoEvents
        
        lScanState% = lSTATE_ITERATE_VALUES
        lDescendingIndex(lDescentLevel%) = 0
      Else
        lScanState% = lSTATE_FINISHED
      End If
      
      
    Case lSTATE_ITERATE_VALUES
      llRetVal = RegEnumValue(lDescendinghKey(lDescentLevel%), lDescendingIndex(lDescentLevel%), lpValueName, lpType, lpValue)
    
      Select Case llRetVal
      Case ERROR_SUCCESS
        Select Case lpType
        Case 1
        
        Case Else
        End Select
'          Debug.Print "Adding Value <"; lDescendingKeyLong$(lDescentLevel%); ">", "<"; lDescendingKeyLong$(lDescentLevel%) & "\" & lpValueName; ">", "<"; lpValueName & "=" & lpValue; ">"
          Set nodX = pTree.Nodes.Add(lDescendingKeyLong$(lDescentLevel%), tvwChild, "v" & lDescendingKeyLong$(lDescentLevel%) & "\" & lpValueName, lpValueName & "=" & lpValue, tvwTextOnly) ' LoadPicture("D:\vb4-32\icons\writing\note10.ico")
          nodX.Expanded = True
        lDescendingIndex(lDescentLevel%) = lDescendingIndex(lDescentLevel%) + 1
      
      Case 259
        llRetVal = ERROR_SUCCESS
        lScanState% = lSTATE_ITERATE_KEYS
        lDescendingIndex(lDescentLevel%) = 0
      
      Case Else
        lScanState% = lSTATE_FINISHED
        
      End Select
    
    
    Case lSTATE_ITERATE_KEYS
      llRetVal = RegEnumKey(lDescendinghKey(lDescentLevel%), lDescendingIndex(lDescentLevel%), lpValueName)
    
      Select Case llRetVal
      Case ERROR_SUCCESS
        lScanState% = lSTATE_PUSH_KEY
      
      Case 259
        llRetVal = ERROR_SUCCESS
        lScanState% = lSTATE_POP_KEY
      
      Case Else
        lScanState% = lSTATE_FINISHED
        
      End Select
    
        
    Case lSTATE_PUSH_KEY
      ' Going down
      lLastDescentLevel% = lDescentLevel%
      lDescentLevel% = lDescentLevel% + 1
      
      lDescendingKey$(lDescentLevel%) = lpValueName
        
      If lDescentLevel% >= lMAX_DESCENT Then
        lScanState% = lSTATE_POP_KEY  ' Too far down
      Else
        lScanState% = lSTATE_OPEN_KEY
      End If
      
      
    Case lSTATE_POP_KEY
      ' Going up
      lDescentLevel% = lDescentLevel% - 1
      lLastDescentLevel% = lDescentLevel% - 1
      
      If lDescentLevel% >= 1 Then
        lDescendingIndex(lDescentLevel%) = lDescendingIndex(lDescentLevel%) + 1
        lScanState% = lSTATE_ITERATE_KEYS
      Else
        lScanState% = lSTATE_FINISHED
      End If
      
    Case Else
    End Select
  Loop
  
  
  ' Make the first node visible
  If (pTree.Nodes.Count > 0) Then pTree.Nodes(1).EnsureVisible
  
'  Debug.Print "Key count = "; lKeyCount
  
  
  Exit Sub
  
  
  
ViewRegistrySubTree_ERH:
  
Dim objNode
  
  Debug.Print
  
'  For Each objNode In pTree.Nodes
'    Debug.Print objNode.Index, "<"; objNode.Key; ">", "<"; objNode.Text; ">"
'  Next objNode
  Debug.Print "**************************************"
  Debug.Print Err, Err.Description
  Debug.Print "lDescentLevel% = "; lDescentLevel%
  Debug.Print "Key count = "; lKeyCount
  Debug.Print lDescendingKeyLong$(lDescentLevel%), "v" & lDescendingKeyLong$(lDescentLevel%) & "\" & lpValueName, lpValueName & "=" & lpValue
  Debug.Print
  
  Exit Sub
End Sub



