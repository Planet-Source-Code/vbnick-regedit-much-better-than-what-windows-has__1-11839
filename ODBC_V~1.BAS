Attribute VB_Name = "ODBC_Viewing"
' ODBC Viewing

'   Notes:
' These wrapper functions are direct functional wrappers for the existing API functions
' Their purpose is to make the data types easier to work with under VB
' Please don't use the API functions directly, as this is a perennial source of obscure VB bugs
'
' When extending this module, don't place any sophisticated code in here
' - use Registry.BAS instead for higher level functions.
'
' This code makes no attempt to support multiple platforms.
' It's enough work to get it running under plain '95 for starters !

'
Option Explicit


' Display ODBC data sources and their main details
'
' Works by displaying the hive of registry entries from
' H_KEY_CURRENT_USER\Sotware\ODBC\ODBC.INI\
'
Public Sub ShowODBC(ptvwODBC As TreeView)

Dim hKey As Long
Dim hSubKey As Long
Dim dwValueIndex As Long

' Fragments of keys
Dim lpSubkeyName As String
Dim lpValueName As String
Dim lpValue As String

' Entire keys
Dim lSubKey$

Dim llRetVal As Long

Dim lpType As Long

Dim lcolDataSources As New Collection
Dim lobjDataSource
Dim nodX As Node


  ptvwODBC.Nodes.Clear
  

  ' get a collection of the ODBC DSNs
  ListODBCDataSources lcolDataSources
  
  ' Use the collection to find the data sources
  For Each lobjDataSource In lcolDataSources

    lpSubkeyName = lobjDataSource
    lSubKey$ = mODBCkeyName & lobjDataSource
    
    Set nodX = ptvwODBC.Nodes.Add(, , lpSubkeyName, lpSubkeyName, tvwTextOnly)
    
    llRetVal = RegOpenKey(HKEY_CURRENT_USER, lSubKey$, KEY_READ, hSubKey)
    If (llRetVal = ERROR_SUCCESS) Then
      ' Loop through all values for this key
      dwValueIndex = 0
      Do While (llRetVal = ERROR_SUCCESS)
        llRetVal = RegEnumValue(hSubKey, dwValueIndex, lpValueName, lpType, lpValue)
        
        If (llRetVal = ERROR_SUCCESS) Then
          ' Filter out interesting values
          Select Case lpValueName
          Case "Server", "Database"
            Set nodX = ptvwODBC.Nodes.Add(lpSubkeyName, tvwChild, lSubKey$ & lpValueName, lpValueName & " = """ & lpValue & """", tvwTextOnly)
          Case "Description"
            If (Len(lpValue) > 0) Then Set nodX = ptvwODBC.Nodes.Add(lpSubkeyName, tvwChild, lSubKey$ & lpValueName, lpValueName & " = """ & lpValue & """", tvwTextOnly)
          Case Else
          End Select
          dwValueIndex = dwValueIndex + 1
        End If
      Loop
      nodX.EnsureVisible
    Else
      ' Couldn't open key
    End If
  Next lobjDataSource
  
  
  ' Make the first node visible
  If (ptvwODBC.Nodes.Count > 0) Then ptvwODBC.Nodes(1).EnsureVisible
End Sub



