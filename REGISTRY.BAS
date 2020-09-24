Attribute VB_Name = "Registry_Management"
' Registry management

'
'
' Handling of registry data types is limited.
' Only the simple string & longint types are supported as yet
'
'
'

Option Explicit




' Get the format into ".XXX"
'
Private Function FormatFileExtensionForRegistry$(pExtension$)

Dim lExtension$


  lExtension$ = pExtension$
  If (left$(lExtension$, 1) <> ".") Then lExtension$ = "." & lExtension$
  lExtension$ = LCase$(left$(lExtension$, 4))
  
  FormatFileExtensionForRegistry$ = lExtension$
End Function

' A function in the spirit of GetPrivateProfileString(),
' but using the Registry.
'
' Registry usage should comply with the recommendations
' in the Win '95 Interface Guidelines
'
Public Function GetPrivateRegistryString$(ByVal pKey As Long, pSectionName As String, pEntryName As String, pDefault As String)
Attribute GetPrivateRegistryString.VB_Description = "Simple replacement for GetPrivateProfileString()"

  GetPrivateRegistryString = GetRegistryString(pKey, GetRegistryKeyNameForApp$() & "\" & pSectionName, pEntryName, pDefault)
End Function



' Find the "Server" or "Open" app
'
' pFindServer is true for the server, false for the Open command
'
' NB - pExe$ will probably include command line parameters of "%1" substitution parameters
' particularly for the "Open" item
'
' Where there is more than one document type registered to this extension,
' the default type will be returned
'
Function GetRegistryAppByExtension(pExtension$, pExe$, ByVal pFindServer As Boolean) As Long

Dim lExtension$

Dim lRetVal As Long
Dim lhExtensionKey As Long, lhOpenKey As Long
Dim lProgID$


  lRetVal = ERROR_SUCCESS
  pExe$ = ""
  
  ' Get the format into ".XXX"
  lExtension$ = FormatFileExtensionForRegistry$(pExtension$)
  
    
  ' Use the extension to find ProgID, then try to find the default property of
  ' either HKEY_CLASSES_ROOT\ProgID\shell\open\command
  ' or HKEY_CLASSES_ROOT\ProgID\protocol\StdFileEditing\server
  lRetVal = RegOpenKey(HKEY_CLASSES_ROOT, lExtension$, KEY_READ, lhExtensionKey)
  If (lRetVal = ERROR_SUCCESS) Then lRetVal = RegQueryValue(lhExtensionKey, "", lProgID$)
  If (lRetVal = ERROR_SUCCESS) Then
    If pFindServer Then
      lRetVal = RegOpenKey(HKEY_CLASSES_ROOT, lProgID$ & "\protocol\StdFileEditing\server", KEY_READ, lhOpenKey)
    Else
      lRetVal = RegOpenKey(HKEY_CLASSES_ROOT, lProgID$ & "\shell\open\command", KEY_READ, lhOpenKey)
    End If
  End If
  If (lRetVal = ERROR_SUCCESS) Then lRetVal = RegQueryValue(lhOpenKey, "", pExe$)
  
  GetRegistryAppByExtension = lRetVal
End Function

' A function in the spirit of GetPrivateProfileString() et al,
' but using the Registry.
'
' Registry usage should comply with the recommendations
' in the Win '95 Interface Guidelines
'
Public Function GetRegistryLong(ByVal pKey As Long, pKeyName As String, pValueName As String, pDefault As Long) As Long
  
Dim llRetVal As Long
Dim lhOurKey As Long
Dim pValue As Variant

  
  GetRegistryLong = pDefault
  
  llRetVal = RegOpenKey(pKey, pKeyName, KEY_READ, lhOurKey)
  If (llRetVal = ERROR_SUCCESS) Then llRetVal = RegQueryValue(lhOurKey, pValueName, pValue)
  If (llRetVal = ERROR_SUCCESS) Then
    ' KLUDGE - This is an implicit assumption about the type of data returned
    ' Need to clarify how RegQueryValue works under these circumstances
    GetRegistryLong = pValue
  End If
  
End Function

' Read a string value from the registry
'
' Registry usage should comply with the recommendations
' in the Win '95 Interface Guidelines
'
Public Function GetRegistryString$(ByVal pKey As Long, pKeyName As String, pValueName As String, pDefault As String)
Attribute GetRegistryString.VB_Description = "Simple replacement for GetProfileString()"
  
Dim llRetVal As Long
Dim lhOurKey As Long
Dim pValue As Variant

  
  GetRegistryString$ = pDefault
  
  llRetVal = RegOpenKey(pKey, pKeyName, KEY_READ, lhOurKey)
  If (llRetVal = ERROR_SUCCESS) Then llRetVal = RegQueryValue(lhOurKey, pValueName, pValue)
  If (llRetVal = ERROR_SUCCESS) Then GetRegistryString$ = pValue
  
End Function


' Loads a form's position from the registry
' Complementary to SaveFormPosition
'
' To use this, the form must have its Tag property set
'
' Errors are returned, but most may be ignored
'
' Positions are stored under:
' HKEY_CURRENT_USER\Software\App.CompanyName\App.Title\pForm.Tag\X, Y, cX, cY
'
Public Function LoadFormPosition(pForm As Form) As Long
Attribute LoadFormPosition.VB_Description = "Makes a form's position ""sticky"""

Dim llRetVal As Long
Dim lKeyName$

Dim X As Long, Y As Long, cx As Long, cy As Long


  llRetVal = ERROR_SUCCESS
  
  ' Tag property must be set
  If (Len(pForm.Tag) = 0) Then llRetVal = ERROR_BADKEY
  
  If (llRetVal = ERROR_SUCCESS) Then
    lKeyName$ = GetRegistryKeyNameForApp$() & "\Form Positions\" & pForm.Tag
    pForm.WindowState = GetRegistryLong(HKEY_CURRENT_USER, lKeyName$, "WindowState", pForm.WindowState)
  End If
  If (llRetVal = ERROR_SUCCESS) Then X = GetRegistryLong(HKEY_CURRENT_USER, lKeyName$, "X", pForm.left)
  If (llRetVal = ERROR_SUCCESS) Then Y = GetRegistryLong(HKEY_CURRENT_USER, lKeyName$, "Y", pForm.top)
  If (llRetVal = ERROR_SUCCESS) Then cx = GetRegistryLong(HKEY_CURRENT_USER, lKeyName$, "cX", pForm.Width)
  If (llRetVal = ERROR_SUCCESS) Then cy = GetRegistryLong(HKEY_CURRENT_USER, lKeyName$, "cY", pForm.Height)
  
  If (llRetVal = ERROR_SUCCESS) Then
    cx = Minimum(Maximum(cx, 32), Screen.Width)
    cy = Minimum(Maximum(cy, 32), Screen.Height)
    X = Minimum(Maximum(X, 0), (Screen.Width - cx))
    Y = Minimum(Maximum(Y, 0), (Screen.Height - cy))
  End If
  
  If (llRetVal = ERROR_SUCCESS) Then
    pForm.Width = cx
    pForm.Height = cy
    pForm.left = X
    pForm.top = Y
  End If
  
  LoadFormPosition = llRetVal
End Function

' A function in the spirit of PutPrivateProfileString(),
' but using the Registry.
'
' Returns usual Windows error codes
'   =0 (ERROR_SUCCESS) is OK
'  <>0 is a failure
'
' Registry usage should comply with the recommendations
' in the Win '95 Interface Guidelines
'
Public Function PutPrivateRegistryString(ByVal pKey As Long, pSectionName As String, pEntryName As String, pValue As String) As Long
Attribute PutPrivateRegistryString.VB_Description = "Simple replacement for PutPrivateProfileString()"
  
  PutPrivateRegistryString = PutRegistryString(pKey, GetRegistryKeyNameForApp$() & "\" & pSectionName, pEntryName, pValue)
End Function


' A function in the spirit of PutPrivateProfileString(),
' but using the Registry.
'
' Returns usual Windows error codes
'   =0 (ERROR_SUCCESS) is OK
'  <>0 is a failure
'
' Registry usage should comply with the recommendations
' in the Win '95 Interface Guidelines
'
Public Function PutPrivateRegistryLong(ByVal pKey As Long, pSectionName As String, pEntryName As String, pValue As Long) As Long
  
  PutPrivateRegistryLong = PutRegistryLong(pKey, GetRegistryKeyNameForApp$() & "\" & pSectionName, pEntryName, pValue)
End Function

' A function in the spirit of PutPrivateProfileString() et al,
' but using the Registry.
'
' Returns usual Windows error codes
'   =0 (ERROR_SUCCESS) is OK
'  <>0 is a failure
'
' Registry usage should comply with the recommendations
' in the Win '95 Interface Guidelines
'
Public Function PutRegistryLong(ByVal pKey As Long, pKeyName As String, pValueName As String, pLongValue As Long) As Long

Dim llRetVal As Long
Dim lhOurKey As Long


  ' Construct new keys as necessary
  llRetVal = RegCreateKey(pKey, pKeyName, lhOurKey)
  
    ' Insert our value
  If (llRetVal = ERROR_SUCCESS) Then llRetVal = RegSetValue(lhOurKey, pValueName, REG_DWORD, pLongValue)
  
  ' Close the key. This will cause Windows to flush it, but not immediately
  If (llRetVal = ERROR_SUCCESS) Then llRetVal = RegCloseKey(lhOurKey)
  PutRegistryLong = llRetVal
End Function

' Save an "application path" for an application
'
'   Parameters:
' pAppName$   Identifier - application's EXE name, "MyApp.EXE"
'
' pAppEXE$    Full path & filename to the executable
'
' pAppPath$   Path entries to be added to the PATH when the executable is run
'             A suffix of ";" will be appended if necessary
'
' If either of the entries is an empty string, that value will not be written to the registry
'
' If both entries are empty strings, the key will be removed
'
Public Function SaveRegistryAppPath(pAppName$, pAppEXE$, pAppPath$) As Long
Attribute SaveRegistryAppPath.VB_Description = "Register an application path"

Dim lRetVal As Long
Dim lKey$


  lKey$ = "Software\Microsoft\Windows\CurrentVersion\App Paths\" & pAppName$
  
  
  lRetVal = ERROR_SUCCESS
  
'  If (lRetVal = ERROR_SUCCESS) Then
    If (Len(pAppEXE$) > 0) Then
      lRetVal = PutRegistryString(HKEY_LOCAL_MACHINE, lKey$, "", pAppEXE$)
    Else
      lRetVal = RegDeleteValue(HKEY_LOCAL_MACHINE, lKey$, "")
    End If
'  End If
  
  If (lRetVal = ERROR_SUCCESS) Then
    If (Len(pAppPath$) > 0) Then
      lRetVal = PutRegistryString(HKEY_LOCAL_MACHINE, lKey$, "Path", pAppPath$ & IIf((right$(pAppPath$, 1) <> ";"), ";", ""))
    Else
      lRetVal = RegDeleteValue(HKEY_LOCAL_MACHINE, lKey$, "Path")
    End If
  End If
  
  If (lRetVal = ERROR_SUCCESS) Then
    If (Len(pAppEXE$) = 0) And (Len(pAppPath$) = 0) Then
      lRetVal = RegDeleteKey(HKEY_LOCAL_MACHINE, lKey$)
    End If
  End If
  
  SaveRegistryAppPath = lRetVal
End Function

' Save an association for a file extension,
' i.e. the application used to open files of this type
'
' If either pOpenCommand$ or pIconPath$ are empty,
' that section will be ignored
'
Public Function SaveRegistryAssociation(pExtension$, pProgID$, _
  pOpenCommand$, _
  pIconPath$, pIconNumber%) As Long

Dim lExtension$
Dim lRetVal As Long
Dim lhKey As Long, lhSubKey As Long

  
  lRetVal = ERROR_SUCCESS

  lExtension$ = FormatFileExtensionForRegistry$(pExtension$)
  
  ' Check the extension & ProgID for validity
  If Not ((Len(lExtension$) > 0) And (Len(pProgID$) > 0)) Then
    lRetVal = ERROR_BADKEY
  End If
  
  ' Associate the extension with a ProgID
  If (lRetVal = ERROR_SUCCESS) Then
    lRetVal = RegCreateKey(HKEY_CLASSES_ROOT, lExtension$, lhKey)
    If (lRetVal = ERROR_SUCCESS) Then lRetVal = RegSetValue(lhKey, "", REG_SZ, pProgID$)
  End If
  
  ' Save attributes for the ProgID
  If (lRetVal = ERROR_SUCCESS) Then
    lRetVal = RegCreateKey(HKEY_CLASSES_ROOT, pProgID$, lhKey)
    
    ' Open
    If (lRetVal = ERROR_SUCCESS) And (Len(pOpenCommand$) > 0) Then
      lRetVal = RegCreateKey(lhKey, "Shell\Open\Command", lhSubKey)
      If (lRetVal = ERROR_SUCCESS) Then lRetVal = RegSetValue(lhSubKey, "", REG_SZ, pOpenCommand$)
    End If
    
    ' Icon
    If (lRetVal = ERROR_SUCCESS) And (Len(pIconPath$) > 0) Then
      lRetVal = RegCreateKey(lhKey, "DefaultIcon", lhSubKey)
      If (lRetVal = ERROR_SUCCESS) Then lRetVal = RegSetValue(lhSubKey, "", REG_SZ, pIconPath$ & "," & CStr(pIconNumber%))
    End If
  End If

  SaveRegistryAssociation = lRetVal
End Function


' Store a setting in the Registry,
' such that the current application will be NOT be re-started
'
Public Function SaveRegistryDontRunAppAfterRestart() As Long
Attribute SaveRegistryDontRunAppAfterRestart.VB_Description = "Cancels SaveRegistryRunAppAfterRestart()"
  
  SaveRegistryDontRunAppAfterRestart = SaveRegistryRunAtStartup(App.Title, "", True, False)
End Function


' Store a setting in the Registry,
' such that the current application will be re-started
'
Public Function SaveRegistryRunAppAfterRestart(pParameters$) As Long
Attribute SaveRegistryRunAppAfterRestart.VB_Description = "Sets current application to restart automatically when Windows is restarted"
  
  SaveRegistryRunAppAfterRestart = SaveRegistryRunAtStartup(App.Title, App.Path & "\" & App.EXEName & ".EXE " & pParameters$, True, False)
End Function

' Save a setting in the Run or RunOnce section of the Registry,
' such that an executable will be run at next Windows startup
'
' See Reference [1], pp239-240
'
'
' pValueName is descriptive, but generally ignored by Windows
' It must be unique, so may need to have a numeric suffix (or similar) appended to it when needed
'
' pValue is the path to the executable that will be run
' when Windows next starts, including a full path and any parameters
'
' pRunOnce is True for executables that should only be run on the next startup
' False when they should be run every time Windows starts
'
' pService will usually be false. It is only true for executables
' that are run before Windows starts (i.e. system functions such as virus scanners)
'
Public Function SaveRegistryRunAtStartup(pValueName$, pValue$, _
    ByVal pRunOnce As Boolean, pService As Boolean) As Long
Attribute SaveRegistryRunAtStartup.VB_Description = "Sets an application to be run when Windows restarts"

Dim lhKey As Long, lKey$


  lhKey = IIf(pService, HKEY_LOCAL_MACHINE, HKEY_CURRENT_USER)
  lKey$ = "Software\Microsoft\Windows\CurrentVersion\" & IIf(pRunOnce, "RunOnce", "Run")
  
  SaveRegistryRunAtStartup = PutRegistryString(lhKey, lKey$, pValueName$, pValue$)
End Function



' Store a string value anywhere in the registry
' The key chain will be built if necessary
'
' Returns usual Windows error codes
'   =0 (ERROR_SUCCESS) is OK
'  <>0 is a failure
'
' Registry usage should comply with the recommendations
' in the Win '95 Interface Guidelines
'
Public Function PutRegistryString(ByVal pKey As Long, pKeyName As String, pValueName As String, pStringValue As String) As Long
Attribute PutRegistryString.VB_Description = "Simple replacement for PutProfileString()"

Dim llRetVal As Long
Dim lhOurKey As Long


  ' Construct new keys as necessary
  llRetVal = RegCreateKey(pKey, pKeyName, lhOurKey)
  
    ' Insert our value
  If (llRetVal = ERROR_SUCCESS) Then llRetVal = RegSetValue(lhOurKey, pValueName, REG_SZ, pStringValue$)
  
  ' Close the key. This will cause Windows to flush it, but not immediately
  If (llRetVal = ERROR_SUCCESS) Then llRetVal = RegCloseKey(lhOurKey)
  PutRegistryString = llRetVal
End Function



' Returns the suggested key name for this application's setting
'
' i.e. "Software\Codesmiths\RegistryProgram"
'
Public Function GetRegistryKeyNameForApp$()
Attribute GetRegistryKeyNameForApp.VB_Description = "Returns a suggested Registry sub-key for this application"

Dim lSniff$, lFullLen%, lChopLen%

Dim lCompanyName$
Dim lAppName$
Dim lAppVersion$
Dim lWork$


  ' Company name
  lCompanyName$ = App.CompanyName
  
  ' Chop off any trailing "Ltd." (which must have a space before it)
  lFullLen% = Len(lCompanyName$)
  lChopLen% = 0
  If (right$(lCompanyName$, 1) = ".") Then lChopLen% = 1
  lSniff$ = UCase$(Mid$(lCompanyName$, lFullLen% - (lChopLen% + 2), 3))
  If (lSniff$ = " CO") Then
    lChopLen% = lChopLen% + 3
  Else
    lSniff$ = UCase$(Mid$(lCompanyName$, lFullLen% - (lChopLen% + 3), 4))
    If (lSniff$ = " LTD") Or (lSniff$ = " INC") Or (lSniff$ = " PLC") Then
      lChopLen% = lChopLen% + 4
    End If
  End If
  ' A special for "XXX Group Ltd"
  If (UCase$(Mid$(lCompanyName$, lFullLen% - (lChopLen% + 5), 6)) = " GROUP") Then lChopLen% = lChopLen% + 6
  If (lChopLen% > 0) Then lCompanyName$ = left$(lCompanyName$, Len(lCompanyName$) - lChopLen%)
  
  ' Application name
  lAppName$ = App.Title
  
  ' Version Number
  ' According to the Microsoft references ([1], p238)
  ' Our key should include a version number
  ' Common practice is to either ignore this, or to store everything under a key "CurrentVersion"
  ' First option:
  lAppVersion$ = ""
  ' Second option:
'  lAppVersion$ = "CurrentVersion"
  ' Third option:
'  lAppVersion$ = CStr(App.Major) & "." & CStr(App.Minor)


  ' Assemble all the pieces
  lWork$ = "Software"
  If (Len(lCompanyName$) > 0) Then lWork$ = lWork$ & "\" & lCompanyName$
  If (Len(lAppName$) > 0) Then lWork$ = lWork$ & "\" & lAppName$
  If (Len(lAppVersion$) > 0) Then lWork$ = lWork$ & "\" & lAppVersion$
  GetRegistryKeyNameForApp$ = lWork$
End Function

' Saves a form's position in the registry
' Complementary to LoadFormPosition
'
' To use this, the form must have its Tag property set
'
' Errors are returned, but most may be ignored
'
' Positions are stored under:
' HKEY_CURRENT_USER\Software\App.CompanyName\App.Title\pForm.Tag\X, Y, cX, cY
'
Public Function SaveFormPosition(pForm As Form) As Long
Attribute SaveFormPosition.VB_Description = "Makes a form's position ""sticky"""

Dim llRetVal As Long
Dim lKeyName$


  llRetVal = ERROR_SUCCESS
  
  ' Tag property must be set
  If (Len(pForm.Tag) = 0) Then llRetVal = ERROR_BADKEY
  
  If (llRetVal = ERROR_SUCCESS) Then
    lKeyName$ = GetRegistryKeyNameForApp$() & "\Form Positions\" & pForm.Tag
    llRetVal = PutRegistryLong(HKEY_CURRENT_USER, lKeyName$, "WindowState", pForm.WindowState)
  End If
  If (llRetVal = ERROR_SUCCESS) Then llRetVal = PutRegistryLong(HKEY_CURRENT_USER, lKeyName$, "X", pForm.left)
  If (llRetVal = ERROR_SUCCESS) Then llRetVal = PutRegistryLong(HKEY_CURRENT_USER, lKeyName$, "Y", pForm.top)
  If (llRetVal = ERROR_SUCCESS) Then llRetVal = PutRegistryLong(HKEY_CURRENT_USER, lKeyName$, "cX", pForm.Width)
  If (llRetVal = ERROR_SUCCESS) Then llRetVal = PutRegistryLong(HKEY_CURRENT_USER, lKeyName$, "cY", pForm.Height)
  
  SaveFormPosition = llRetVal
End Function




