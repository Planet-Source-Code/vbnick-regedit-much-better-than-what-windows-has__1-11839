Attribute VB_Name = "Registry_API_Declares"
' API declarations and simple wrapper functions for the Registry
'
'
'
' JET 2 engine required, no known JET 3 features used
'
'
'   Modified:
' AD  4/ 9/95   Initial version
'

Option Explicit




#If Win16 Then
'  ERROR - This code does not work under 16 bit Windows
  ' The Registry does exist from W3.1 onwards,
  ' but only for use with OLE registration. This
  ' code does not attempt to support these functions
#End If

#If Win32 Then
#Else
'  ERROR - This code requires 32 bit Windows
#End If


' NB - We can't nest a #Const within a conditional compile block
' #If Win32 Then
  ' Define this true to use '95 style enhanced functions with access control (RegOpenKeyEx etc.)
  ' False uses the 32 bit version of the 16 bit functions
  #Const UseRegExFunctions = True
' #Else
  ' For 16 bits, these functions will simply be unavailable
'  #Const UseRegExFunctions = False
' #End If





'Registration functions
#If Win16 Then
' Old 16 bit functions
Public Declare Function RegCloseKey Lib "SHELL.DLL" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "SHELL.DLL" (ByVal hKey As Long, ByVal lpszSubKey As String, lphKey As Long) As Long
Public Declare Function RegDeleteKey Lib "SHELL.DLL" (ByVal hKey As Long, ByVal lpszSubKey As String) As Long
Public Declare Function RegEnumKey Lib "SHELL.DLL" (ByVal hKey As Long, ByVal iSubKey As Long, ByVal lpszBuffer As String, ByVal cbBuffer As Long) As Long
Public Declare Function RegOpenKey Lib "SHELL.DLL" (ByVal hKey As Long, ByVal lpszSubKey As String, lphkResult As Long) As Long
Public Declare Function RegQueryValue Lib "SHELL.DLL" (ByVal hKey As Long, ByVal lpszSubKey As String, ByVal lpszValue As String, lpcbValue As Long) As Long
Public Declare Function RegSetValue Lib "SHELL.DLL" (ByVal hKey As Long, ByVal lpszSubKey As String, ByVal fdwType As Long, ByVal lpszValue As String, ByVal dwLength As Long) As Long
#End If


#If Win32 Then
' 32 bit declares for Win32s

' NB - Don't use the supplied WIN32API.TXT instead. It has bugs in the translation
' of its function headers from C to VB.
' RegEnumValue & RegSetValue (et al) have lpData declared "lpData As Byte", when it needs to be "lpData As Any"
' if we are to manipulate pointers to arrays of bytes from VB


' These declares will be qualified as "Private" as each new wrapper function is implemented
' Exception - we won't wrap these two, as their headers are trivially compatible with VB
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegCreateKeyA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKeyA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValueA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyA Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
'Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValueA Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Private Declare Function RegOpenKeyA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValueA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


Public Declare Function GetEnvironmentStrings Lib "Kernel32" Alias "GetEnvironmentStringsA" () As String
Public Declare Function GetEnvironmentVariable Lib "Kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SetEnvironmentVariable Lib "Kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long

Public Declare Function ExpandEnvironmentStrings Lib "Kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nDstSize As Long) As Long
#End If



' Return codes from Registration functions
Public Const ERROR_SUCCESS = 0&

#If Win16 Then
' These constants are taken from 16-bit sources
' They appear to be incompatible with the 32-bit constants
Public Const ERROR_BADDB = 1&
Public Const ERROR_BADKEY = 2&
Public Const ERROR_CANTOPEN = 3&
Public Const ERROR_CANTREAD = 4&
Public Const ERROR_CANTWRITE = 5&
#End If

Public Const ERROR_OUTOFMEMORY = 6&
Public Const ERROR_INVALID_PARAMETER = 7&
Public Const ERROR_ACCESS_DENIED = 8&




#If Win32 Then
'   More data is available.
Public Const ERROR_MORE_DATA = 234 '  dderror


'   The configuration registry database is corrupt.
Public Const ERROR_BADDB = 1009&

'   The configuration registry key is invalid.
Public Const ERROR_BADKEY = 1010&

'   The configuration registry key could not be opened.
Public Const ERROR_CANTOPEN = 1011&

'   The configuration registry key could not be read.
Public Const ERROR_CANTREAD = 1012&

'   The configuration registry key could not be written.
Public Const ERROR_CANTWRITE = 1013&

'   One of the files in the Registry database had to be recovered
'   by use of a log or alternate copy.  The recovery was successful.
Public Const ERROR_REGISTRY_RECOVERED = 1014&

'   The Registry is corrupt. The structure of one of the files that contains
'   Registry data is corrupt, or the system's image of the file in memory
'   is corrupt, or the file could not be recovered because the alternate
'   copy or log was absent or corrupt.
Public Const ERROR_REGISTRY_CORRUPT = 1015&

'   An I/O operation initiated by the Registry failed unrecoverably.
'   The Registry could not read in, or write out, or flush, one of the files
'   that contain the system's image of the Registry.
Public Const ERROR_REGISTRY_IO_FAILED = 1016&

'   The system has attempted to load or restore a file into the Registry, but the
'   specified file is not in a Registry file format.
Public Const ERROR_NOT_REGISTRY_FILE = 1017&

'   Illegal operation attempted on a Registry key which has been marked for deletion.
Public Const ERROR_KEY_DELETED = 1018&

'   System could not allocate the required space in a Registry log.
Public Const ERROR_NO_LOG_SPACE = 1019&

'   Cannot create a symbolic link in a Registry key that already
'   has subkeys or values.
Public Const ERROR_KEY_HAS_CHILDREN = 1020&

'   Cannot create a stable subkey under a volatile parent key.
Public Const ERROR_CHILD_MUST_BE_VOLATILE = 1021&
#End If



#If Win32 Then
'   These are the generic rights.

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_ALL = &H10000000
#End If





'Registration constants


Public Const READ_CONTROL = &H20000
Public Const SYNCHRONIZE = &H100000

Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const STANDARD_RIGHTS_ALL = &H1F0000


'//
'// Registry Specific Access Rights.
'//

Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20

Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
                                  KEY_QUERY_VALUE Or _
                                  KEY_ENUMERATE_SUB_KEYS Or _
                                  KEY_NOTIFY) _
                                  And _
                                 (Not SYNCHRONIZE))


Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or _
                                  KEY_SET_VALUE Or _
                                  KEY_CREATE_SUB_KEY) _
                                  And _
                                 (Not SYNCHRONIZE))

Public Const KEY_EXECUTE = ((KEY_READ) _
                                  And _
                                 (Not SYNCHRONIZE))

Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or _
                                  KEY_QUERY_VALUE Or _
                                  KEY_SET_VALUE Or _
                                  KEY_CREATE_SUB_KEY Or _
                                  KEY_ENUMERATE_SUB_KEYS Or _
                                  KEY_NOTIFY Or _
                                  KEY_CREATE_LINK) _
                                  And _
                                 (Not SYNCHRONIZE))



Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005

Public Const MAX_PATH = 256&


'//
'//
'// Predefined Value Types.
'//

Public Const REG_NONE = 0                     '// No value type
Public Const REG_SZ = 1                       ' // Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2                ' // Unicode nul terminated string
                                              ' // (with environment variable references)
Public Const REG_BINARY = 3                   ' // Free form binary
Public Const REG_DWORD_LITTLE_ENDIAN = 4
Public Const REG_DWORD_BIG_ENDIAN = 5
Public Const REG_DWORD = REG_DWORD_LITTLE_ENDIAN
Public Const REG_LINK = 6                     ' // Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7                 ' // Multiple Unicode strings
Public Const REG_RESOURCE_LIST = 8            ' // Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9 ' // Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10




' This value for "largest possible Registry buffer"
' is based on Microsoft's recommendation that any entries larger
' than this should be stored as files, with a UNC to them in the Registry
Private Const mREG_BUFFER_SIZE = 2048





' Reformat the Registry "Buffer O' Bytes" data type into something more VB friendly
'
' Values are returned in pValue & pBuffer will be unchanged
'
' NB - pBufferLen is the size of the buffer,
' and for strings it wil include the terminating nulls
' (this is different from how Windows returns the value names)
'
Private Sub MakeRegistryBufferIntoVB(pValueName As String, pBuffer As String, ByVal pBufferLen As Long, ByVal pType As Long, pValue As Variant)

Dim lHack As Long

Dim lLongValue As Long
Dim lStringValue As String
Dim llLength As Long


    ' We put this in a variable, not a constant, because of problems with VB
    ' Constants are implicitly typed as integers, and that would cause an overflow in this case
    lHack = &H100
    lLongValue = 0
    pValue = 0

    
    Select Case pType
    Case REG_BINARY
    
    Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
      lLongValue = (&H1000000 * (Asc(Mid$(pBuffer, 4, 1)) Mod &H80)) _
        + (&H10000 * Asc(Mid$(pBuffer, 3, 1))) _
        + (lHack * Asc(Mid$(pBuffer, 2, 1))) _
        + (&H1 * Asc(Mid$(pBuffer, 1, 1)))
      If ((Asc(Mid$(pBuffer, 4, 1)) And &H80) <> 0) Then lLongValue = -(lLongValue Xor &H7FFFFFFF)
      pValue = lLongValue
      
    Case REG_DWORD_BIG_ENDIAN
      lLongValue = (&H1000000 * (Asc(Mid$(pBuffer, 1, 1)) Mod &H80)) _
        + (&H10000 * Asc(Mid$(pBuffer, 2, 1))) _
        + (lHack * Asc(Mid$(pBuffer, 3, 1))) _
        + (&H1 * Asc(Mid$(pBuffer, 4, 1)))
      If ((Asc(Mid$(pBuffer, 1, 1)) And &H80) <> 0) Then lLongValue = -(lLongValue Xor &H7FFFFFFF)
      pValue = lLongValue
    
    Case REG_LINK
    
    Case REG_MULTI_SZ
      ' We probably want to return these as an array of strings, now that VB4 Variants support arrays of strings
      lStringValue = left$(pBuffer, InStr(pBuffer, Chr$(0) & Chr$(0)) - 1)
      pValue = lStringValue
    
    Case REG_RESOURCE_LIST
    
    Case REG_SZ
      ' pBufferLen is the buffer length, not the string length (i.e. it includes the terminating null)
      lStringValue = left$(pBuffer, pBufferLen - 1)
      pValue = lStringValue
      
    Case REG_EXPAND_SZ
      lStringValue = String$(mREG_BUFFER_SIZE, Chr$(0))
      llLength = ExpandEnvironmentStrings(ByVal pBuffer, ByVal lStringValue, mREG_BUFFER_SIZE)
      pValue = left$(lStringValue, llLength)
    
    Case Else ' REG_NONE
    End Select
    
End Sub

' Reformat a VB variant into a buffer for use with the Registry
'
' Values are returned in pValue & pBuffer will be unchanged
'
'
'
' If pType = REG_NONE (0) an automatic selection will be made
'
Private Sub MakeVBIntoRegistryBuffer(pValue As Variant, pType As Long, pBuffer As String, pBufferLen As Long)

  ' Sort out the type of data
  If (pType = REG_NONE) Then
    Select Case VarType(pValue)
    Case vbInteger, vbLong, vbByte, vbBoolean
      pType = REG_DWORD
    Case vbString
      pType = REG_SZ
    Case Else ' Default to a string representation for the moment - maybe should be binary if we work out how to manipulate them
      pType = REG_SZ
    End Select
  End If
    
  
  Select Case pType
  
  Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
    pBuffer = Chr$((CLng(pValue)) And &HFF) _
      & Chr$((CLng(pValue) \ &H100) And &HFF) _
      & Chr$((CLng(pValue) \ &H10000) And &HFF) _
      & Chr$((CLng(pValue) \ &H1000000) And &HFF)
    pBufferLen = 4
    
  Case REG_DWORD_BIG_ENDIAN
    pBuffer = Chr$((CLng(pValue) \ &H1000000) And &HFF) _
      & Chr$((CLng(pValue) \ &H10000) And &HFF) _
      & Chr$((CLng(pValue) \ &H100) And &HFF) _
      & Chr$((CLng(pValue)) And &HFF)
    pBufferLen = 4
 
  Case REG_SZ, REG_EXPAND_SZ
    Select Case VarType(pValue)
    Case vbString
      pBuffer = pValue & Chr$(0)
    Case Else
      pBuffer = Format$(pValue) & Chr$(0)
    End Select
    pBufferLen = Len(pBuffer)
    
  Case Else
    ' Fail - we don't recognise this type
    Error -1
  End Select
  
  
End Sub

' Wrapper function for RegCreateKey
'
Public Function RegCreateKey(ByVal hKey As Long, lpSubKey As String, phkResult As Long) As Long

  RegCreateKey = RegCreateKeyA(hKey, ByVal lpSubKey, phkResult)
End Function

' Wrapper function for RegDeleteKey
'
Public Function RegDeleteKey(ByVal hKey As Long, lpSubKey As String) As Long

  RegDeleteKey = RegDeleteKeyA(hKey, ByVal lpSubKey)
End Function

' Wrapper function for RegDeleteValue
'
' Unlike the API RegDeleteValue, this function supports the deletion
' of a value from a sub-key, lpSubKey
'
' lpSubKey may otherwise be empty, in which case the value to be deleted
' should be a value of hKey itself
'
Public Function RegDeleteValue(ByVal hKey As Long, lpSubKey As String, lpValueName As String) As Long

Dim lRetVal As Long, lMinorRetVal As Long
Dim lhKey As Long


  lRetVal = ERROR_SUCCESS
  ' Open any subkeys as necessary
  If (Len(lpSubKey) = 0) Then
    lhKey = hKey
  Else
#If UseRegExFunctions Then
    lRetVal = RegOpenKey(hKey, lpSubKey, KEY_WRITE, lhKey)
#Else
    lRetVal = RegOpenKey(hKey, lpSubKey, lhKey)
#End If
  End If
  
  ' Delete the value
  If (lRetVal = ERROR_SUCCESS) Then lRetVal = RegDeleteValueA(lhKey, ByVal lpValueName)
  
  ' Close the key
  If (lRetVal = ERROR_SUCCESS) Then
    lMinorRetVal = RegCloseKey(lhKey)
    If (lRetVal = ERROR_SUCCESS) Then lRetVal = lMinorRetVal
  End If
  
  RegDeleteValue = lRetVal
End Function


' Wrapper function for RegEnumKey
'
Public Function RegEnumKey(ByVal hKey As Long, ByVal dwIndex As Long, lpName As String) As Long

Dim llRetVal As Long
Dim lcbName As Long


  lcbName = mREG_BUFFER_SIZE
  lpName = String$(lcbName, Chr$(0))
  llRetVal = RegEnumKeyA(hKey, dwIndex, ByVal lpName, lcbName)
  ' lcbName is not set on return from RegEnumKeyA, so we must search for the end of the string ourselves
  ' (see RegEnumValue)
  lpName = left$(lpName, InStr(lpName, Chr$(0)) - 1)
  
  RegEnumKey = llRetVal
End Function




' Wrapper function for RegEnumValue
'
' Values are returned as a VB Variant
'
Public Function RegEnumValue(ByVal hKey As Long, ByVal dwIndex As Long, lpValueName As String, lpType As Long, pValue As Variant) As Long

Dim lpLongValue As Long

Dim lRetVal As Long
Dim cbValueName As Long
Dim lpValueBuffer As String
Dim cbValue As Long

Dim lpReserved As Long


  cbValueName = mREG_BUFFER_SIZE
  lpValueName = String$(cbValueName, Chr$(0))
  cbValue = mREG_BUFFER_SIZE
  lpValueBuffer = String$(cbValue, Chr$(0))
  
  lpReserved = 0
  lRetVal = RegEnumValueA(hKey, dwIndex, ByVal lpValueName, cbValueName, lpReserved, lpType, ByVal lpValueBuffer, cbValue)
  If lRetVal = ERROR_SUCCESS Then
    ' cbValueName is set to the length of the string on return,
    ' cbValue is set to the length of the buffer
    ' cbValue includes the terminating null, cbValueName doesn't
    ' (See RegEnumKey for a 3rd way of passing these parameters !)
    lpValueName = left$(lpValueName, cbValueName)
    MakeRegistryBufferIntoVB lpValueName, lpValueBuffer, cbValue, lpType, pValue
  End If
  
  RegEnumValue = lRetVal
End Function



' Wrapper function for RegOpenKey
'
' When calling this under 32-bit code, pass KEY_READ or KEY_WRITE (or similar) constants
'
#If UseRegExFunctions Then
Public Function RegOpenKey(ByVal hKey As Long, lpSubKey As String, ByVal pSAM As Long, phkResult As Long) As Long
#Else
Public Function RegOpenKey(ByVal hKey As Long, lpSubKey As String, phkResult As Long) As Long
#End If


#If UseRegExFunctions Then
Dim lReserved As Long
 
  RegOpenKey = RegOpenKeyExA(hKey, ByVal lpSubKey, lReserved, pSAM, phkResult)
#Else
  RegOpenKey = RegOpenKeyA(hKey, ByVal lpSubKey, phkResult)
#End If
End Function


' Wrapper function for RegQueryValue
'
Public Function RegQueryValue(ByVal phKey As Long, pSubKey As String, pValue As Variant) As Long
  
Dim lRetVal As Long
Dim lpReserved As Long
Dim lpType As Long
Dim lpData As String
Dim lpcbValue As Long


  lpReserved = 0
  lpcbValue = mREG_BUFFER_SIZE
  lpData = String$(lpcbValue, Chr$(0))
  
' We can only use RegQueryValueEx if we used RegOpenKeyEx to open it
#If UseRegExFunctions Then
  lRetVal = RegQueryValueExA(phKey, ByVal pSubKey, lpReserved, lpType, ByVal lpData, lpcbValue)
#Else
  lRetVal = RegQueryValueA(phKey, ByVal pSubKey, ByVal lpData, lpcbValue)
  lpType = REG_SZ ' 16 bit Registry only supports strings
#End If

  If lRetVal = ERROR_SUCCESS Then
    MakeRegistryBufferIntoVB pSubKey, lpData, lpcbValue, lpType, pValue
  End If
  RegQueryValue = lRetVal
End Function


' Wrapper function for RegSetValue
'
' phKey   - key to a hive
' pSubKey - string for the sub-key from phKey
' pType   - data type, REG_SZ etc. as defined in Reg_API.BAS
'           Use pType = REG_NONE to give an automatic selection
' pData   - data to be stored
'
Public Function RegSetValue(ByVal phKey As Long, pSubKey As String, ByVal pType As Long, ByVal pData As Variant) As Long

Dim lBuffer As String
Dim lBufferLen As Long


  MakeVBIntoRegistryBuffer pData, pType, lBuffer, lBufferLen


#If UseRegExFunctions Then
Dim lReserved As Long

  lReserved = 0
  RegSetValue = RegSetValueExA(phKey, ByVal pSubKey, lReserved, pType, ByVal lBuffer, lBufferLen)
#Else
  RegSetValue = RegSetValueA(phKey, ByVal pSubKey, pType, ByVal lBuffer, lBufferLen)
#End If
End Function


