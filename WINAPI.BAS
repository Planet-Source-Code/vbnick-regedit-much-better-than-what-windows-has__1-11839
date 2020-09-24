Attribute VB_Name = "WinAPI"
'
'  Windows API stuff that is useful to any project
'since written in VB4-32 bit , might have problems with other
'versions...
Option Explicit


Type Rect
   left As Integer
   top As Integer
   right As Integer
   bottom As Integer
End Type




Global Const HINSTANCE_ERROR = 32
#If Win16 Then
Declare Function LoadLibrary Lib "Kernel" (ByVal lpLibFileName As String) As Integer
Declare Sub FreeLibrary Lib "Kernel" (ByVal hLibModule As Integer)
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags As Integer) As Long
Declare Function WinExec Lib "Kernel" (ByVal lpCmdLine As String, ByVal nCmdShow As Integer) As Integer
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long

#Else
#If Win32 Then

Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long
Declare Function WinExec Lib "Kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Declare Function GetVersion Lib "Kernel32" () As Long
#End If
#End If



' ***  System metrics  ***
Global Const SM_CXSCREEN = 0    ' Width of screen
Global Const SM_CYSCREEN = 1    ' Height of screen
Global Const SM_CXVSCROLL = 2
Global Const SM_CXBORDER = 5    ' Width of window frame that cannot be sized
Global Const SM_CYBORDER = 6    ' Height of window frame that cannot be sized
Global Const SM_CXDLGFRAME = 7
Global Const SM_CXICON = 11
Global Const SM_CXCURSOR = 13
Global Const SM_CYMENU = 15     ' Height of single line menu bar
Global Const SM_CXFULLSCREEN = 16 ' Width of window client area for a full-screen window
Global Const SM_CXMIN = 28      ' Minimum width of window
Global Const SM_CYMIN = 29
Global Const SM_CXSIZE = 30     ' Width of bitmaps contained in the title bar
Global Const SM_CYSIZE = 31     ' Height of bitmaps contained in the title bar
Global Const SM_CXFRAME = 32    ' Width of window frame that can be sized
Global Const SM_CYFRAME = 33    ' Height of window frame that can be sized


#If Win16 Then
Declare Function GetSystemMetrics Lib "User" Alias "getSystemMetrics" (ByVal nIndex As Integer) As Integer
#Else
#If Win32 Then
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If
#End If



' User interfaces

#If Win16 Then
Declare Sub MessageBeep Lib "User" (ByVal wType As Integer)
#Else
#If Win32 Then
Declare Sub MessageBeep Lib "user32" (ByVal wType As Long)

Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
#End If
#End If


' ***  Cursors  ***

Type POINTAPI
  X As Long
  Y As Long
End Type


#If Win16 Then
Declare Sub GetCursorPos Lib "User" (lpPoint As POINTAPI)
#Else
#If Win32 Then
Declare Sub GetCursorPos Lib "user32" (lpPoint As POINTAPI)
#End If
#End If


' ***  Windows  ***
'Now define the constants for Task list functions etc.
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2


#If Win16 Then
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function ShowWindow Lib "User" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function SetFocusAPI Lib "User" Alias "SetFocus" (ByVal hwnd As Integer) As Integer
Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function GetClassName Lib "User" (ByVal hwnd%, ByVal lpClassName$, ByVal nMaxCount%) As Integer
Declare Function SetParent Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
Declare Function SetWindowPos Lib "User" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer

Declare Sub GetClientRect Lib "User" (ByVal hwnd As Integer, lpRect As Rect)
Declare Function RedrawWindow Lib "User" (ByVal hwnd As Integer, lprcUpdate As Any, ByVal hrgnUpdate As Integer, ByVal fuRedraw As Integer) As Integer
Declare Sub InvalidateRect Lib "User" (ByVal hwnd As Integer, lpRect As Rect, ByVal bErase As Integer)
Declare Sub UpdateWindow Lib "User" (ByVal hwnd As Integer)

Declare Function GetNextWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal sBuff As String, ByVal wLgth As Integer) As Integer

Declare Function FlashWindow Lib "User" (ByVal hwnd As Integer, ByVal bInvert As Integer) As Integer

Declare Function GetFocus% Lib "User" ()
Declare Function GetParent Lib "User" (ByVal hwnd As Integer) As Integer

#Else
#If Win32 Then

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function GetNextWindow Lib "user32" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetFocus Lib "user32" () As Long

#End If
#End If



' ***  Menus  ***
#If Win16 Then
Declare Function GetSystemMenu Lib "User" (ByVal hwnd%, ByVal bRevert%) As Integer
Declare Function DeleteMenu Lib "User" (ByVal hMenu%, ByVal iditem%, ByVal wFlags%) As Integer
Declare Function CreatePopupMenu Lib "User" () As Integer
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function TrackPopupMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nReserved As Integer, ByVal hwnd As Integer, lpReserved As Any) As Integer
#Else
#If Win32 Then
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As Rect) As Long
#End If
#End If


Public Const SC_SIZE = &HF000
Public Const SC_MOVE = &HF010
Public Const SC_CLOSE = &HF060
Public Const SC_RESTORE = &HF120
Public Const SC_TASKLIST = &HF130
Public Const MF_BYCOMMAND = &H0


' GetTempFileName() Flags
'
Public Const TF_FORCEDRIVE = &H80

#If Win16 Then
Declare Function GetTempDrive Lib "Kernel" (ByVal cDriveLetter As Integer) As Integer
Declare Function GetTempFileName Lib "Kernel" (ByVal cDriveLetter As Integer, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer
#Else
#If Win32 Then
Declare Function GetTempFileName Lib "Kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
#End If
#End If


' ***  Graphics  ***

Global Const RDW_INVALIDATE = &H1
Global Const RDW_INTERNALPAINT = &H2
Global Const RDW_ERASE = &H4


Global Const TA_LEFT = 0
Global Const TA_RIGHT = 2
Global Const TA_CENTER = 6

Global Const TA_TOP = 0
Global Const TA_BOTTOM = 8
Global Const TA_BASELINE = 24

Global Const TA_NOUPDATECP = 0
Global Const TA_UPDATECP = 1



#If Win16 Then
Declare Function GetDC Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function ReleaseDC Lib "User" (ByVal hwnd As Integer, ByVal hDC As Integer) As Integer
Declare Function GetSysColor Lib "User" (ByVal nIndex As Integer) As Long

Declare Function SetTextAlign Lib "GDI" (ByVal hDC As Integer, ByVal wFlags As Integer) As Integer

Declare Function TextOut Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long) As Long

Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function PatBlt Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal dwRop As Long) As Integer

Declare Function FloodFill Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Long) As Integer
Declare Function ExtFloodFill Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Long, ByVal wFillType As Integer) As Integer

Declare Function CreatePatternBrush% Lib "GDI" (ByVal hBitmap As Integer)

Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer
#Else
#If Win32 Then
#End If
#End If


' ***  Printing  ***
#If Win16 Then
Declare Function Escape% Lib "GDI" (ByVal hDC%, ByVal nEsc%, ByVal nLen%, lpData As Any, lpOut As Any)

#Else
#If Win32 Then

#End If
#End If

Global Const PORTRAIT = 1
Global Const LANDSCAPE = 2
Global Const GETSETPAPERORIENT = 30



' ***  Messages  ***
#If Win16 Then
Declare Function PostMessage Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Integer
Declare Function SendMessage Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
#Else
#If Win32 Then

Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
#End If
#End If

Global Const WM_SETREDRAW = &HB
Global Const WM_PAINT = &HF
Global Const WM_CLOSE = &H10
Global Const WM_NCHITTEST = &H84
Global Const WM_NCACTIVATE = &H86
Global Const WM_NCLBUTTONDOWN = &HA1
Global Const WM_NCLBUTTONDBLCLK = &HA3
Global Const WM_COMMAND = &H111
Global Const WM_SYSCOMMAND = &H112

Global Const WM_USER = &H400

Global Const CB_SHOWDROPDOWN = WM_USER + 15

Global Const EM_GETMODIFY = WM_USER + 8
Global Const EM_SETMODIFY = WM_USER + 9
Global Const EM_LIMITTEXT = WM_USER + 21
Global Const EM_CANUNDO = WM_USER + 22
Global Const EM_UNDO = WM_USER + 23
Global Const EM_SETREADONLY = WM_USER + 31

Global Const LB_FINDSTRING = WM_USER + 16
Global Const LB_SETTABSTOPS = WM_USER + 19


' ***  User Profile Routines  ***
' NOTE: The lpKeyName argument for GetProfileString, WriteProfileString,
'       GetPrivateProfileString, and WritePrivateProfileString can be either
'       a string or NULL.  This is why the argument is defined as "By Any".
'          For example, to pass a string specify   ByVal "wallpaper"
'          To pass NULL specify                    ByVal 0&
'       You can also pass NULL for the lpString argument for WriteProfileString
'       and WritePrivateProfileString
#If Win16 Then
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer

#Else
#If Win32 Then

Declare Function GetProfileInt Lib "Kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, lpKeyName As Any, ByVal nDefault As Long) As Long
Declare Function GetProfileString Lib "Kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function WriteProfileString Lib "Kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Declare Function GetProfileSection Lib "Kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function WriteProfileSection Lib "Kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long

Declare Function GetPrivateProfileInt Lib "Kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileSection Lib "Kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "Kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

#End If
#End If


#If Win16 Then
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

#Else
#If Win32 Then

Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

#End If
#End If




' Clipboards
#If Win16 Then
Declare Function SetClipboardData Lib "User" (ByVal wFormat As Integer, ByVal hMem As Integer) As Integer

#Else
#If Win32 Then

#End If
#End If





'*************************** Device capabilities ***************************
'
#If Win32 Then
'  size of a device name string
Public Const CCHDEVICENAME = 32

'  size of a form name string
Public Const CCHFORMNAME = 32

Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type


Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As DEVMODE) As Long


Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long

Public Const PRINTER_ENUM_DEFAULT = &H1
Public Const PRINTER_ENUM_LOCAL = &H2
Public Const PRINTER_ENUM_CONNECTIONS = &H4
Public Const PRINTER_ENUM_FAVORITE = &H4
Public Const PRINTER_ENUM_NAME = &H8
Public Const PRINTER_ENUM_REMOTE = &H10
Public Const PRINTER_ENUM_SHARED = &H20
Public Const PRINTER_ENUM_NETWORK = &H40


Type PRINTER_INFO_1
        flags As Long
        pDescription As String
        pName As String
        pComment As String
End Type

#End If


'*************************** Windows 3.1 Extensions ***************************
' Window 3.1 commands to pass WinHelp()
'Global Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
'Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
'Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
'Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
'Global Const HELP_COMMAND = &H102    ' Execute Help macro
'Global Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
'Global Const HELP_SETWINPOS = &H203  ' Display and position Help window


#If Win16 Then
Declare Function lstrcpy Lib "Kernel" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
' Pointer weirdness here, where a pointer is being passed as a long
' This doesn't match the API spec, but it's how MSoft do it in the setupkit
' If this is ever changed (fixed ?) it will affect VERSION.BAS
Declare Function lstrcpyn Lib "Kernel" (ByVal lpszString1 As Any, ByVal lpszString2 As Long, ByVal cChars As Integer) As Long

#Else
#If Win32 Then

Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcmp Lib "Kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcmpi Lib "Kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcpy Lib "Kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcpyn Lib "Kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
Declare Function lstrlen Lib "Kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

#End If
#End If



' Retrieves the class name for the specified window.
'
' Input Parameters: hWnd%  - Window handle for which to obtain the class
'                            name.
'
' Return Value    : String - Class name (lpClassName$) for the specified
'                            window.
'
' Global Variables: None
'
' Objects         : None
'
Function cmnGetClassName$(hwnd%)
  
Const lCAPI_MAX_CLASS_NAME_LENGTH = 255

Dim lClassNameLength%
Dim lpClassName$


  ' Preallocate class name buffer.
  lpClassName$ = String$(lCAPI_MAX_CLASS_NAME_LENGTH, " ")

  ' Retrieve class name.
  lClassNameLength% = GetClassName(hwnd%, lpClassName$, lCAPI_MAX_CLASS_NAME_LENGTH)

  cmnGetClassName$ = Trim$(left$(lpClassName$, lClassNameLength%))
End Function

' Determines if the contents of an edit control have been changed.
'
' Input Parameters: txtTextBox - The text box to check for modified contents.
'
' Return Value    : Integer    - Boolean that contains True if the text box
'                                has been modified or False if the text box
'                                has not been modified.
'
' Global Variables: None
'
' Objects         : None
'
Function cmnGetModifiedState%(ctrl As Control)
  cmnGetModifiedState% = SendMessage(ctrl.hwnd, EM_GETMODIFY, 0, 0)
End Function

' Retrieves an integer for the specified entry in an initialization file.
'
' Input Parameters: lpApplicationName$ - The section to search for the entry.
'                   lpKeyName$         - The key name or entry to retrieve.
'                   Default%           - The default value to return if the
'                                        specified entry is not found.
'                   lpFileName$        - The name of the private
'                                        initialization file.  If a full path
'                                        name is not specified, Windows will
'                                        search for the file in the Windows
'                                        directory.
'
' Return Value    : Integer            - The value for the INI file key name
'                                        or entry, or the default value if the
'                                        specified entry is not found.
'
' Global Variables: None
'
' Objects         : None
'
Function cmnGetPrivateProfileInt%(lpApplicationName$, lpKeyName$, ByVal Default%, lpFileName$)

  cmnGetPrivateProfileInt% = GetPrivateProfileInt(ByVal lpApplicationName$, ByVal lpKeyName$, Default%, lpFileName$)
End Function

' Retrieves the string for the specified entry in an initialization file.
'
' Input Parameters: lpApplicationName$ - The section to search for the entry.
'                   lpKeyName$         - The key name or entry to retrieve.
'                   lpszDefault$       - The default value to return if the
'                                        specified entry is not found.
'                   lpFileName$        - The name of the private
'                                        initialization file.  If a full path
'                                        name is not specified, Windows will
'                                        search for the file in the Windows
'                                        directory.
'                   nSize%             - The maximum number of characters to
'                                        load into the lpszReturnString
'                                        buffer.
'
' Return Value    : String             - The value for the INI file key name
'                                        or entry, or the default value if the
'                                        specified entry is not found.
'
' Global Variables: None
'
' Objects         : None
'
Function cmnGetPrivateProfileString$(lpApplicationName$, lpKeyName$, lpszDefault$, lpFileName$, nSize%)
  
Dim lpszReturnString$
Dim lStringLength%
  

  ' 1 is added to nSize%, to allow space for the terminating null
  lpszReturnString$ = String$(nSize% + 1, 0)
  If Len(lpKeyName$) > 0 Then
    lStringLength% = GetPrivateProfileString(lpApplicationName$, ByVal lpKeyName$, lpszDefault$, lpszReturnString$, nSize% + 1, lpFileName$)
  Else
    lStringLength% = GetPrivateProfileString(lpApplicationName$, Null, lpszDefault$, lpszReturnString$, nSize% + 1, lpFileName$)
  End If
  cmnGetPrivateProfileString$ = left$(lpszReturnString$, lStringLength%)
End Function

' Retrieves the string for the specified entry in the WIN.INI initialization
' file.
'
' Input Parameters: lpApplicationName$ - The section to search for the entry.
'                   lpKeyName$         - The key name or entry to retrieve.
'                   lpszDefault$       - The default value to return if the
'                                        specified entry is not found.
'                   nSize%             - The maximum number of characters to
'                                        load into the lpszReturnString
'                                        buffer.
'
' Return Value    : String             - The value for the WIN.INI file key
'                                        name or entry, or the default value
'                                        if the specified entry is not found.
'
' Global Varaibles: None
'
' Objects         : None
'
Function cmnGetProfileString$(lpApplicationName$, lpKeyName$, lpszDefault$, nSize%)

Dim lpszReturnString$
Dim lStringLength%


  ' Blank fill return string.
  lpszReturnString$ = String$(nSize% + 1, 0)
  
  ' Retrieve entry value.
  If Len(lpKeyName$) > 0 Then
    lStringLength% = GetProfileString(lpApplicationName$, ByVal lpKeyName$, lpszDefault$, lpszReturnString$, nSize% + 1)
  Else
    lStringLength% = GetProfileString(lpApplicationName$, Null, lpszDefault$, lpszReturnString$, nSize% + 1)
  End If
  
  ' Return blank stripped entry value
  cmnGetProfileString$ = left$(lpszReturnString$, lStringLength%)
End Function

' This routine sets the form mouse pointer, to the specified pointer, for all
' of the loaded forms.
'
' Input Parameters : lMousePointerType% - The type of mouse pointer to set
'                                         the forms mouse pointer to.
'
' Output parameters: None
'
' Global Variables : None
'
' Objects          : Forms
'
Sub cmnSetFormsMousePointer(ByVal lMousePointerType%)
  
Dim i%
  
  For i% = 0 To (Forms.Count - 1)
    Forms(i%).MousePointer = lMousePointerType%
  Next i%
End Sub

' Set the modification flag for an edit control.
'
' Input Parameters  : hWnd%             - The handle of the window whose edit
'                                         control the modification flag is
'                                         being set for.
'                     lModifyFlagValue% - Boolean that contains True or False
'                                         to indicate the value to set
'                                         modification flag to.
'
' Output Parameters : None
'
' Global Variables  : None
'
' Objects           : None
'
Sub cmnSetModifyFlag(ctrl As Control, lModifyFlagValue%)

Dim lReturnValue%


  lReturnValue% = SendMessage(ctrl.hwnd, EM_SETMODIFY, lModifyFlagValue%, 0)
End Sub

' Set an edit control to read only.
'
' Input Parameters  : hWnd%             - The handle of the window whose edit
'                                         control is being set to read only.
'                     lSetReadOnly%     - Boolean that contains True to set
'                                         an edit control to read only or
'                                         False not to set an edit control
'                                         to read only.
'
' Output Pararmeters: None
'
' Global Varaiables : None
'
' Objects           : None
'
Sub cmnSetReadOnlyFlag(ctrl As Control, lSetReadOnly%)

Dim lReturnValue%

  lReturnValue% = SendMessage(ctrl.hwnd, EM_SETREADONLY, lSetReadOnly%, 0)
End Sub

' Writes a string into the specified enty in an initialization file.
'
' Input Parameters: lpSection$         - The section to search for the entry.
'                   lpEntry$           - The key name or entry to retrieve.
'                   lpWriteString$     - The value to write
'                   lpFileName$        - The name of the private
'                                        initialization file.  If a full path
'                                        name is not specified, Windows will
'                                        search for the file in the Windows
'                                        directory.
'
' If lpApplicationName$ or lpKeyName$ are blank, the relevant sections will be deleted
'
' Return Value    : None
'
' Global Variables: None
'
' Objects         : None
'
Sub cmnWritePrivateProfileString(lpSection$, lpEntry$, lpWriteString$, lpFileName$)
  
  If (Len(lpEntry$) = 0) Then
    ' Delete section
    If WritePrivateProfileString(ByVal lpSection$, ByVal 0&, ByVal lpWriteString$, ByVal lpFileName$) <> 0 Then
      ' Success
    End If

  ElseIf (Len(lpWriteString$) = 0) Then
    ' Delete entry
    If WritePrivateProfileString(ByVal lpSection$, ByVal lpEntry$, ByVal 0&, ByVal lpFileName$) <> 0 Then
      ' Success
    End If

  Else
    ' Write an entry
    If WritePrivateProfileString(ByVal lpSection$, ByVal lpEntry$, ByVal lpWriteString$, ByVal lpFileName$) <> 0 Then
      ' Success
    End If
  End If
End Sub

