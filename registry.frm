VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRegistry 
   Caption         =   "Registry  ==>BY KAYHAN TANRISEVEN...IF YOU LIKE THE CODE....:)"
   ClientHeight    =   6420
   ClientLeft      =   855
   ClientTop       =   1800
   ClientWidth     =   7905
   Icon            =   "registry.frx":0000
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6420
   ScaleWidth      =   7905
   Tag             =   "frmRegistry"
   Begin VB.PictureBox picView 
      BorderStyle     =   0  'None
      Height          =   5292
      Index           =   4
      Left            =   240
      ScaleHeight     =   5295
      ScaleWidth      =   7455
      TabIndex        =   49
      Top             =   600
      Width           =   7452
      Begin VB.Frame fraExtensions 
         Caption         =   "Extensions"
         Height          =   1215
         Left            =   0
         TabIndex        =   51
         Top             =   4080
         Width           =   7455
         Begin VB.CommandButton cmdRegister 
            Caption         =   "Register"
            Height          =   375
            Left            =   5040
            TabIndex        =   57
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optExtensionFindAppOpenWith 
            Caption         =   "Open With"
            Height          =   255
            Left            =   3120
            TabIndex        =   56
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton optExtensionFindAppServer 
            Caption         =   "Server"
            Height          =   255
            Left            =   3120
            TabIndex        =   55
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdFindAppByExtension 
            Caption         =   "Find App"
            Height          =   375
            Left            =   1800
            TabIndex        =   54
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtExtensionApp 
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   810
            Width           =   6975
         End
         Begin VB.TextBox txtExtension 
            Height          =   285
            Left            =   240
            TabIndex        =   52
            Top             =   360
            Width           =   1455
         End
      End
      Begin ComctlLib.TreeView tvwExtensions 
         Height          =   3975
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   7011
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.PictureBox picView 
      BorderStyle     =   0  'None
      Height          =   5292
      Index           =   1
      Left            =   240
      ScaleHeight     =   5295
      ScaleWidth      =   7455
      TabIndex        =   2
      Top             =   600
      Width           =   7452
      Begin VB.Frame fraRegistryKey 
         Caption         =   "Initial Key"
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   4560
         Width           =   7215
         Begin VB.TextBox txtRegistryKey 
            Height          =   285
            Left            =   3720
            TabIndex        =   24
            Top             =   240
            Width           =   3135
         End
         Begin VB.OptionButton optRegistryKey 
            Caption         =   "Microsoft"
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optRegistryKey 
            Caption         =   "Application"
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optRegistryKey 
            Caption         =   "Custom"
            Height          =   315
            Index           =   3
            Left            =   2520
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
      End
      Begin ComctlLib.TreeView tvwRegistry 
         Height          =   4095
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7223
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
      Begin ComctlLib.TabStrip tabRegistry 
         Height          =   5295
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   9340
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "H_KEY_CURRENT_USER"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "H_KEY_LOCAL_MACHINE"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "H_KEY_CLASSES_ROOT"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picView 
      BorderStyle     =   0  'None
      Height          =   5292
      Index           =   3
      Left            =   240
      ScaleHeight     =   5295
      ScaleWidth      =   7455
      TabIndex        =   39
      Top             =   600
      Width           =   7452
      Begin VB.Frame fraAppPaths 
         Caption         =   "Application Path"
         Height          =   1695
         Left            =   0
         TabIndex        =   40
         Top             =   3480
         Width           =   7455
         Begin VB.TextBox txtAppPathPath 
            Height          =   285
            Left            =   1800
            TabIndex        =   47
            Top             =   1080
            Width           =   4095
         End
         Begin VB.CommandButton cmdAppPathAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   840
            TabIndex        =   43
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtAppPathEXEPath 
            Height          =   285
            Left            =   1800
            TabIndex        =   42
            Top             =   480
            Width           =   4095
         End
         Begin VB.TextBox txtAppPathName 
            Height          =   285
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Path"
            Height          =   255
            Left            =   1800
            TabIndex        =   48
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Application"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Executable"
            Height          =   255
            Left            =   1800
            TabIndex        =   44
            Top             =   240
            Width           =   1575
         End
      End
      Begin ComctlLib.TreeView tvwAppPaths 
         Height          =   3255
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5741
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.PictureBox picView 
      BorderStyle     =   0  'None
      Height          =   5292
      Index           =   2
      Left            =   240
      ScaleHeight     =   5295
      ScaleWidth      =   7455
      TabIndex        =   18
      Top             =   600
      Width           =   7452
      Begin VB.Frame fraRunThisApplication 
         Caption         =   "Run This Application"
         Height          =   1695
         Left            =   3960
         TabIndex        =   25
         Top             =   3480
         Width           =   3495
         Begin VB.CommandButton cmdRemoveApp 
            Caption         =   "&Remove"
            Height          =   375
            Left            =   1920
            TabIndex        =   37
            Top             =   960
            Width           =   735
         End
         Begin VB.CommandButton cmdRestartApp 
            Caption         =   "&Add"
            Height          =   375
            Left            =   960
            TabIndex        =   35
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtRestartParameters 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label5 
            Caption         =   "Parameters"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame fraRun 
         Caption         =   "Run"
         Height          =   1695
         Left            =   0
         TabIndex        =   26
         Top             =   3480
         Width           =   3735
         Begin VB.CheckBox chkRunStartup 
            Caption         =   "Run at startup"
            Height          =   375
            Left            =   360
            TabIndex        =   36
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtRunValueName 
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtRunValue 
            Height          =   285
            Left            =   1800
            TabIndex        =   29
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmdRunAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   2160
            TabIndex        =   28
            Top             =   960
            Width           =   735
         End
         Begin VB.CheckBox chkRunOnce 
            Caption         =   "Run Once"
            Height          =   375
            Left            =   360
            TabIndex        =   27
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblValue 
            Caption         =   "Value"
            Height          =   255
            Left            =   1800
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblValueName 
            Caption         =   "Value Name"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   1095
         End
      End
      Begin ComctlLib.TreeView tvwRun 
         Height          =   3255
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5741
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.PictureBox picView 
      BorderStyle     =   0  'None
      Height          =   5292
      Index           =   5
      Left            =   240
      ScaleHeight     =   5295
      ScaleWidth      =   7455
      TabIndex        =   1
      Top             =   600
      Width           =   7452
      Begin ComctlLib.TreeView tvwODBC 
         Height          =   5292
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   7452
         _ExtentX        =   13150
         _ExtentY        =   9340
         _Version        =   327682
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.PictureBox picView 
      BorderStyle     =   0  'None
      Height          =   5292
      Index           =   6
      Left            =   240
      ScaleHeight     =   5295
      ScaleWidth      =   7455
      TabIndex        =   3
      Top             =   600
      Width           =   7452
      Begin VB.TextBox txtDescription 
         Height          =   360
         Left            =   1320
         TabIndex        =   16
         Top             =   1920
         Width           =   2412
      End
      Begin VB.CommandButton cmdDeleteODBC 
         Caption         =   "Delete ODBC"
         Height          =   492
         Left            =   4080
         TabIndex        =   15
         Top             =   1440
         Width           =   972
      End
      Begin VB.ComboBox cboDSN 
         Height          =   360
         Left            =   1320
         TabIndex        =   14
         Top             =   840
         Width           =   2412
      End
      Begin VB.ComboBox cboServer 
         Height          =   360
         Left            =   1320
         TabIndex        =   13
         Top             =   1200
         Width           =   2412
      End
      Begin VB.TextBox txtDatabase 
         Height          =   360
         Left            =   1320
         TabIndex        =   9
         Top             =   1560
         Width           =   2412
      End
      Begin VB.CommandButton cmdCreateODBC 
         Caption         =   "Create ODBC"
         Height          =   492
         Left            =   4080
         TabIndex        =   8
         Top             =   840
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "Description:"
         Height          =   252
         Left            =   360
         TabIndex        =   17
         Top             =   1920
         Width           =   852
      End
      Begin VB.Label Label3 
         Caption         =   "Database:"
         Height          =   252
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   852
      End
      Begin VB.Label Label2 
         Caption         =   "Server:"
         Height          =   252
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "DSN:"
         Height          =   252
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   852
      End
      Begin VB.Label lblTabName 
         Alignment       =   2  'Center
         Caption         =   "Change ODBC settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   7452
      End
   End
   Begin ComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   38
      Top             =   6165
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13520
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TabStrip tabView 
      Height          =   5892
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7692
      _ExtentX        =   13573
      _ExtentY        =   10398
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Registry"
            Object.Tag             =   ""
            Object.ToolTipText     =   "View registry settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Run"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Show programs scheduled for auto-run"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "App Paths"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Show paths stored for applications"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Extensions"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ODBC"
            Object.Tag             =   ""
            Object.ToolTipText     =   "View ODBC data sources"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Change ODBC"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Change setttings for ODBC data sources"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'all done by the help of other REGEDIT samples on PSC (NO code IS stolen by any of these projects..)
'Done By Kayhan Tanriseven....Any Comments to kayahan@programmner.net NOT to Feedback area because I dont
'think I will be checking it everyday...
'   Notes:
'
'
'   Code Status:
' VB4 32 bit,VB5,VB6
'

'

Option Explicit

Dim mFormActivated As Boolean

' Tab IDs for the "top level"
Const mTAB_ID_Registry = 1
Const mTAB_ID_Run = 2
Const mTAB_ID_AppPaths = 3
Const mTAB_ID_Extensions = 4
Const mTAB_ID_ODBC = 5
Const mTAB_ID_ODBC_Change = 6


' Tab IDs for the registry viewing options
Const mTAB_HKEY_ID_CURRENT_USER = 1
Const mTAB_HKEY_ID_LOCAL_MACHINE = 2
Const mTAB_HKEY_ID_CLASSES_ROOT = 3


' Flag for stopping program changes to txtRegistryKey cause a _Change event
Dim mtxtRegistryKeyFreezeChanges As Boolean
Private Sub PopulateDSN()

Dim lcolDSN As New Collection
Dim lobjDSN


  cboDSN.Clear
  ListODBCDataSources lcolDSN
  For Each lobjDSN In lcolDSN
    cboDSN.AddItem lobjDSN
  Next lobjDSN
End Sub

'
'
Private Sub ShowAppPaths()

Dim lRootKey As Long
Dim lInitialKey$

  
  DoEvents

  lRootKey = HKEY_LOCAL_MACHINE
  lInitialKey$ = "Software\Microsoft\Windows\CurrentVersion\App Paths"
    
  SetStatusBarMsg lInitialKey$
  ViewRegistrySubTree lRootKey, lInitialKey$, Me.tvwAppPaths
End Sub

' Show a list of registered extensions
'
Private Sub ShowExtensions()

Dim lhRootKey As Long, lhProductKey As Long, lhExtensionKey As Long
Dim lExtensionKeyIndex As Long
Dim llProductRetVal As Long, lRetVal As Long
Dim lInitialKey$, lExtension$
Dim lNodNew As Node

Dim lProgID$, lAppRetVal As Long

Dim lAttr$
  
  
  DoEvents

  lhRootKey = HKEY_CLASSES_ROOT
  lInitialKey$ = "\"
    
  SetStatusBarMsg lInitialKey$
  Me.tvwExtensions.Nodes.Clear
  
  ' Set lhProductKey to contain the key which contains the sub-keys for each extension
  llProductRetVal = RegOpenKey(lhRootKey, "", KEY_READ, lhProductKey)
  
  ' Loop for file extensions
  lExtensionKeyIndex = 0
  Do While (llProductRetVal = ERROR_SUCCESS)
    llProductRetVal = RegEnumKey(lhProductKey, lExtensionKeyIndex, lExtension$)
    If (llProductRetVal <> ERROR_SUCCESS) Then Exit Do

    If (Len(lExtension$) <= 4) And (left$(lExtension$, 1) = ".") Then
    
      ' Find the application identifier
      lRetVal = RegOpenKey(HKEY_CLASSES_ROOT, lExtension$, KEY_READ, lhExtensionKey)
      If (lRetVal = ERROR_SUCCESS) Then lRetVal = RegQueryValue(lhExtensionKey, "", lProgID$)
      
      If (lRetVal = ERROR_SUCCESS) Then
        ' I have no idea why this is necessary, but the TreeView control doesn't like a key of ".386"
        If (lExtension$ = ".386") Then lExtension$ = "_" & Mid$(lExtension$, 2)
        
        
        Set lNodNew = Me.tvwExtensions.Nodes.Add(, , lExtension$, lExtension$ & " - " & lProgID$)
        lNodNew.Expanded = True
        
        ' Find the attributes for this extension
        If (GetRegistryAppByExtension(lExtension$, lAttr$, False) = ERROR_SUCCESS) Then
          Set lNodNew = Me.tvwExtensions.Nodes.Add(lExtension$, tvwChild, lExtension$ & "\O" & lAttr$, "Open with: " & lAttr$)
          lNodNew.Expanded = False
        End If
        If (GetRegistryAppByExtension(lExtension$, lAttr$, True) = ERROR_SUCCESS) Then
          Set lNodNew = Me.tvwExtensions.Nodes.Add(lExtension$, tvwChild, lExtension$ & "\S" & lAttr$, "Server: " & lAttr$)
          lNodNew.Expanded = False
        End If
      End If
    End If
    
    lExtensionKeyIndex = lExtensionKeyIndex + 1
  Loop
  
  llProductRetVal = RegCloseKey(lhProductKey)
  
End Sub

Private Sub ShowRegistry()

  DoEvents
  tabRegistry_Click
End Sub

' Show the "Run" & "RunOnce" entries
'
' These are string values stored under \Software\Microsoft\Windows\CurtrentVersion\Run
' (or RunOnce)
'
' The root key will be HKEY_CURRENT_USER for most purposes,
' or HKEY_LOCAL_MACHINE for services that will be run before Windows proper starts
'
' The value name is descriptive, but not used by Windows
' As far as the "Run on Startup" behaviour is concerned, only the value itself is relevant
'
' If the #Const ShowAllEntries is True, then this routine will show all "Run" entries stored for any vendor.
' Although the use of these is unknown, they have been noted in some non-Microsoft products
' (Windows doesn't support them and they are presumed to be for the application's own use)
'
Private Sub ShowRun()

#Const ShowAllEntries = False


Dim llRetVal As Long
Dim llVendorRetVal As Long, llProductRetVal As Long
Dim i%, j%

Dim lInitialKey$, lPartKey$, lCurrentKey$
Dim lBaseKey$, lVendorKey$, lProductKey$, lPartRootKey$
  
Dim nodX As Node

Dim lpValueName As String
Dim lpType As Long, lpValue As Variant

Dim lhRootKey As Long, lRootKeyName$, lhCurrentKey As Long
Dim lVendorKeyIndex As Long, lProductKeyIndex As Long, lValueIndex As Long
Dim lhVendorKey As Long, lhProductKey As Long, lhValueKey As Long

Dim lbRootShown As Boolean, lbVendorShown As Boolean, lbProductShown As Boolean
Dim lEntryCount As Long
  
  
  DoEvents
  Me.MousePointer = vbHourglass
  tvwRun.Nodes.Clear
  
  lBaseKey$ = "Software"

  SetStatusBarMsg "Searching Registry..."
  DoEvents

  ' Loop for classes of entry
  For i% = 1 To 3
    Select Case i%
    Case 1
      lPartKey$ = "Run"
    Case 2
      lPartKey$ = "RunOnce"
    Case 3
      lPartKey$ = "RunServicesOnce"
    End Select
    
    Set nodX = tvwRun.Nodes.Add(, , lPartKey$, lPartKey$, tvwTextOnly)
    nodX.Expanded = True
    lEntryCount = 0
    
    ' Loop for root keys
    For j% = 1 To 2
      Select Case j%
      Case 1
        lhRootKey = HKEY_LOCAL_MACHINE
        lRootKeyName$ = "Before Windows start-up"
      Case 2
        lhRootKey = HKEY_CURRENT_USER
        lRootKeyName$ = "Normal"
      End Select
      lbRootShown = False
      
      
#If ShowAllEntries Then
      ' Loop for vendors
      llVendorRetVal = RegOpenKey(lhRootKey, lBaseKey$, KEY_READ, lhVendorKey)
      lVendorKeyIndex = 0
      Do While (llVendorRetVal = ERROR_SUCCESS)
        llVendorRetVal = RegEnumKey(lhVendorKey, lVendorKeyIndex, lVendorKey$)
        If (llVendorRetVal <> ERROR_SUCCESS) Then Exit Do
        
        lbVendorShown = False
        
        ' Skip hives which aren't vendor products
        If (lVendorKey$ = "Classes") Then GoTo SkipVendor
        
        ' Loop for products
        llProductRetVal = RegOpenKey(lhRootKey, lBaseKey$ & "\" & lVendorKey$, KEY_READ, lhProductKey)
        lProductKeyIndex = 0
        Do While (llProductRetVal = ERROR_SUCCESS)
          llProductRetVal = RegEnumKey(lhProductKey, lProductKeyIndex, lProductKey$)
          If (llProductRetVal <> ERROR_SUCCESS) Then Exit Do
          
          lbProductShown = False
          lInitialKey$ = lBaseKey$ & "\" & lVendorKey$ & "\" & lProductKey$ & "\CurrentVersion\"
          
            
#Else
          lInitialKey$ = lBaseKey$ & "\Microsoft\Windows\CurrentVersion\"
#End If

          SetStatusBarMsg lInitialKey$
          
          lCurrentKey$ = lInitialKey$ & lPartKey$
          llRetVal = RegOpenKey(lhRootKey, lCurrentKey$, KEY_READ, lhCurrentKey)
          If llRetVal = ERROR_SUCCESS Then
                    
            lValueIndex = 0
            llRetVal = RegEnumValue(lhCurrentKey, lValueIndex, lpValueName, lpType, lpValue)
            
            ' Loop for entries
            Do While (llRetVal = ERROR_SUCCESS)
              ' An entry has been found
              
              ' Make sure that the necessary parent keys are in place
              If Not lbRootShown Then
                lPartRootKey$ = lPartKey$ & "\" & lRootKeyName$
                Set nodX = tvwRun.Nodes.Add(lPartKey$, tvwChild, lPartRootKey$, lRootKeyName$, tvwTextOnly)
                nodX.Expanded = True
                lbRootShown = True
              End If
#If ShowAllEntries Then
              If Not lbVendorShown Then
                Set nodX = tvwRun.Nodes.Add(lPartRootKey$, tvwChild, lPartRootKey$ & "\" & lVendorKey$, lVendorKey$, tvwTextOnly)
                nodX.Expanded = True
                lbVendorShown = True
              End If
              If Not lbProductShown Then
                Set nodX = tvwRun.Nodes.Add(lPartRootKey$ & "\" & lVendorKey$, tvwChild, lPartRootKey$ & "\" & lVendorKey$ & "\" & lProductKey$, lProductKey$, tvwTextOnly)
                nodX.Expanded = True
                lbProductShown = True
              End If
              Set nodX = tvwRun.Nodes.Add(lPartRootKey$ & "\" & lVendorKey$ & "\" & lProductKey$, tvwChild, lPartRootKey$ & "\" & lVendorKey$ & "\" & lProductKey$ & "\" & lpValueName, lpValueName & " : " & lpValue, tvwTextOnly)
#Else
              Set nodX = tvwRun.Nodes.Add(lPartRootKey$, tvwChild, lPartRootKey$ & "\" & lpValueName, lpValueName & " : " & lpValue, tvwTextOnly)
#End If
              
              nodX.Expanded = True
              lEntryCount = lEntryCount + 1
              
              ' Get the next one
              lValueIndex = lValueIndex + 1
              llRetVal = RegEnumValue(lhCurrentKey, lValueIndex, lpValueName, lpType, lpValue)
            Loop
            If (259 = llRetVal) Then llRetVal = ERROR_SUCCESS
          End If
          
#If ShowAllEntries Then
          lProductKeyIndex = lProductKeyIndex + 1
        Loop
        
SkipVendor:
        lVendorKeyIndex = lVendorKeyIndex + 1
      Loop
#Else
#End If

    Next j%
    
    ' If no entries for this service, then explain
    If (lEntryCount = 0) Then
      Set nodX = tvwRun.Nodes.Add(lPartKey$, tvwChild, lPartKey$ & "\", "< No Entries >", tvwTextOnly)
      nodX.Expanded = True
    End If
  Next i%
  
  SetStatusBarMsg ""
  Me.MousePointer = vbDefault
End Sub

Public Sub SetStatusBarMsg(pMsg$)

  sbrMain.Panels(1).Text = pMsg$
  DoEvents
End Sub


Private Sub cmdAppPathAdd_Click()

Dim lRetVal As Long


  lRetVal = SaveRegistryAppPath(txtAppPathName, txtAppPathEXEPath, txtAppPathPath)
  If (lRetVal <> 0) Then MsgBox "Error " & CStr(lRetVal) & ":", vbCritical, "Error"
  
  ' Refresh to show them
  DoEvents
  tabView_Click
End Sub

Private Sub cmdCreateODBC_Click()

Dim llRetVal As Long


  If Len(Me.cboDSN) = 0 Then
    Me.cboDSN.SetFocus
    Exit Sub
  End If

  If Len(Me.cboServer) = 0 Then
    Me.cboServer.SetFocus
    Exit Sub
  End If

  If Len(Me.txtDatabase) = 0 Then
    Me.txtDatabase.SetFocus
    Exit Sub
  End If


  llRetVal = CreateODBCDataSource(cboDSN, "", cboServer, txtDatabase, "")
  If llRetVal <> ERROR_SUCCESS Then
    MsgBox "CreateODBCDataSource() failed : " & Str$(llRetVal), vbExclamation, ""
  
  End If
  PopulateDSN
End Sub

Private Sub cmdDeleteODBC_Click()

Dim llRetVal As Long


  If Len(Me.cboDSN) = 0 Then
    Me.cboDSN.SetFocus
    Exit Sub
  End If

  
  llRetVal = DeleteODBCDataSource(cboDSN)
  If llRetVal <> ERROR_SUCCESS Then
    MsgBox "DeleteODBCDataSource() failed : " & Str$(llRetVal), vbExclamation, ""
  End If
  
  PopulateDSN
End Sub


Private Sub cmdFindAppByExtension_Click()

Dim lRetVal As Long
Dim lAppName$


  lRetVal = GetRegistryAppByExtension(txtExtension, lAppName$, (optExtensionFindAppServer))
  If (lRetVal = ERROR_SUCCESS) Then
    txtExtensionApp = lAppName$
  ElseIf (lRetVal = ERROR_BADKEY) Or (lRetVal = 2) Then
    txtExtensionApp = "No " & IIf((optExtensionFindAppServer), "server", "application") & " registered for extension '" & (txtExtension) & "'"
  Else
    txtExtensionApp = "Error " & Str$(lRetVal)
  End If
End Sub

' This is just a dummy
'
Private Sub cmdRegister_Click()

Dim lRetVal As Long
  
  lRetVal = SaveRegistryAssociation(txtExtension, "Prog-ID", "C:\WINDOWS\NOTEPAD.EXE %1", "C:\WINDOWS\NOTEPAD.EXE", 3)
  If (lRetVal <> 0) Then MsgBox "Error " & CStr(lRetVal) & ":", vbCritical, "Error"
  
  ' Refresh to show them
  DoEvents
  tabView_Click
End Sub

Private Sub cmdRemoveApp_Click()

Dim lRetVal As Long

  lRetVal = SaveRegistryDontRunAppAfterRestart()
  If (lRetVal <> 0) Then MsgBox "Error " & CStr(lRetVal) & ":", vbCritical, "Error"
  
  ' Refresh to show them
  DoEvents
  tabView_Click
End Sub

Private Sub cmdRestartApp_Click()

Dim lRetVal As Long

  lRetVal = SaveRegistryRunAppAfterRestart(txtRestartParameters)
  If (lRetVal <> 0) Then MsgBox "Error " & CStr(lRetVal) & ":", vbCritical, "Error"
  
  ' Refresh to show them
  DoEvents
  tabView_Click
End Sub

Private Sub cmdRunAdd_Click()

Dim lRetVal As Long

  lRetVal = SaveRegistryRunAtStartup((txtRunValueName), (txtRunValue), (chkRunOnce = vbChecked), (chkRunStartup = vbChecked))
  If (lRetVal <> 0) Then MsgBox "Error " & CStr(lRetVal) & ":", vbCritical, "Error"
  
  ' Refresh to show them
  DoEvents
  tabView_Click
End Sub


Private Sub Form_Activate()
  
  DoEvents
  If Not mFormActivated Then
    mFormActivated = True

    ' Controls to default values
    optRegistryKey(1) = True

    ' Run
    chkRunOnce = vbChecked
    txtRunValueName = "Registry Viewer"
    txtRunValue = App.Path & "\" & App.EXEName & ".EXE"

    ' App paths
    txtAppPathName = App.EXEName & ".EXE"
    txtAppPathEXEPath = App.Path & "\" & App.EXEName & ".EXE"
    txtAppPathPath = App.Path

    ' Extensions
    optExtensionFindAppServer = True
    
    ' ODBC
    PopulateDSN


    ' Show the Registry first
    picView(mTAB_ID_Registry).ZOrder vbBringToFront
    
    
    DoEvents
    tabView_Click
    
  End If
End Sub


Private Sub Form_Load()

  mFormActivated = False
  If (LoadFormPosition(Me) <> ERROR_SUCCESS) Then CentreForm Me
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim lRetVal As Long


  ' Make our application "sticky", if it is closed by Windows
  If (UnloadMode = vbAppWindows) Then
    ' If Windows is forcing the closure, then make this application "sticky"
    lRetVal = SaveRegistryRunAppAfterRestart(txtRestartParameters)
    ' Ignore errors - we're probably well broken by now
  Else
    ' If we close for some other reason, then _don't_ restart with Windows
    ' We need to save this explicitly in case of the (rare) case where
    ' a _QueryUnload from Windows to our app is later rejected by another app
    ' then we manually close our app before trying to close Windows again.
    lRetVal = SaveRegistryDontRunAppAfterRestart()
  End If
End Sub

' Each Tab has a (smaller) picture control placed upon it
' the ZOrder of the picture being used to show the tabs
'
' Further controls are contained by the Picture controls
' and may extend to the full extent of the Picture's area
'
Private Sub Form_Resize()

Dim i%
Dim lXTemp!, lYTemp!


  If (Me.WindowState <> vbMinimized) Then
    For i% = picView.LBound() To picView.UBound()
      picView(i%).left = (tabView.left * 2)
    Next i%
  
    tabView.Width = Me.ScaleWidth - (tabView.left * 2)
    lXTemp! = tabView.Width - (picView(mTAB_ID_ODBC).left + 120)
    
    tabView.Height = Me.ScaleHeight - ((tabView.top * 3) + sbrMain.Height)
    lYTemp! = tabView.Height - (picView(mTAB_ID_ODBC).top + 120)
    For i% = picView.LBound() To picView.UBound()
      With picView(i%)
        .Width = lXTemp!
        .Height = lYTemp!
        End With
    Next i%
    With tvwODBC
      .Width = picView(mTAB_ID_ODBC).ScaleWidth
      .Height = picView(mTAB_ID_ODBC).ScaleHeight
      End With
      
    ' Registry
    With tabRegistry
      .Width = picView(mTAB_ID_Registry).ScaleWidth
      .Height = picView(mTAB_ID_Registry).ScaleHeight
      End With
    With fraRegistryKey
      .Width = picView(mTAB_ID_Registry).ScaleWidth - (2 * .left)
      .top = picView(mTAB_ID_Registry).ScaleHeight - (.Height + 120)
      End With
    With tvwRegistry
      .Width = picView(mTAB_ID_Registry).ScaleWidth - (2 * .left)
      .Height = fraRegistryKey.top - (.top + 120)
      End With
      
    ' Run
    fraRun.top = picView(mTAB_ID_Run).ScaleHeight - (fraRun.Height + 120)
    fraRunThisApplication.top = fraRun.top
    fraRun.Width = (picView(mTAB_ID_Run).ScaleWidth - 120) / 2
    fraRunThisApplication.Width = fraRun.Width
    fraRunThisApplication.left = fraRun.left + fraRun.Width + 120
    With tvwRun
      .Width = picView(mTAB_ID_Run).ScaleWidth
      .Height = fraRun.top - 120
      End With
      
    ' App Paths
    fraAppPaths.top = picView(mTAB_ID_AppPaths).ScaleHeight - (fraAppPaths.Height + 120)
    fraAppPaths.Width = (picView(mTAB_ID_AppPaths).ScaleWidth - 120)
    txtAppPathEXEPath.Width = fraAppPaths.Width - (txtAppPathEXEPath.left + 120)
    txtAppPathPath.Width = txtAppPathEXEPath.Width
    With tvwAppPaths
      .Width = picView(mTAB_ID_AppPaths).ScaleWidth
      .Height = fraAppPaths.top - 120
      End With
      
    ' Extensions
    fraExtensions.top = picView(mTAB_ID_Extensions).ScaleHeight - (fraExtensions.Height + 120)
    fraExtensions.Width = (picView(mTAB_ID_Extensions).ScaleWidth - 120)
'    txtAppPathEXEPath.Width = fraExtensions.Width - (txtAppPathEXEPath.left + 120)
'    txtAppPathPath.Width = txtAppPathEXEPath.Width
    With tvwExtensions
      .Width = picView(mTAB_ID_Extensions).ScaleWidth
      .Height = fraExtensions.top - 120
      End With
    
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim llRetVal As Long

  llRetVal = SaveFormPosition(Me)
End Sub

Private Sub optExtensionFindAppOpenWith_Click()

  If (Len(txtExtensionApp) > 0) Then cmdFindAppByExtension_Click
End Sub

Private Sub optExtensionFindAppServer_Click()

  If (Len(txtExtensionApp) > 0) Then cmdFindAppByExtension_Click
End Sub


Private Sub optRegistryKey_Click(Index As Integer)

  mtxtRegistryKeyFreezeChanges = True
  Select Case Index
  Case 1
    txtRegistryKey = "Software\Microsoft"
  Case 2
    txtRegistryKey = GetRegistryKeyNameForApp$()
  End Select
  mtxtRegistryKeyFreezeChanges = False
  tabRegistry_Click
End Sub


Private Sub tabRegistry_Click()

Dim lRootKey As Long
Dim lInitialKey$

  If (tabRegistry.SelectedItem.Index > 0) And (tabRegistry.SelectedItem.Index <= 3) Then
    Select Case tabRegistry.SelectedItem.Index
    Case mTAB_HKEY_ID_CURRENT_USER
      lRootKey = HKEY_CURRENT_USER
      lInitialKey$ = txtRegistryKey
    Case mTAB_HKEY_ID_LOCAL_MACHINE
      lRootKey = HKEY_LOCAL_MACHINE
      lInitialKey$ = txtRegistryKey
    Case mTAB_HKEY_ID_CLASSES_ROOT
      lRootKey = HKEY_CLASSES_ROOT
      lInitialKey$ = txtRegistryKey
    End Select
    
    ' If the Registry tab isn't selected, then bale out
    If (tabView.SelectedItem.Index <> mTAB_ID_Registry) Then Exit Sub
    
    Me.MousePointer = vbHourglass
    SetStatusBarMsg lInitialKey$
    ViewRegistrySubTree lRootKey, lInitialKey$, Me.tvwRegistry
    Me.MousePointer = vbDefault
  End If
End Sub

' A new "top level" function has been seelcted
'
Private Sub tabView_Click()

  If (tabView.SelectedItem.Index > 0) And (tabView.SelectedItem.Index <= picView.Count) Then
    picView(tabView.SelectedItem.Index).ZOrder vbBringToFront
    
    Me.MousePointer = vbHourglass
    Select Case tabView.SelectedItem.Index
    Case mTAB_ID_Registry
      ShowRegistry
    
    Case mTAB_ID_Run
      ShowRun
    
    Case mTAB_ID_AppPaths
      ShowAppPaths
    
    Case mTAB_ID_Extensions
      ShowExtensions
    
    Case mTAB_ID_ODBC
      ShowODBC Me.tvwODBC
      
    End Select
    Me.MousePointer = vbDefault
  End If
End Sub




Private Sub tvwExtensions_DblClick()
  DoEvents
  optExtensionFindAppOpenWith = True
  cmdFindAppByExtension_Click
End Sub

' Node selected, so display details for this extension
'
Private Sub tvwExtensions_NodeClick(ByVal Node As Node)
  DoEvents
  txtExtension = Node.Key
End Sub

Private Sub txtExtension_Change()
  txtExtensionApp = ""
End Sub


Private Sub txtRegistryKey_Change()

  If Not mtxtRegistryKeyFreezeChanges Then
    optRegistryKey(3) = True
  End If
End Sub


