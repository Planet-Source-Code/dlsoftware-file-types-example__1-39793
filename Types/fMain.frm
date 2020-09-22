VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1200
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox pLarge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5040
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2760
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   4440
   End
   Begin MSComctlLib.ListView lvwFile 
      Height          =   5010
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Right click for additional information"
      Top             =   0
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   8837
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FileExt"
         Text            =   "File Extension"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "FileType"
         Text            =   "File Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuOfCommands 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuOptions 
         Caption         =   "Display Additional Information"
         Index           =   1
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Backup Registry Entry (.reg)"
         Index           =   2
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Exit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Const HKEY_CLASSES_ROOT As Long = &H80000000
    Private Const GOOD_RETURN_CODE As Long = 0
    Private Const STARTS_WITH_A_PERIOD As Long = 46
    Private Const MAX_PATH_LENGTH As Long = 260
    Private Const REG_SZ = (1)
    Private Const REG_EXPAND_SZ = (2)
    Private Const ILD_TRANSPARENT = &H1
    Private Const STANDARD_RIGHTS_READ As Long = &H20000
    Private Const KEY_QUERY_VALUE As Long = &H1
    Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
    Private Const KEY_NOTIFY As Long = &H10
    Private Const SYNCHRONIZE As Long = &H100000
    Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    Private Const SH_USEFILEATTRIBUTES As Long = &H10
    Private Const SH_TYPENAME As Long = &H400
    Private Const SH_DISPLAYNAME = &H200
    Private Const SH_EXETYPE = &H2000
    Private Const SH_SYSICONINDEX = &H4000
    Private Const SH_LARGEICON = &H0
    Private Const SH_SMALLICON = &H1
    Private Const SH_SHELLICONSIZE = &H4
    Private Const FILE_ATTRIBUTE_NORMAL = &H80
    Private Const BASIC_SH_FLAGS = SH_TYPENAME Or SH_SHELLICONSIZE Or SH_SYSICONINDEX Or SH_DISPLAYNAME Or SH_EXETYPE
    Private Type FILETIME: dwLowDateTime As Long: dwHighDateTime As Long: End Type
    Private Type SHFILEINFO: hIcon As Long: iIcon As Long: dwAttributes As Long: szDisplayName As String * MAX_PATH_LENGTH: szTypeName As String * 80: End Type
    Private Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
    Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
    Private Declare Function RegEnumKeyEx Lib "advapi32" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
    Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
    Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
    Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Private shinfo As SHFILEINFO
    Private bResize As Boolean
    Private lSortCol As Long
    Private iPos As Integer
    Private lIcon As Long
    Private lIco2 As Long
    Private sBarValue As Long
    Private sCount As Long
    Private sw As Long
    Private lResult As Long
    Private lResultEnumKey As Long
    Private lrc As Long
    Private rc1 As Long
    Private rc2 As Long
    Private rc3 As Long
    Private cch As Long
    Private sActionValue As String
    Private TheRecord As String
    Private sActionCommand As String
    Private sActionKey As String
    Private sAction As String
    Private sTitle As String
    Private lType As Long
    Private vValue As Variant
    Private sValue As String
    Private sKey As String
    Private lRegKeyIndex As Long
    Private sImageList1Key As String
    Private sFileTypeName As String
    Private Buffer As String
    Private CurDir As String
    Private q As String
    Private sFileExtension As String
    Private sExtension
    Private sRegSubkey As String * MAX_PATH_LENGTH
    Private sRegKeyClass As String * MAX_PATH_LENGTH
    Private ftime As FILETIME
    Private lvi As ListItem
    Private iSmall As ListImage
    Private iLarge As ListImage

Private Sub Form_Load()
    On Error Resume Next
    With lvwFile
        .SmallIcons = ImageList1
        .ColumnHeaders("FileExt").Width = .Width * 0.2
        .ColumnHeaders("FileType").Width = .Width * 0.8
    End With
    Buffer = Space(255)
    lrc = GetCurrentDirectory(Len(Buffer), Buffer)
    CurDir = TrimNull(Buffer)
    q = """"
    fMain.Visible = False
    fSplash.Show , fMain
    DoEvents
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If bResize Then Exit Sub
    If Me.WindowState = vbMinimized Then Exit Sub
    bResize = True
    If Me.Width <= 7440 Then Me.Width = 7440
    If Me.Height <= 5010 Then Me.Height = 5010
    With lvwFile
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top
    End With
    bResize = False
End Sub

Private Sub lvwFile_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    On Error Resume Next
    lSortCol = ColumnHeader.Index - 1
    With lvwFile
        If .SortKey = lSortCol Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortKey = lSortCol
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub

Private Sub lvwFile_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    sFileExtension = Item.Text
    sFileTypeName = lvwFile.SelectedItem.SubItems(1)
End Sub

Private Sub lvwFile_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If (Button = vbRightButton) Then
       lvwFile.SetFocus
       Me.PopupMenu mnuOfCommands
   End If
End Sub

Private Sub TrimValue()
    On Error Resume Next
    If vValue = "" Then
       sValue = ""
    Else
       sValue = vValue
       sValue = TrimNull(sValue)
    End If
End Sub

Private Sub mnuOptions_Click(Index As Integer)
    On Error Resume Next
    If Index = 3 Then Exit Sub
    If Index = 2 Then GoTo RegistryBackup
    sFileExtension = "." & sFileExtension
    fValues.Text1.Text = sFileExtension
    rc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sFileExtension, 0, KEY_READ, lResult)
    If rc1 = 0 Then
       rc2 = QueryValueEx(lResult, vbNullString, vValue)
       TrimValue
       fValues.Text2.Text = sValue
       sKey = sValue
       rc2 = QueryValueEx(lResult, "content type", vValue)
       TrimValue
       fValues.Text3.Text = sValue
    Else
       fValues.Text2.Text = ""
       fValues.Text3.Text = ""
    End If
    RegCloseKey HKEY_CLASSES_ROOT
    If sKey = "" Then
       sKey = sFileExtension
    End If
    rc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKey & "\DefaultIcon", 0, KEY_READ, lResult)
    If rc1 = 0 Then
       rc2 = QueryValueEx(lResult, vbNullString, vValue)
       TrimValue
       fValues.Text4.Text = sValue
    Else
       fValues.Text4.Text = ""
    End If
    RegCloseKey HKEY_CLASSES_ROOT
    With fValues.lvwFile
        .ColumnHeaders("Action").Width = .Width * 0.2
        .ColumnHeaders("Command").Width = .Width * 1.3
    End With
    rc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKey & "\shell", 0, KEY_READ, lResult)
    If rc1 = 0 Then
       rc2 = QueryValueEx(lResult, vbNullString, vValue)
       TrimValue
       fValues.Text5.Text = sValue
    Else
       fValues.Text5.Text = ""
    End If
    lRegKeyIndex = 0
    Do While RegEnumKeyEx(lResult, lRegKeyIndex, sRegSubkey, MAX_PATH_LENGTH, 0, sRegKeyClass, MAX_PATH_LENGTH, ftime) = GOOD_RETURN_CODE
       sAction = TrimNull(sRegSubkey)
       sActionKey = sKey & "\shell\" & sAction
       rc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sActionKey, 0, KEY_READ, lResultEnumKey)
       If rc1 = 0 Then
          rc2 = QueryValueEx(lResultEnumKey, vbNullString, vValue)
          TrimValue
          sActionValue = sValue
       Else
          sActionValue = ""
       End If
       If sActionValue = "" Then
          sActionValue = sAction
       End If
       sActionKey = sKey & "\shell\" & sAction & "\command"
       rc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sActionKey, 0, KEY_READ, lResultEnumKey)
       If rc1 = 0 Then
          rc2 = QueryValueEx(lResultEnumKey, vbNullString, vValue)
          TrimValue
          sActionCommand = sValue
       Else
          sActionCommand = ""
       End If
       Set lvi = fValues.lvwFile.ListItems.Add(, , sActionValue)
           lvi.SubItems(1) = sActionCommand
       lRegKeyIndex = lRegKeyIndex + 1
    Loop
    RegCloseKey HKEY_CLASSES_ROOT
    sFileExtension = Right(sFileExtension, Len(sFileExtension) - 1)
    sImageList1Key = "#" & sFileExtension & "#"
    pSmall.Picture = ImageList1.ListImages(sImageList1Key).Picture
    fValues.pSmall.Picture = pSmall.Picture
    pLarge.Picture = ImageList2.ListImages(sImageList1Key).Picture
    fValues.pLarge.Picture = pLarge.Picture
    fValues.Text6.Text = sFileTypeName
    fValues.Show , fMain
    Refresh
    Exit Sub
    
RegistryBackup:
    sExtension = UCase(sFileExtension)
    sFileExtension = "." & sFileExtension
    rc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sFileExtension, 0, KEY_READ, lResult)
    If rc1 = 0 Then
       rc2 = QueryValueEx(lResult, vbNullString, vValue)
       TrimValue
       sKey = sValue
    Else
       sKey = sFileExtension
    End If
    RegCloseKey HKEY_CLASSES_ROOT
    Kill CurDir & "\" & sExtension & "#1.reg"
    Kill CurDir & "\" & sExtension & "#2.reg"
    sValue = "regedit.exe /e" & q & " " & q & CurDir & "\" & sExtension & "#1.reg" & q & " " & q & "HKEY_CLASSES_ROOT\" & sFileExtension & q
    rc1 = Shell(sValue, vbNormalFocus)
    fBackup.Text1.Text = CurDir & "\" & sExtension & "#1.reg"
    fBackup.Text2.Text = ""
    If sKey <> sFileExtension And sKey <> "" Then
       sValue = "regedit.exe /e" & q & " " & q & CurDir & "\" & sExtension & "#2.reg" & q & " " & q & "HKEY_CLASSES_ROOT\" & sKey & q
       rc1 = Shell(sValue, vbNormalFocus)
       fBackup.Text2.Text = CurDir & "\" & sExtension & "#2.reg"
    End If
    fBackup.Show , fMain
    Refresh
End Sub

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
      On Error GoTo QueryValueExError
      rc3 = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
      If rc3 <> 0 Then Error 5
      Select Case lType
             Case REG_SZ, REG_EXPAND_SZ:
                  sValue = String(cch, 0)
                  lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
                  If lrc = 0 Then
                     vValue = Left(sValue, cch)
                  Else
                     vValue = Empty
                  End If
             Case Else
                  vValue = Empty
                  lrc = -1
      End Select
QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:
      vValue = Empty
      Resume QueryValueExExit
End Function

Private Sub Timer1_Timer()
    On Error Resume Next
    Timer1.Enabled = False
    Screen.MousePointer = vbArrowHourglass
    fSplash.pBar1.Value = 0
    DoEvents
    Do While RegEnumKeyEx(HKEY_CLASSES_ROOT, lRegKeyIndex, sRegSubkey, MAX_PATH_LENGTH, 0, sRegKeyClass, MAX_PATH_LENGTH, ftime) = GOOD_RETURN_CODE
       If Asc(sRegSubkey) = STARTS_WITH_A_PERIOD Then
          sCount = sCount + 1
          sBarValue = ((sCount / 600) * 100)
          If sBarValue > 100 Then
             sBarValue = 100
          End If
          fSplash.Label3.Caption = sBarValue & "%"
          fSplash.pBar1.Value = sBarValue
          lIco2 = SHGetFileInfo(sRegSubkey, FILE_ATTRIBUTE_NORMAL, shinfo, Len(shinfo), SH_USEFILEATTRIBUTES Or BASIC_SH_FLAGS Or SH_LARGEICON)
          lIcon = SHGetFileInfo(sRegSubkey, FILE_ATTRIBUTE_NORMAL, shinfo, Len(shinfo), SH_USEFILEATTRIBUTES Or BASIC_SH_FLAGS Or SH_SMALLICON)
          sFileTypeName = TrimNull(shinfo.szTypeName)
          sFileExtension = TrimNull(sRegSubkey)
          sFileExtension = Right(sFileExtension, Len(sFileExtension) - 1)
          pSmall.Picture = LoadPicture()
          Call ImageList_Draw(lIcon, shinfo.iIcon, pSmall.hDC, 0, 0, ILD_TRANSPARENT)
          pSmall.Picture = pSmall.Image
          pLarge.Picture = LoadPicture()
          Call ImageList_Draw(lIco2, shinfo.iIcon, pLarge.hDC, 0, 0, ILD_TRANSPARENT)
          pLarge.Picture = pLarge.Image
          sImageList1Key = "#" & sFileExtension & "#"
          Set iSmall = ImageList1.ListImages.Add(, sImageList1Key, pSmall.Picture)
          Set iLarge = ImageList2.ListImages.Add(, sImageList1Key, pLarge.Picture)
          Set lvi = lvwFile.ListItems.Add(, , sFileExtension)
                    lvi.SmallIcon = ImageList1.ListImages(sImageList1Key).Key
                    lvi.SubItems(1) = sFileTypeName
          DoEvents
       End If
       lRegKeyIndex = lRegKeyIndex + 1
    Loop
    sTitle = "File Types Example   -   " & CStr(sCount) & " registered file extensions"
    fMain.Caption = sTitle
    fMain.Visible = True
    Unload fSplash
    Set fSplash = Nothing
    Screen.MousePointer = vbDefault
    DoEvents
End Sub

Private Function TrimNull(startstr As String) As String
    On Error Resume Next
    iPos = InStr(startstr, Chr$(0))
    If iPos Then
       TrimNull = Left$(startstr, iPos - 1)
       Exit Function
    End If
    TrimNull = startstr
End Function

