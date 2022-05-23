VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5280
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Main"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215.568
   ScaleMode       =   0  'User
   ScaleWidth      =   4387.096
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ListBox lstRenamed 
      Height          =   2400
      ItemData        =   "Main.frx":0CCA
      Left            =   3120
      List            =   "Main.frx":0CD1
      TabIndex        =   17
      Top             =   2280
      Width           =   2895
   End
   Begin VB.ListBox lstOriginal 
      Height          =   2400
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton btnRename 
      Caption         =   "Rena&me"
      Height          =   375
      Left            =   4680
      TabIndex        =   19
      Top             =   4800
      Width           =   1335
   End
   Begin TabDlg.SSTab tabCounter 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2778
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Add Text"
      TabPicture(0)   =   "Main.frx":0CE1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAddBefore"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAddAfter"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtAddBefore"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAddAfter"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Replace Text"
      TabPicture(1)   =   "Main.frx":0CFD
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtReplace"
      Tab(1).Control(1)=   "txtFind"
      Tab(1).Control(2)=   "lblReplace"
      Tab(1).Control(3)=   "lblFind"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "C&ounter"
      TabPicture(2)   =   "Main.frx":0D19
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkAfterName"
      Tab(2).Control(1)=   "chkBeforeName"
      Tab(2).Control(2)=   "txtStart"
      Tab(2).Control(3)=   "updown"
      Tab(2).Control(4)=   "lblStart"
      Tab(2).ControlCount=   5
      Begin VB.CheckBox chkAfterName 
         Caption         =   "A&fter name"
         Height          =   255
         Left            =   -70800
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkBeforeName 
         Caption         =   "&Before name"
         Height          =   255
         Left            =   -72240
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtStart 
         Height          =   405
         Left            =   -73680
         TabIndex        =   10
         Text            =   "0"
         Top             =   480
         Width           =   840
      End
      Begin MSComCtl2.UpDown updown 
         Height          =   405
         Left            =   -72840
         TabIndex        =   11
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   714
         _Version        =   393216
         BuddyControl    =   "txtStart"
         BuddyDispid     =   196615
         OrigLeft        =   2040
         OrigTop         =   600
         OrigRight       =   2295
         OrigBottom      =   975
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtReplace 
         Height          =   375
         Left            =   -73680
         TabIndex        =   8
         Top             =   980
         Width           =   4455
      End
      Begin VB.TextBox txtFind 
         Height          =   375
         Left            =   -73680
         TabIndex        =   6
         Top             =   500
         Width           =   4455
      End
      Begin VB.TextBox txtAddAfter 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   980
         Width           =   4455
      End
      Begin VB.TextBox txtAddBefore 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   500
         Width           =   4455
      End
      Begin VB.Label lblStart 
         Caption         =   "Start:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblReplace 
         Caption         =   "Replace:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   1100
         Width           =   975
      End
      Begin VB.Label lblFind 
         Caption         =   "Find:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   5
         Top             =   620
         Width           =   975
      End
      Begin VB.Label lblAddAfter 
         Caption         =   "After name:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1100
         Width           =   975
      End
      Begin VB.Label lblAddBefore 
         Caption         =   "Before name:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   620
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   9240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblRenamed 
      Caption         =   "Renamed:"
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblOriginal 
      Caption         =   "Original:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuWebSite 
         Caption         =   "Visit &Web Site"
      End
      Begin VB.Menu mnuHelpSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private FilesOpened As Boolean

Private RootDirectory As String
Private FilesToRename() As FileToRename

Private Sub OpenFiles(DirectoryName As String, FileNames() As String)
    Dim Files() As FileToRename
    ReDim Files(UBound(FileNames))
    
    Dim i As Integer
    
    SortArray FileNames
    
    For i = 0 To UBound(FileNames)
        Files(i) = OpenFile(DirectoryName, CStr(FileNames(i)))
    Next
    
    mnuOpen.Enabled = False
    mnuClose.Enabled = True
    FilesOpened = True
    
    RootDirectory = DirectoryName
    FilesToRename = Files
    
    ReloadOriginalFileList
    ReloadRenamedFileList
End Sub

Private Sub CloseFiles()
    mnuOpen.Enabled = True
    mnuClose.Enabled = False
    FilesOpened = False
    
    RootDirectory = ""
    Erase FilesToRename
    
    ReloadOriginalFileList
    ReloadRenamedFileList
End Sub

Private Sub ClearFields()
    txtAddBefore.Text = ""
    txtAddAfter.Text = ""
    
    txtFind.Text = ""
    txtReplace.Text = ""
    
    txtStart.Text = 0
    chkBeforeName.Value = False
    chkAfterName.Value = False
End Sub

Private Sub ReloadOriginalFileList()
    lstOriginal.Clear
    
    If FilesOpened Then
        Dim i As Integer
        
        For i = 0 To UBound(FilesToRename)
            Dim File As FileToRename: File = FilesToRename(i)
            
            If File.FileExists Then
                lstOriginal.AddItem File.OldFileName
            End If
        Next
    End If
End Sub

Private Sub ReloadRenamedFileList()
    lstRenamed.Clear
    
    If FilesOpened Then
        Dim i As Integer
        
        For i = 0 To UBound(FilesToRename)
            Dim File As FileToRename: File = FilesToRename(i)
            
            If File.FileExists Then
                lstRenamed.AddItem File.NewFileName
            End If
        Next
    End If
End Sub

' Selection sort
Private Sub SortArray(ByRef List As Variant)
    Dim SmallestIndex As Integer
    Dim SmallestValue As Variant
    
    Dim Min As Integer: Min = LBound(List)
    Dim Max As Integer: Max = UBound(List) - 1
    
    Dim i As Integer
    Dim j As Integer
    
    For i = Min To Max
        SmallestValue = List(i)
        SmallestIndex = i
        
        For j = i + 1 To Max
            If List(j) < SmallestValue Then
                SmallestValue = List(j)
                SmallestIndex = j
            End If
        Next j
        
        List(SmallestIndex) = List(i)
        List(i) = SmallestValue
    Next i
End Sub

Private Sub EditText()
    If FilesOpened Then
        Dim i As Integer
        
        For i = 0 To UBound(FilesToRename)
            Dim File As FileToRename: File = FilesToRename(i)
            
            If File.FileExists Then
                Dim NewFileName As String: NewFileName = File.OldFileName
                
                NewFileName = txtAddBefore.Text & NewFileName & txtAddAfter.Text
                NewFileName = Replace(NewFileName, txtFind.Text, txtReplace.Text)
                
                If chkBeforeName.Value Then
                    NewFileName = (txtStart.Text + i) & NewFileName
                End If
                
                If chkAfterName.Value Then
                    NewFileName = NewFileName & (txtStart.Text + i)
                End If
                
                FilesToRename(i).NewFileName = NewFileName
            End If
        Next
        
        ReloadRenamedFileList
    End If
End Sub

Private Sub RenameFiles()
    On Error GoTo RENAME_ERROR
    
    If FilesOpened Then
        Dim i As Integer
        
        For i = 0 To UBound(FilesToRename)
            Dim File As FileToRename: File = FilesToRename(i)
            
            With File
                If .FileExists Then
                    Dim FullPath As String: FullPath = GetFullPath(.DirectoryName, .NewFileName)
                    
                    If Not CheckIfFileExists(FullPath) Then
                        RenameFile File
                        
                        FilesToRename(i).OldFileName = .NewFileName
                    Else
                        MsgBox "A file named """ & .NewFileName & """ already exists in the directory.", _
                            vbCritical Or vbOKOnly, _
                            "Error"
                    End If
                End If
            End With
        Next
        
        ClearFields
        
        ReloadOriginalFileList
        ReloadRenamedFileList
    End If
    
    Exit Sub
    
RENAME_ERROR:
    MsgBox "There was an error renaming the files.", _
        vbCritical Or vbOKOnly, _
        "Error"
    
    Exit Sub
End Sub

Private Sub OpenArguments(Paths() As String)
    Dim FileNames() As String
    ReDim FileNames(UBound(Paths))
    
    Dim ErrorEncountered As Boolean
    Dim PreviousDirectoryName As String
    Dim DirectoryName As String
    Dim i As Integer
    
    For i = 0 To UBound(Paths)
        Dim Path As String: Path = Paths(i)
        DirectoryName = Left(Path, InStrRev(Path, "\"))
        MsgBox DirectoryName
        
        If PreviousDirectoryName = "" _
        Or DirectoryName = PreviousDirectoryName Then
            PreviousDirectoryName = DirectoryName
            FileNames(i) = Mid(Path, InStrRev(Path, "\") + 1)
        Else
            MsgBox "Files must be located in the same directory.", _
                vbCritical Or vbOKOnly, _
                "Error"
            ErrorEncountered = True
            
            Exit For
        End If
    Next
    
    If Not ErrorEncountered Then
        OpenFiles DirectoryName, FileNames
    Else
        CloseFiles
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title
    mnuAbout.Caption = "&About " & App.Title
        
    If Command <> "" Then
        ' Splits arguments string into files
        Dim Arguments As String: Arguments = Mid(Command, 2, Len(Command) - 2)
        Dim Paths() As String: Paths = Split(Arguments, """ """)
        
        OpenArguments Paths
    End If
    
    ReloadOriginalFileList
    ReloadRenamedFileList
End Sub


Private Sub mnuOpen_Click()
    On Error GoTo CANCEL_ERROR
    
    With cdlg
        .Filter = "All Files (*.*)|*.*"
        .CancelError = True
        .Flags = cdlOFNAllowMultiselect _
            Or cdlOFNExplorer _
            Or cdlOFNLongNames _
            Or cdlOFNHideReadOnly
        .ShowOpen
    End With
    
    Dim Arguments() As String: Arguments = Split(cdlg.FileName, Chr(0), 2)
    Dim DirectoryName As String: DirectoryName = Arguments(0)
    Dim FileNames() As String: FileNames = Split(Arguments(1), Chr(0))
    
    OpenFiles DirectoryName, FileNames
    
    Exit Sub
    
CANCEL_ERROR:
    Exit Sub
End Sub

Private Sub mnuClose_Click()
    CloseFiles
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuWebSite_Click()
    frmAbout.OpenWebSite
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub txtAddBefore_Change()
    EditText
End Sub

Private Sub txtAddAfter_Change()
    EditText
End Sub

Private Sub txtFind_Change()
    EditText
End Sub

Private Sub txtReplace_Change()
    EditText
End Sub

Private Sub txtStart_Change()
    ' Forces to integer
    txtStart.Text = Val(txtStart.Text)
    
    EditText
End Sub

Private Sub chkAfterName_Click()
    EditText
End Sub

Private Sub chkBeforeName_Click()
    EditText
End Sub

Private Sub btnClear_Click()
    ClearFields
End Sub

Private Sub btnRename_Click()
    RenameFiles
End Sub
