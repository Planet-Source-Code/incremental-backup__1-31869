VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Incremental Backup"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Backup Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Return to Selection Screen"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         ToolTipText     =   "Original files are newer than backup files."
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         ToolTipText     =   "Original files are not newer than backup files."
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Files Copied"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Files Skipped"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Files to Copy"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Copy Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   2760
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   8415
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         ScaleHeight     =   225
         ScaleWidth      =   5385
         TabIndex        =   23
         Top             =   960
         Width           =   5415
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Text            =   "Text8"
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         ScaleHeight     =   225
         ScaleWidth      =   5385
         TabIndex        =   16
         Top             =   1320
         Width           =   5415
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "kBytes Remaining"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "kBytes Copied"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Copying"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VB.ListBox FindFilesTmpDirs 
         Height          =   450
         Left            =   1920
         TabIndex        =   18
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ListBox FindFilesTmpResults 
         Height          =   645
         Left            =   1920
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ListBox FindFilesResults 
         Height          =   4350
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   10695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Src As String, Dest As String, Junk As String, Fol As String
    Dim Trash As String
    Dim OrigFiles As Single, CopiedFiles As Single, SkippedFiles As Single
    Dim FCurrent, FDest, Fs  'scripting
    Dim DateCurrent As Date, DateDest As Date
    Dim X As Integer
    Dim CurFile As String
    Dim DestPath As String
    Dim CurDir As String
    Dim BarString As String
    Dim KCopied As Single, i As Single, KLeft As Single, KSkipped As Single
    Dim TotalK As Double, DoneK As Double
    
Private Sub cmdReturn_Click()
    SetDrives.Show
    Unload Me
End Sub

Private Sub Form_Load()
    'collect Parameters
    Src = SetDrives.Label3(0)
    Dest = SetDrives.Label3(1)
    Unload SetDrives
    
    'defaults
    OrigFiles = 0
    CopiedFiles = 0
    SkippedFiles = 0
    KCopied = 0
    KSkipped = 0
    DoneK = 0
    Picture1.ForeColor = RGB(0, 0, 255)
    Picture2.ForeColor = RGB(0, 0, 255)
    cmdReturn.Enabled = False

    'display
    Frame1.Caption = "Building File List: 0 Files"
    Me.Visible = True
    Me.Refresh
    Screen.MousePointer = vbHourglass

    FindFilesTmpResults.Clear
    FindFilesTmpDirs.Clear
    FindFilesResults.Clear
    
    'find files in source
    SearchFilesInDir Src, "*.*"
    
    'parse source for all directories
    Frame1.Caption = "Analysing " & FindFilesResults.ListCount & " Files for Backup"
    For X = 4 To Len(Src)
        If Mid(Src, X, 1) = "\" Then
            Junk = "[dir]  " & Left(Src, X)
            FindFilesResults.AddItem Junk

        End If
    Next X

    FindFilesResults.Refresh
    Text3 = FindFilesResults.ListCount
    Text3.Refresh

    'build destination folder structure
    For X = 0 To FindFilesResults.ListCount - 1
        Junk = FindFilesResults.List(X)
        If Left(Junk, 7) = "[dir]  " Then
            Junk = Right(Junk, Len(Junk) - 10)
            Junk = Dest & Junk
            'test if present in destination
            Trash = Dir(Junk, vbDirectory)

            If Trash = "" Then
                MkDir Junk
                CopiedFiles = CopiedFiles + 1
                Text5 = CopiedFiles
                Text5.Refresh
            End If
        Else    'exit loop (list sorted)
            GoTo AfterX
        End If
    Next X
AfterX:

    'files
    TotalK = Val(Text8)
    Copy_Files
    
    Text7 = 0
    
    Screen.MousePointer = vbDefault
    cmdReturn.Enabled = True
    DoEvents

End Sub

Private Sub Copy_Files()
    On Error GoTo EHandler
    Frame2.Visible = True
    Frame3.Visible = True
    Me.Refresh
    For X = 0 To FindFilesResults.ListCount - 1
        BarString = "Files "
        If X + 1 <= FindFilesResults.ListCount Then i = ((X + 1) / FindFilesResults.ListCount) * 100
        UpdateProgress Picture1, i
        FindFilesResults.Selected(X) = True
        Junk = FindFilesResults.List(X)
        CurFile = Right(Junk, Len(Junk) - 7)
        DestPath = Dest & Right(CurFile, Len(CurFile) - 3)

        'If (GetAttr(CurFile) And vbDirectory) <> vbDirectory Then
        If Left(Junk, 7) = "[file] " Then
            'copy if curfile is newer than destpath
            Set Fs = CreateObject("scripting.filesystemobject")
            'current file
            DateCurrent = #1/1/2001#
            Set FCurrent = Fs.getfile(CurFile)
            DateCurrent = FCurrent.datelastmodified
            Junk = Format(DateCurrent, "dd/mm/yyyy")
            DateCurrent = CDate(Junk)
            'destination file
            DateDest = #1/1/1991#
            Junk = Dir(DestPath)
            If Junk = "" Then Junk = Dir(DestPath, vbDirectory)
            If Junk <> "" Then
                Set FDest = Fs.getfile(DestPath)
                DateDest = FDest.datelastmodified
                Junk = Format(DateDest, "dd/mm/yyyy")
                DateDest = CDate(Junk)
            End If

            If DateCurrent > DateDest Then
                Call FileCopy(CurFile, DestPath)
                CopiedFiles = CopiedFiles + 1
                Text5 = CopiedFiles
                Text1 = CurFile
                Text2 = DestPath
                Text1.Refresh
                Text2.Refresh
                KCopied = KCopied + (FileLen(CurFile) / 1024)
                Text6 = Format(KCopied, "##########,##")
                KLeft = Val(Text8) - FileLen(CurFile)
                Text8 = KLeft
                Text7 = Format(KLeft / 1024, "##########,##")
                DoneK = DoneK + FileLen(CurFile)
                BarString = "kBytes "
                If DoneK <= TotalK Then i = (DoneK / TotalK) * 100
                UpdateProgress Picture2, i
                DoEvents
            Else
                SkippedFiles = SkippedFiles + 1
                Text4 = SkippedFiles
                Text4.Refresh
                FindFilesResults.List(X) = FindFilesResults.List(X) & "      **SKIPPED **"
                KLeft = Val(Text8) - FileLen(CurFile)
                Text8 = KLeft
                Text7 = Format(KLeft / 1024, "##########,##")
                DoneK = DoneK + FileLen(CurFile)
                BarString = "kBytes "
                If DoneK <= TotalK Then i = (DoneK / TotalK) * 100
                UpdateProgress Picture2, i
            End If
        Else
            SkippedFiles = SkippedFiles + 1
            Text4 = SkippedFiles
            Text4.Refresh
        End If
NxtX:
    Next X
    Frame2.ForeColor = &HFF&
    Frame2.Caption = "Backup Completed"
    Text5 = Val(Text3) - Val(Text4)
    Exit Sub
EHandler:
    Select Case Err.Number
        Case 70
            Resume Next
    End Select
End Sub

Sub UpdateProgress(PB As Control, ByVal percent)
    Dim num$        'use percent
    If Not PB.AutoRedraw Then      'picture in memory ?
        PB.AutoRedraw = -1          'no, make one
    End If
    PB.Cls                      'clear picture in memory
    PB.ScaleWidth = 100         'new scalemodus
    PB.DrawMode = 10            'not XOR Pen Modus
    num$ = BarString & Format$(percent, "###") + "%"
    PB.CurrentX = 50 - PB.TextWidth(num$) / 2
    PB.CurrentY = (PB.ScaleHeight - PB.TextHeight(num$)) / 2
    PB.Print num$               'print percent
    PB.Line (0, 0)-(percent, PB.ScaleHeight), , BF
    PB.Refresh          'show difference
End Sub
