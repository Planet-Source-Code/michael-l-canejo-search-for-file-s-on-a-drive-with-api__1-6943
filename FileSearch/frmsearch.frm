VERSION 5.00
Begin VB.Form frmsearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Finder By: MiKE3D"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   240
      MaxLength       =   100
      TabIndex        =   12
      Text            =   "Example"
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   240
      MaxLength       =   100
      TabIndex        =   11
      Text            =   "Example"
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox zFileExt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   9
      Text            =   ".exe"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox zFileExt2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "frmsearch.frx":0000
      Left            =   3600
      List            =   "frmsearch.frx":0002
      TabIndex        =   7
      Top             =   2400
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmdend 
      Cancel          =   -1  'True
      Caption         =   "End"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CheckBox chkhidelist 
      Caption         =   "Hide lists when searching"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.ListBox lstdirs 
      Height          =   1620
      Left            =   3600
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.ListBox lstfiles 
      Height          =   2010
      ItemData        =   "frmsearch.frx":0004
      Left            =   3600
      List            =   "frmsearch.frx":0006
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ßy: MiKE3D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MiKE_3D@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Label lbltime 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time taken: 0 second(s)"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   3015
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CountDir As Integer
Dim CountFiles As Integer

Private Sub cmdend_Click()
End
End Sub

Private Sub cmdgo_Click()
Dim starttime As Single
Label1.Caption = ""
Command1.Enabled = True
Drive1.Enabled = False
lstfiles.Clear
Text1.Enabled = False
zFileExt.Enabled = False
starttime = Timer
lstdirs.AddItem Drive1.Drive & "\"
List1.AddItem Drive1.Drive & "\"
Do
lbltime = "Searching . . . " & lstdirs.List(0)
findfilesapi lstdirs.List(0), LCase$("*.*")
lstdirs.RemoveItem 0
Loop Until lstdirs.ListCount = 0
lstdirs.Visible = True
lstfiles.Visible = True
CountFiles = lstfiles.ListCount
lbltime = "Time taken: " & Timer - starttime & " second(s)"
Command1.Caption = "Search !"
Label1.Caption = "Directories: " & CountDir & "   Files:  " & CountFiles
CountDir = 0
CountFiles = 0
Drive1.Enabled = True
Text1.Enabled = True
zFileExt.Enabled = True
chkhidelist.Enabled = True
List1.Visible = True
lstfiles.Visible = True
cmdend.Left = 7440
Me.Width = 9180
End Sub

Sub findfilesapi(DirPath As String, FileSpec As String)
On Error Resume Next
Dim FindData As WIN32_FIND_DATA
Dim FindHandle As Long
Dim FindNextHandle As Long
Dim filestring As String
Dim GetExta As String, GetExtb As String, GetExtc As String
DirPath = Trim$(DirPath)
If Right(DirPath, 1) <> "\" Then
DirPath = DirPath & "\"
End If
FindHandle = FindFirstFile(DirPath & FileSpec, FindData)
If FindHandle <> 0 Then
If FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
If Left$(FindData.cFileName, 1) <> "." And Left$(FindData.cFileName, 2) <> ".." Then
filestring = DirPath & Trim$(FindData.cFileName) & "\"
lstdirs.AddItem LCase(filestring), 1
List1.AddItem LCase(filestring), 1
End If
Else
filestring = DirPath & Trim$(FindData.cFileName)
GetExta$ = Mid(filestring, InStr(filestring, "."))
zFileExt2 = LCase(GetExta$)
If zFileExt = zFileExt2 Then lstfiles.AddItem LCase(filestring)
End If
End If
If FindHandle <> 0 Then
Do
DoEvents
FindNextHandle = FindNextFile(FindHandle, FindData)
If FindNextHandle <> 0 Then
If FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
' It's a directory
If Left$(FindData.cFileName, 1) <> "." And Left$(FindData.cFileName, 2) <> ".." Then
filestring = DirPath & Trim$(FindData.cFileName) & "\"
lstdirs.AddItem LCase(filestring), 1
List1.AddItem LCase(filestring), 1
CountDir = CountDir + 1
End If
Else
filestring = DirPath & Trim$(FindData.cFileName)
GetExta$ = Mid(filestring, InStr(filestring, "."))
GetExtb$ = Mid(FindData.cFileName, InStr(FindData.cFileName, "."))
GetExtb$ = Replace(FindData.cFileName, GetExtb$, "")
Text2 = LCase(GetExtb$)
zFileExt2 = LCase(GetExta$)
If zFileExt = "*.*" Then GoTo here
If Text1 = "" Then
If zFileExt = zFileExt2 Then
lstfiles.AddItem LCase(filestring)
GoTo here2
End If
End If
If zFileExt = zFileExt2 And Text2 = LCase(Text1) Then
here:
lstfiles.AddItem LCase(filestring)
here2:
End If
End If
Else
Exit Do
End If
Loop
End If
Call FindClose(FindHandle)
End Sub
'Instead of using findfilesapi you could use this method. Only thing its slower
Public Sub findfilesdir(DirPath As String, FileSpec As String)
On Error Resume Next
Dim filestring As String
DirPath = Trim$(DirPath)
If Right$(DirPath, 1) <> "\" Then
DirPath = DirPath & "\"
End If
filestring = Dir$(DirPath & FileSpec, vbArchive Or vbHidden Or vbSystem Or vbDirectory)
Do
DoEvents
If filestring = "" Then
Exit Do
Else
If (GetAttr(DirPath & filestring) And vbDirectory) = vbDirectory Then
If Left$(filestring, 1) <> "." And Left$(filestring, 2) <> ".." Then
lstdirs.AddItem LCase(DirPath) & LCase(filestring) & "\", 1
List1.AddItem LCase(DirPath) & LCase(filestring) & "\", 1
CountDir = CountDir + 1
End If
Else
lstfiles.AddItem LCase(DirPath) & LCase(filestring)
End If
End If
filestring = Dir$
Loop
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Search !" Then
chkhidelist.Enabled = False
lstfiles.Clear
lstdirs.Clear
List1.Clear
Command1.Caption = "Continue"
GoTo here
End If
here:
If Command1.Caption = "Continue" Then
Command1.Enabled = False
Command1.Caption = "Stop!"
If chkhidelist.Value = 1 Then
List1.Visible = False
lstfiles.Visible = False
cmdend.Left = 1920
Me.Width = 3585
End If
If chkhidelist.Value = 0 Then
List1.Visible = True
lstfiles.Visible = True
cmdend.Left = 7440
Me.Width = 9180
End If
cmdgo_Click
Exit Sub
End If
If Command1.Caption = "Stop!" Then
lstdirs.Visible = True
lstfiles.Visible = True
CountFiles = lstfiles.ListCount
lbltime = "Time taken: 0 second(s)"
Command1.Caption = "Search !"
Label1.Caption = "Directories: " & CountDir & "   Files:  " & CountFiles
CountDir = 0
CountFiles = 0
Drive1.Enabled = True
Text1.Enabled = True
zFileExt.Enabled = True
chkhidelist.Enabled = True
lstfiles.Clear
lstdirs.Clear
List1.Visible = True
lstfiles.Visible = True
cmdend.Left = 7440
Me.Width = 9180
Do
DoEvents
Loop
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If List1.ListCount < 1 Then Exit Sub
Dim a As Integer, b As Integer
For a = 0 To List1.ListCount - 1
For b = a + 1 To List1.ListCount - 1
If List1.List(a) = List1.List(b) Then
List1.RemoveItem b
b = b - 1
End If
Next b
Next a
End Sub

Private Sub Form_Load()
cmdend.Left = 1920
Me.Width = 3585
End Sub

Private Sub lbltime_Change()
If Len(lbltime) > 200 Then lbltime.Caption = lstdirs.List(0) & "..."
End Sub

Private Sub zFileExt_Change()
zFileExt = LCase(zFileExt)
End Sub
