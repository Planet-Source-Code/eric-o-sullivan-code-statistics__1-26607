VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Basic 6.0 Code Statistics"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frmStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdScan 
      Caption         =   "&Scan"
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame framProj 
      Caption         =   "Statistics"
      Height          =   2415
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   5055
      Begin VB.Label lblCode 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   24
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblDCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Code Lines :"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblDTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Lines :"
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblMod 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblDMod 
         BackStyle       =   0  'Transparent
         Caption         =   "Modules :"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblForm 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblDForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Forms :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblFunc 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblDFunc 
         BackStyle       =   0  'Transparent
         Caption         =   "Functions :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblProc 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblDProc 
         BackStyle       =   0  'Transparent
         Caption         =   "Procedures :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblComm 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblDComm 
         BackStyle       =   0  'Transparent
         Caption         =   "Comment Lines :"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblDBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Blank Lines :"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "v1.0.0"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblDVer 
         BackStyle       =   0  'Transparent
         Caption         =   "Version :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Project1"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblDName 
         BackStyle       =   0  'Transparent
         Caption         =   "Project Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   120
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog cdgFiles 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const FormStartCode = "Attribute VB_Exposed "
Const ModStartCode = "Attribute VB_Name "
Const VbpTitle = "Title"
Const VbpMajor = "MajorVer"
Const VbpMinor = "MinorVer"
Const VbpRevision = "RevisionVer"
Const VbpForm = "Form"
Const VbpMod = "Module"

Dim NumBlank As Long
Dim NumProc As Long
Dim NumFunc As Long
Dim NumComments As Long
Dim NumForms As Long
Dim NumModules As Long
Dim NumCode As Long
Dim Version As String

Public Sub ReadProject(Path As String)
'This will read an entire project and set the values for statistics

Dim FileNum As Integer 'used for the .vbp file
Dim Line As String
Dim ProjectName As String
Dim FileName As String
Dim StartScan As Boolean

'if path is invalid, then quit
If Dir(Path) = "" Then
    Exit Sub
End If

'reset values
StartScan = False
lblName.Caption = ""
Version = "v"
NumBlank = 0
NumProc = 0
NumFunc = 0
NumComments = 0
NumForms = 0
NumModules = 0
NumCode = 0

'open project
FileNum = FreeFile
Open Path For Input As #FileNum
    While Not EOF(FileNum)
        Line Input #FileNum, Line
        
        Select Case GetBefore(Line)
'        Case FormStartCode
'        Case ModStartCode
        Case VbpTitle
            lblName.Caption = GetAfter(Line)
        
        Case VbpMajor, VbpMinor
            Version = Version & GetAfter(Line) & "."
        
        Case VbpRevision
            Version = Version & GetAfter(Line)
        
        Case VbpForm
            'scan form
            NumForms = NumForms + 1
            Call ScanFile((GetPath(Path) & "\" & GetAfter(Line)), FormStartCode)
            
        Case VbpMod
            'scan module
            NumModules = NumModules + 1
            Call ScanFile((GetPath(Path) & "\" & GetMod(Line)), ModStartCode)
        
        End Select
        
    Wend
Close #FileNum

'display results
If Trim(lblName.Caption) = "" Then
    'if the project name is blank then use the default name
    lblName.Caption = "Project1"
End If
lblVersion.Caption = Version
lblBlank.Caption = Format(NumBlank, "0")
lblComm.Caption = Format(NumComments, "0")
lblForm.Caption = Format(NumForms, "0")
lblMod.Caption = Format(NumModules, "0")
lblProc.Caption = Format(NumProc, "0")
lblFunc.Caption = Format(NumFunc, "0")
lblCode.Caption = Format(NumCode, "0")

'total results accounting for headers/footers of procedures/functions
lblTotal.Caption = Format((NumBlank + NumComments + (NumProc + NumFunc * 2) + NumCode), "0")
End Sub

Public Sub IncrementVal(Line As String)
'This will increment the appropiate values based on the text

Const ProcA = "Public Sub"
Const ProcB = "Private Sub"
Const EndProc = "End Sub"
Const FuncA = "Public Function"
Const FuncB = "Private Function"
Const EndFunc = "End Function"
Const Comment = "'"
Const Blank = ""

'Functions
If Left(Line, Len(ProcA)) = ProcA Then
    NumProc = NumProc + 1
    Exit Sub
End If
If Left(Line, Len(ProcB)) = ProcB Then
    NumProc = NumProc + 1
    Exit Sub
End If

'Functions
If Left(Line, Len(FuncA)) = FuncA Then
    NumFunc = NumFunc + 1
    Exit Sub
End If
If Left(Line, Len(FuncB)) = FuncB Then
    NumFunc = NumFunc + 1
    Exit Sub
End If

'Comments
If Left(Line, 1) = Comment Then
    NumComments = NumComments + 1
    Exit Sub
End If

'Blanks
If Line = Blank Then
    NumBlank = NumBlank + 1
    Exit Sub
End If

'the footers of the functions and procedures
If Left(Line, Len(EndProc)) = EndProc Then
    Exit Sub
End If
If Left(Line, Len(EndFunc)) = EndFunc Then
    Exit Sub
End If

'else the line is code
NumCode = NumCode + 1
End Sub

Public Function GetPath(Address As String) As String
'This function returns the path from a string containing the full
'path and filename of a file.

Dim Counter As Integer
Dim LastPos As Integer

'find the position of the last "\" mark in the string
LastPos = 1
For Counter = 1 To Len(Address)
    If Mid(Address, Counter, 1) = "\" Then
        LastPos = Counter
    End If
Next Counter

'return everything before the last "\" mark
GetPath = Left(Address, (LastPos - 1))
End Function

Public Function GetBefore(Sentence As String) As String
'This procedure returns all the character of a
'string before the "=" sign.

Const Sign = "="

Dim Counter As Integer
Dim Before As String

'find the position of the equals sign
Counter = InStr(1, Sentence, Sign)

If (Counter <> Len(Sentence)) And (Counter <> 0) Then
    Before = Left(Sentence, (Counter - 1))
Else
    Before = ""
End If

GetBefore = Before
End Function

Public Function GetAfter(Sentence As String) As String
'This procedure returns all the character of a
'string after the "=" sign.

Const Sign = "="

Dim Counter As Integer
Dim Rest As String

'find the position of the equals sign
Counter = InStr(1, Sentence, Sign)

If Counter <> Len(Sentence) Then
    Rest = Right(Sentence, (Len(Sentence) - Counter))
Else
    Rest = ""
End If

GetAfter = Rest
End Function

Public Function GetMod(Sentence As String) As String
'This procedure returns all the character of a
'string after the "=" sign.

Const ModName = ";"

Dim Rest As String
Dim ModPos As Integer

'find the position of the ; sign
ModPos = InStr(1, Sentence, ModName) + 1

If ModPos <> Len(Sentence) Then
    Rest = Right(Sentence, (Len(Sentence) - ModPos))
Else
    Rest = ""
End If

GetMod = Rest
End Function

Private Sub cmdBrowse_Click()
cdgFiles.Filter = "VB Project *.Vbp|*.Vbp|All Files *.*|*.*"
cdgFiles.InitDir = GetPath(txtPath.Text)
cdgFiles.ShowOpen
txtPath.Text = cdgFiles.FileName
End Sub

Private Sub cmdScan_Click()
Call ReadProject(txtPath.Text)
End Sub

Private Sub ScanFile(Path As String, Start As String)
'This procedure will scan a file starting at the first point with the
'specified starting string.

Dim FileNum As Integer
Dim Line As String
Dim StartScan As Boolean

FileNum = FreeFile

If Dir(Path) = "" Then
    'invalid path
    Exit Sub
End If

Open Path For Input As #FileNum
    'scan file
    While Not EOF(FileNum)
        Line Input #FileNum, Line
        If StartScan Then
            Call IncrementVal(LTrim(Line))
        End If
        
        If Left(Line, Len(Start)) = Start Then
            'scan code
            StartScan = True
        End If
    Wend
Close #FileNum
End Sub

Private Sub Form_Load()
txtPath.Text = App.Path
txtPath.SelLength = Len(txtPath.Text)
End Sub
