VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ikonz v1.0"
   ClientHeight    =   8685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12330
   Icon            =   "extract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "extract.frx":0782
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   822
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   8160
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8160
      Width           =   1410
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9720
      TabIndex        =   14
      Text            =   "0"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10920
      TabIndex        =   13
      Top             =   3600
      Width           =   855
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFC0&
      Height          =   1650
      Hidden          =   -1  'True
      Left            =   9960
      Pattern         =   "*.dll;*.ocx;*.exe"
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   7920
      TabIndex        =   11
      Top             =   1440
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFC0&
      Height          =   1665
      Left            =   7920
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFC0&
      Height          =   1815
      Left            =   1440
      TabIndex        =   8
      Top             =   6120
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Extract Selected Icon"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      Height          =   3255
      Left            =   840
      ScaleHeight     =   3195
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   840
      Width           =   5175
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   33375
         Left            =   0
         ScaleHeight     =   33315
         ScaleWidth      =   4995
         TabIndex        =   3
         Top             =   0
         Width           =   5055
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   0
            Left            =   -540
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   4
            Top             =   10
            Width           =   480
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "www.gauravcreations.cjb.net"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   360
            TabIndex        =   7
            Top             =   1320
            Width           =   4455
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3015
      LargeChange     =   10
      Left            =   6120
      Max             =   100
      Min             =   1
      SmallChange     =   5
      TabIndex        =   1
      Top             =   1080
      Value           =   1
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "icons"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   16
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Show files with min. icons"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Icons Displayed"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   4320
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ikonz v1.1'
' Source Code By Gaurav dhup '
' www.gauravcreations.cjb.net '

Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Path As String
Dim x As Integer ' Variable for index of picture box 2
Dim setload As Integer ' Variable for determining whether icons are loaded
Dim oldtop As Integer ' variable for scrolling icons to the next line
Dim portind As Integer ' variable for determining the index of the selected icons
Dim CurRgn, TempRgn As Long
'For Dragging Borderless Forms...
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2


' Save Command '
Private Sub Command1_Click()
On Error GoTo ClassHandler1
ClassHandler1:
    If Err.Number = 53 Then
    End If
    Resume Next
CommonDialog1.Filter = "Bmp File|*.bmp|Icon File|*.ico"
CommonDialog1.Action = 2
SavePicture Picture2(portind).Image, CommonDialog1.FileName
End Sub

' Searching for specified number of icons '
Private Sub Command2_Click()
' Erasing the previously displayed list '
If List1.ListCount <> 0 Then
listc = List1.ListCount - 1
For cfile = listc To 0 Step -1
    List1.RemoveItem cfile
Next cfile
End If
ProgressBar1.Visible = True
If File1.ListCount <> 0 Then
    ProgressBar1.Max = File1.ListCount
End If

' Checking for specified value and transfering the appropriate file paths to list box from file list box
For cfile = 0 To File1.ListCount - 1
    File1.ListIndex = cfile
    pathcheck = File1.Path + "\" + File1.FileName
    return1& = ExtractIcon(Me.hWnd, pathcheck, -1)
    If return1& >= Val(Text3.Text) Then
       List1.AddItem pathcheck
    End If
    ProgressBar1.Value = cfile
Next cfile
ProgressBar1.Visible = False
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo ClassHandler1
ClassHandler1:
    If Err.Number = 68 Then
        response1 = MsgBox("Drive not ready", vbOKOnly, "Ikons")
    End If
    Resume Next
Dir1.Path = Drive1.Drive
End Sub

' Extract Icons from File list box on Dbl Click '
Private Sub File1_dblClick()
Path = File1.Path + "\" + File1.FileName
Call extract
End Sub

Private Sub Form_Load()
'Form1.Picture = LoadPicture(App.Path & "\frame.gif")
AutoFormShape Form1, RGB(254, 254, 254)
setload = 0
oldtop = Picture2(x).Top
Unload Form2
End Sub

Private Sub Label4_Click()
Shell "Explorer http://www.gauravcreations.cjb.net"
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.ForeColor = &HFF0000
End Sub

Private Sub List1_Click()
Path = List1.List(List1.ListIndex)
Call extract
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label4.ForeColor = &HFF&
End Sub

' Selecting Icons displayed for extraction '
Private Sub Picture2_Click(Index As Integer)
For x = 1 To Val(Text1.Text)
    Picture2(x).BorderStyle = 0
Next x
Picture2(Index).BorderStyle = 1
portind = Index
End Sub

Private Sub VScroll1_Change()
   Picture1.Top = -VScroll1.Value
End Sub

' Main function to extract the icons from Dll, Exes and Ocx files
Private Sub extract()
 return1& = ExtractIcon(Me.hWnd, Path, -1)
If setload = 1 Then
   For x = 1 To Val(Text1.Text)
       Unload Picture2(x)
   Next x
   setload = 0
End If
Text1.Text = return1&
If Val(Text1.Text) > 18 Then
   Label4.Visible = False
End If
If Val(Text1.Text) < 18 Then
   Label4.Visible = True
End If

For x = 1 To Val(Text1.Text)
    Load Picture2(x)
    Picture2(x).Left = Picture2(x - 1).Left + 560
    Picture2(x).Top = oldtop
    If Picture2(x).Left > (9 * 560) Then
       Picture2(x).Left = 10
       Picture2(x).Top = Picture2(x).Top + 560
       oldtop = Picture2(x).Top
    End If
    Picture2(x).Visible = True
    Picture2(x).Picture = LoadPicture()
    return2& = ExtractIcon(Me.hWnd, Path, return1& - x)
    return3& = DrawIcon(Picture2(x).hdc, 0, 0, return2&)
Next x
setload = 1
oldtop = 10
VScroll1.Max = (Val(Text1.Text) \ 11) * 560
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error Resume Next
'Move the borderless form...
Call ReleaseCapture
Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub


