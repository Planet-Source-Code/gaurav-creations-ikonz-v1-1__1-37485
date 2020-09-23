VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   LinkTopic       =   "Form2"
   Picture         =   "splash.frx":0000
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   332
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   600
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurRgn, TempRgn As Long
Dim counter As Integer
Private Sub Form_Load()
    'Form2.Picture = LoadPicture(App.Path & "\ikonz.gif")
    AutoFormShape Form2, RGB(254, 254, 254)
End Sub

Private Sub Timer1_Timer()
Form1.Show
End Sub

