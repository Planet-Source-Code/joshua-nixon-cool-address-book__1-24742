VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   LinkTopic       =   "Form2"
   ScaleHeight     =   3795
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      Picture         =   "Address Book2.frx":0000
      ScaleHeight     =   330
      ScaleWidth      =   13755
      TabIndex        =   4
      Top             =   0
      Width           =   13755
      Begin VB.CommandButton Command19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Height          =   255
         Left            =   4120
         Picture         =   "Address Book2.frx":1908
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "End"
         Top             =   40
         Width           =   255
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   3255
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2040
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2340
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
 

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Private Sub Command1_Click()
If Form1.Label7.Caption = "Yes" Then
Let Form1.Text15.Text = Text1.Text
Form1.Label8.Caption = "No"

ElseIf Form1.Label8.Caption = "Yes" Then
Let Form1.Text24.Text = Text1.Text
Form1.Label8.Caption = "No"

ElseIf Form1.Label9.Caption = "Yes" Then
Let Form1.Text25.Text = Text1.Text
Form1.Label9.Caption = "No"

ElseIf Form1.Label10.Caption = "Yes" Then
Let Form1.Text26.Text = Text1.Text
Form1.Label10.Caption = "No"

ElseIf Form1.Label11.Caption = "Yes" Then
Let Form1.Text27.Text = Text1.Text
Form1.Label11.Caption = "No"
End If
End Sub

Private Sub Command19_Click()
Form2.Hide
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
Dim filen As String
If Right(Dir1.Path, 1) = "\" Then
filen = Dir1.Path + File1.FileName
Else
filen = Dir1.Path + "\" + File1.FileName
End If
Text1.Text = filen
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
