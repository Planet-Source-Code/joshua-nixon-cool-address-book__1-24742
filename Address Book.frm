VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Address Book"
   ClientHeight    =   9345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13770
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   13770
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame19 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Address Created (Click on textbox to open)"
      Height          =   2055
      Left            =   4440
      TabIndex        =   68
      Top             =   5880
      Width           =   3495
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   240
         Width           =   2775
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2325
      Left            =   -2640
      Picture         =   "Address Book.frx":0000
      ScaleHeight     =   2325
      ScaleWidth      =   2865
      TabIndex        =   45
      Top             =   5880
      Width           =   2865
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2280
      Top             =   1680
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      Picture         =   "Address Book.frx":3BFB
      ScaleHeight     =   330
      ScaleWidth      =   13755
      TabIndex        =   35
      Top             =   0
      Width           =   13755
      Begin VB.CommandButton Command19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Height          =   255
         Left            =   13460
         Picture         =   "Address Book.frx":128BD
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "End"
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton Command22 
         Height          =   255
         Left            =   13080
         Picture         =   "Address Book.frx":12C74
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   20
         Width           =   255
      End
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   360
      Picture         =   "Address Book.frx":1304C
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Exit"
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   360
      Picture         =   "Address Book.frx":135E6
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Load Address"
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   360
      Picture         =   "Address Book.frx":13BA8
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Save Address"
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   360
      Picture         =   "Address Book.frx":142BA
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "New Address"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4275
      Left            =   120
      Picture         =   "Address Book.frx":148A2
      ScaleHeight     =   4275
      ScaleWidth      =   2820
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   2820
      Begin VB.Shape Shape7 
         Height          =   855
         Left            =   120
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Shape Shape6 
         Height          =   855
         Left            =   120
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Shape Shape5 
         Height          =   855
         Left            =   120
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000000&
         Height          =   855
         Left            =   120
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "      Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   39
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "      Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "      Load"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "      New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Height          =   255
      Left            =   60
      Picture         =   "Address Book.frx":15DB5
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   360
      Width           =   735
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   9090
      Width           =   13770
      _ExtentX        =   24289
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   1
            Object.Width           =   21669
            TextSave        =   "11:36 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/4/2001"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Personal"
      Height          =   2415
      Left            =   4440
      TabIndex        =   16
      Top             =   3360
      Width           =   3495
      Begin VB.TextBox Text29 
         Height          =   285
         Left            =   1320
         TabIndex        =   85
         Text            =   "No"
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   1320
         TabIndex        =   84
         Text            =   "Yes"
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   49
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Insert Picture"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1320
         TabIndex        =   42
         Text            =   "Female"
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1320
         TabIndex        =   41
         Text            =   "Male"
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Relative"
         Height          =   855
         Left            =   140
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "No"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Yes"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Female"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Male"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gender"
         Height          =   975
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Shape Shape12 
         BorderWidth     =   2
         Height          =   735
         Left            =   2520
         Top             =   240
         Width           =   735
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   2640
         Picture         =   "Address Book.frx":16248
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Information"
      Height          =   7815
      Left            =   8160
      TabIndex        =   7
      Top             =   720
      Width           =   5295
      Begin VB.Frame Frame21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Telephone"
         Height          =   2535
         Left            =   120
         TabIndex        =   52
         Top             =   3240
         Width           =   4935
         Begin VB.Frame Frame25 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Company Phone"
            Height          =   615
            Left            =   2400
            TabIndex        =   62
            Top             =   1560
            Width           =   2175
            Begin VB.TextBox Text23 
               Height          =   285
               Left            =   720
               TabIndex        =   64
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox Text22 
               Height          =   285
               Left            =   120
               TabIndex        =   63
               Text            =   "()"
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame24 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cell Phone"
            Height          =   615
            Left            =   120
            TabIndex        =   59
            Top             =   1560
            Width           =   2175
            Begin VB.TextBox Text21 
               Height          =   285
               Left            =   120
               TabIndex        =   61
               Text            =   "()"
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox Text20 
               Height          =   285
               Left            =   720
               TabIndex        =   60
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame23 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Home Phone 2"
            Height          =   615
            Left            =   120
            TabIndex        =   56
            Top             =   840
            Width           =   2175
            Begin VB.TextBox Text19 
               Height          =   285
               Left            =   120
               TabIndex        =   58
               Text            =   "()"
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox Text18 
               Height          =   285
               Left            =   720
               TabIndex        =   57
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame22 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Home Phone"
            Height          =   615
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   2175
            Begin VB.TextBox Text17 
               Height          =   285
               Left            =   720
               TabIndex        =   55
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox Text16 
               Height          =   285
               Left            =   120
               TabIndex        =   54
               Text            =   "()"
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Shape Shape10 
            BorderWidth     =   2
            Height          =   735
            Left            =   4080
            Top             =   240
            Width           =   735
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   4200
            Picture         =   "Address Book.frx":16F12
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Location"
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   6000
         Width           =   5055
         Begin VB.Frame Frame16 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Zip"
            Height          =   615
            Left            =   3000
            TabIndex        =   50
            Top             =   240
            Width           =   975
            Begin VB.TextBox Text10 
               Height          =   285
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame18 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Country"
            Height          =   615
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   2055
            Begin VB.TextBox Text13 
               Height          =   285
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H00E0E0E0&
            Caption         =   "State"
            Height          =   615
            Left            =   1560
            TabIndex        =   24
            Top             =   240
            Width           =   1335
            Begin VB.TextBox Text9 
               Height          =   285
               Left            =   120
               TabIndex        =   25
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "City"
            Height          =   615
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1335
            Begin VB.TextBox Text8 
               Height          =   285
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Shape Shape11 
            BorderWidth     =   2
            Height          =   735
            Left            =   4080
            Top             =   240
            Width           =   735
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   4200
            Picture         =   "Address Book.frx":177DC
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Company Address 2"
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   2895
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Company Address"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   2895
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Home Address 2"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2895
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Home Address"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2895
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Line Line4 
         X1              =   5160
         X2              =   5160
         Y1              =   3120
         Y2              =   5880
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   5160
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line2 
         X1              =   5160
         X2              =   5160
         Y1              =   120
         Y2              =   3120
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5160
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Shape Shape9 
         BorderWidth     =   2
         Height          =   735
         Left            =   4320
         Top             =   240
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4440
         Picture         =   "Address Book.frx":1881E
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Last"
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Middle"
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "First"
      Height          =   615
      Left            =   4560
      TabIndex        =   0
      Top             =   960
      Width           =   1815
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Name"
      Height          =   2415
      Left            =   4440
      TabIndex        =   6
      Top             =   720
      Width           =   3375
      Begin VB.Image Image6 
         Height          =   480
         Left            =   2640
         Picture         =   "Address Book.frx":194E8
         Top             =   360
         Width           =   480
      End
      Begin VB.Shape Shape13 
         BorderWidth     =   2
         Height          =   735
         Left            =   2520
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H00000000&
      Caption         =   "Picture "
      ForeColor       =   &H0000FF00&
      Height          =   2655
      Left            =   120
      TabIndex        =   40
      Top             =   840
      Width           =   2295
      Begin VB.Shape Shape16 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   2295
         Left            =   240
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   2175
      Left            =   2760
      ScaleHeight     =   2115
      ScaleWidth      =   915
      TabIndex        =   48
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "No"
      Height          =   255
      Left            =   7560
      TabIndex        =   83
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "No"
      Height          =   255
      Left            =   7200
      TabIndex        =   82
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "No"
      Height          =   255
      Left            =   6840
      TabIndex        =   81
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "No"
      Height          =   255
      Left            =   6480
      TabIndex        =   80
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "No"
      Height          =   255
      Left            =   6120
      TabIndex        =   79
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "www.yarinteractive.com"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6120
      MousePointer    =   1  'Arrow
      TabIndex        =   65
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Image Image7 
      Height          =   780
      Left            =   4320
      Picture         =   "Address Book.frx":1A52A
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Shape Shape14 
      BorderWidth     =   2
      Height          =   320
      Left            =   0
      Top             =   345
      Width           =   860
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "BY:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   120
      TabIndex        =   44
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      Height          =   2655
      Left            =   4320
      Top             =   600
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      Height          =   8055
      Left            =   8040
      Top             =   600
      Width           =   5535
   End
   Begin VB.Shape Shape3 
      Height          =   2655
      Left            =   4320
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   8655
      Left            =   4200
      Top             =   360
      Width           =   9495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim name1 As String
Dim name2 As String
Dim name3 As String
Dim name4 As String
Dim name5 As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
 

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


Command1.Top = Command1.Top + 100
Command1.Left = Command1.Left + 100

End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape16.Visible = False
CommonDialog1.Filter = "All Image Files|*.jpg;*.bmp;*.gif;*.ico;*.cur"
On Error Resume Next
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
CommonDialog1.Filter = ""
'picture1
If Err <> 32755 Then
Picture4.Visible = False
OpenFileName = CommonDialog1.FileName
MousePointer = 0
Picture4.ScaleMode = 3
Picture4.Picture = LoadPicture()
Picture4.Picture = LoadPicture(OpenFileName)
Picture4.Visible = True
End If
Let Image1.Picture = Picture4.Picture
Command1.Top = Command1.Top - 100
Command1.Left = Command1.Left - 100
Text14.Text = CommonDialog1.FileName
Picture4.Visible = False
End Sub

Private Sub Command11_Click()
Shape16.Visible = True

Me.Text1 = ""
Me.Text2 = ""
Me.Text3 = ""
Me.Text4 = ""
Me.Text5 = ""
Me.Text6 = ""
Me.Text7 = ""
Me.Text8 = ""
Me.Text9 = ""
Me.Text10 = ""
Me.Text11 = ""
Me.Text12 = ""
Me.Text13 = ""
Me.Text16 = ""
Me.Text17 = ""
Me.Text18 = ""
Me.Text19 = ""
Me.Text20 = ""
Me.Text21 = ""
Me.Text22 = ""
Me.Text23 = ""

End Sub

Private Sub Command18_Click()
Picture2.Visible = True
Command6.Visible = True
Command8.Visible = True
Command9.Visible = True
Command11.Visible = True

End Sub

Private Sub Command19_Click()
End
End Sub

Private Sub Command2_Click()
Label7.Caption = "Yes"
name1 = InputBox("Label your member")
Text15.ToolTipText = name1
Form2.Show


End Sub
Sub save()
CommonDialog1.Filter = "Address Files|*.adr"
On Error Resume Next
CommonDialog1.FileName = ""
CommonDialog1.ShowSave
On Error Resume Next
Open CommonDialog1.FileName For Output As #1
Print #1, Text1.Text
Print #1, Text2.Text
Print #1, Text3.Text
Print #1, Text4.Text
Print #1, Text5.Text
Print #1, Text6.Text
Print #1, Text7.Text
Print #1, Text8.Text
Print #1, Text9.Text
Print #1, Text10.Text
Print #1, Text11.Text
Print #1, Text12.Text
Print #1, Text13.Text
Print #1, Text14.Text
Print #1, Text15.Text
Print #1, Text16.Text
Print #1, Text17.Text
Print #1, Text18.Text
Print #1, Text19.Text
Print #1, Text20.Text
Print #1, Text21.Text
Print #1, Text22.Text
Print #1, Text23.Text
Print #1, Text28.Text
Print #1, Text29.Text

Close #1

End Sub

Private Sub Command22_Click()
Form1.WindowState = 1
End Sub

Private Sub Command3_Click()
Label8.Caption = "Yes"
name2 = InputBox("Label your member")
Text24.ToolTipText = name2

Form2.Show
End Sub

Private Sub Command4_Click()
Label9.Caption = "Yes"
name3 = InputBox("Label your member")
Text25.ToolTipText = name3

Form2.Show
End Sub

Private Sub Command5_Click()
Label10.Caption = "Yes"
name4 = InputBox("Label your member")
Text26.ToolTipText = name4

Form2.Show
End Sub

Private Sub Command6_Click()
Call save
End Sub

Private Sub Command7_Click()
Label11.Caption = "Yes"
name5 = InputBox("Label your member")
Text27.ToolTipText = name5

Form2.Show
End Sub

Private Sub Command8_Click()
End
End Sub

Private Sub Command9_Click()
Call load
End Sub
Sub load()
Shape16.Visible = False
CommonDialog1.Filter = "Address Files|*.adr"
On Error Resume Next
    CommondialogDialog1.FileName = ""
    CommonDialog1.ShowOpen
On Error Resume Next
Open CommonDialog1.FileName For Input As #1
 Line Input #1, filens$
 Text1.Text = filens$
 Line Input #1, filens$
 Text2.Text = filens$
 Line Input #1, filens$
 Text3.Text = filens$
 Line Input #1, filens$
 Text4.Text = filens$
 Line Input #1, filens$
 Text5.Text = filens$
 Line Input #1, filens$
 Text6.Text = filens$
 Line Input #1, filens$
 Text7.Text = filens$
 Line Input #1, filens$
 Text8.Text = filens$
 Line Input #1, filens$
 Text9.Text = filens$
 Line Input #1, filens$
 Text10.Text = filens$
 Line Input #1, filens$
 Text11.Text = filens$
 Line Input #1, filens$
 Text12.Text = filens$
 Line Input #1, filens$
 Text13.Text = filens$
 Line Input #1, filens$
 Text14.Text = filens$
 Line Input #1, filens$
 Text16.Text = filens$
 Line Input #1, filens$
 Text17.Text = filens$
 Line Input #1, filens$
 Text18.Text = filens$
 Line Input #1, filens$
 Text19.Text = filens$
 Line Input #1, filens$
 Text20.Text = filens$
 Line Input #1, filens$
 Text21.Text = filens$
 Line Input #1, filens$
 Text22.Text = filens$
 Line Input #1, filens$
 Text23.Text = filens$
 Line Input #1, filens$
 Text28.Text = filens$
 Line Input #1, filens$
 Text29.Text = filens$
 Line Input #1, filens$
  Close #1
If Err <> 32755 Then
Picture4.Visible = False
CommonDialog1.FileName = Text14.Text
OpenFileName = Text14.Text
MousePointer = 0
Picture4.ScaleMode = 3
Picture4.Picture = LoadPicture()
Picture4.Picture = LoadPicture(OpenFileName)
Picture4.Visible = True
End If
Let Image1.Picture = Picture4.Picture
Command1.Top = Command1.Top - 100
Command1.Left = Command1.Left - 100
Picture4.Visible = False
End Sub
Private Sub Form_Click()
Picture2.Visible = False
Command6.Visible = False
Command8.Visible = False
Command9.Visible = False
Command11.Visible = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFF0000
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape4.BorderColor = &HFFC0C0
Shape7.BorderColor = &H0&
Shape5.BorderColor = &H0&
Shape6.BorderColor = &H0&
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.BorderColor = &HFFC0C0
Shape4.BorderColor = &H0&
Shape6.BorderColor = &H0&
Shape7.BorderColor = &H0&
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape6.BorderColor = &HFFC0C0
Shape4.BorderColor = &H0&
Shape5.BorderColor = &H0&
Shape7.BorderColor = &H0&
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape7.BorderColor = &HFFC0C0
Shape4.BorderColor = &H0&
Shape5.BorderColor = &H0&
Shape6.BorderColor = &H0&
End Sub

Private Sub Label6_Click()
ShellExecute hwnd, "open", "http://www.yarinteractive.com", vbNullString, vbNullString, conSwNormal

End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &H80FF&
End Sub

Private Sub List1_Click()
Dim name As String
If List1.Selected(name) = True Then
name = Text1.Text & Text2.Text & Text3.Text
End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Option2.Value = False
Text11.Text = "Male"
Text12.Text = ""
ElseIf Option2.Value = True Then
Option1.Value = False
Text11.Text = ""
Text12.Text = "Female"
End If
End Sub

Private Sub Option2_Click()
If Option1.Value = True Then
Option2.Value = False
Text11.Text = "Male"
Text12.Text = ""
ElseIf Option2.Value = True Then
Option1.Value = False
Text11.Text = ""
Text12.Text = "Female"
End If

End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Text28.Text = "Yes"
Text29.Text = ""
ElseIf Option4.Value = True Then
Text29.Text = "No"
Text28.Text = ""
End If
End Sub

Private Sub Option4_Click()
If Option3.Value = True Then
Text28.Text = "Yes"
Text29.Text = ""
ElseIf Option4.Value = True Then
Text29.Text = "No"
Text28.Text = ""
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape4.BorderColor = &H0&
Shape5.BorderColor = &H0&
Shape6.BorderColor = &H0&
Shape7.BorderColor = &H0&
End Sub

Private Sub Text15_Click()
Shape16.Visible = False
On Error Resume Next
CommonDialog1.FileName = Text15.Text
On Error Resume Next
Open Text15.Text For Input As #1
 Line Input #1, filens$
 Text1.Text = filens$
 Line Input #1, filens$
 Text2.Text = filens$
 Line Input #1, filens$
 Text3.Text = filens$
 Line Input #1, filens$
 Text4.Text = filens$
 Line Input #1, filens$
 Text5.Text = filens$
 Line Input #1, filens$
 Text6.Text = filens$
 Line Input #1, filens$
 Text7.Text = filens$
 Line Input #1, filens$
 Text8.Text = filens$
 Line Input #1, filens$
 Text9.Text = filens$
 Line Input #1, filens$
 Text10.Text = filens$
 Line Input #1, filens$
 Text11.Text = filens$
 Line Input #1, filens$
 Text12.Text = filens$
 Line Input #1, filens$
 Text13.Text = filens$
 Line Input #1, filens$
 Text14.Text = filens$
 Line Input #1, filens$
 Text16.Text = filens$
 Line Input #1, filens$
 Text17.Text = filens$
 Line Input #1, filens$
 Text18.Text = filens$
 Line Input #1, filens$
 Text19.Text = filens$
 Line Input #1, filens$
 Text20.Text = filens$
 Line Input #1, filens$
 Text21.Text = filens$
 Line Input #1, filens$
 Text22.Text = filens$
 Line Input #1, filens$
 Text23.Text = filens$
 Line Input #1, filens$
   Text28.Text = filens$
 Line Input #1, filens$
 Text29.Text = filens$
 Line Input #1, filens$
 
   Close #1
 If Err <> 32755 Then
Picture4.Visible = False
CommonDialog1.FileName = Text14.Text
OpenFileName = Text14.Text
MousePointer = 0
Picture4.ScaleMode = 3
Picture4.Picture = LoadPicture()
Picture4.Picture = LoadPicture(OpenFileName)
Picture4.Visible = True
End If
Let Image1.Picture = Picture4.Picture
Picture4.Visible = False
End Sub

Private Sub Text24_Click()
Shape16.Visible = False
On Error Resume Next
CommonDialog1.FileName = Text24.Text
On Error Resume Next
Open Text24.Text For Input As #1
 Line Input #1, filens$
 Text1.Text = filens$
 Line Input #1, filens$
 Text2.Text = filens$
 Line Input #1, filens$
 Text3.Text = filens$
 Line Input #1, filens$
 Text4.Text = filens$
 Line Input #1, filens$
 Text5.Text = filens$
 Line Input #1, filens$
 Text6.Text = filens$
 Line Input #1, filens$
 Text7.Text = filens$
 Line Input #1, filens$
 Text8.Text = filens$
 Line Input #1, filens$
 Text9.Text = filens$
 Line Input #1, filens$
 Text10.Text = filens$
 Line Input #1, filens$
 Text11.Text = filens$
 Line Input #1, filens$
 Text12.Text = filens$
 Line Input #1, filens$
 Text13.Text = filens$
 Line Input #1, filens$
 Text14.Text = filens$
 Line Input #1, filens$
 Text16.Text = filens$
 Line Input #1, filens$
 Text17.Text = filens$
 Line Input #1, filens$
 Text18.Text = filens$
 Line Input #1, filens$
 Text19.Text = filens$
 Line Input #1, filens$
 Text20.Text = filens$
 Line Input #1, filens$
 Text21.Text = filens$
 Line Input #1, filens$
 Text22.Text = filens$
 Line Input #1, filens$
 Text23.Text = filens$
 Line Input #1, filens$
   Text28.Text = filens$
 Line Input #1, filens$
 Text29.Text = filens$
 Line Input #1, filens$
 
   Close #1
 If Err <> 32755 Then
Picture4.Visible = False
CommonDialog1.FileName = Text14.Text
OpenFileName = Text14.Text
MousePointer = 0
Picture4.ScaleMode = 3
Picture4.Picture = LoadPicture()
Picture4.Picture = LoadPicture(OpenFileName)
Picture4.Visible = True
End If
Let Image1.Picture = Picture4.Picture
Picture4.Visible = False
End Sub

Private Sub Text25_Click()
Shape16.Visible = False
On Error Resume Next
CommonDialog1.FileName = Text25.Text
On Error Resume Next
Open Text25.Text For Input As #1
 Line Input #1, filens$
 Text1.Text = filens$
 Line Input #1, filens$
 Text2.Text = filens$
 Line Input #1, filens$
 Text3.Text = filens$
 Line Input #1, filens$
 Text4.Text = filens$
 Line Input #1, filens$
 Text5.Text = filens$
 Line Input #1, filens$
 Text6.Text = filens$
 Line Input #1, filens$
 Text7.Text = filens$
 Line Input #1, filens$
 Text8.Text = filens$
 Line Input #1, filens$
 Text9.Text = filens$
 Line Input #1, filens$
 Text10.Text = filens$
 Line Input #1, filens$
 Text11.Text = filens$
 Line Input #1, filens$
 Text12.Text = filens$
 Line Input #1, filens$
 Text13.Text = filens$
 Line Input #1, filens$
 Text14.Text = filens$
 Line Input #1, filens$
 Text16.Text = filens$
 Line Input #1, filens$
 Text17.Text = filens$
 Line Input #1, filens$
 Text18.Text = filens$
 Line Input #1, filens$
 Text19.Text = filens$
 Line Input #1, filens$
 Text20.Text = filens$
 Line Input #1, filens$
 Text21.Text = filens$
 Line Input #1, filens$
 Text22.Text = filens$
 Line Input #1, filens$
 Text23.Text = filens$
 Line Input #1, filens$
   Text28.Text = filens$
 Line Input #1, filens$
 Text29.Text = filens$
 Line Input #1, filens$
 
   Close #1
 If Err <> 32755 Then
Picture4.Visible = False
CommonDialog1.FileName = Text14.Text
OpenFileName = Text14.Text
MousePointer = 0
Picture4.ScaleMode = 3
Picture4.Picture = LoadPicture()
Picture4.Picture = LoadPicture(OpenFileName)
Picture4.Visible = True
End If
Let Image1.Picture = Picture4.Picture
Picture4.Visible = False
End Sub

Private Sub Text26_Change()
Shape16.Visible = False
On Error Resume Next
CommonDialog1.FileName = Text26.Text
On Error Resume Next
Open Text26.Text For Input As #1
 Line Input #1, filens$
 Text1.Text = filens$
 Line Input #1, filens$
 Text2.Text = filens$
 Line Input #1, filens$
 Text3.Text = filens$
 Line Input #1, filens$
 Text4.Text = filens$
 Line Input #1, filens$
 Text5.Text = filens$
 Line Input #1, filens$
 Text6.Text = filens$
 Line Input #1, filens$
 Text7.Text = filens$
 Line Input #1, filens$
 Text8.Text = filens$
 Line Input #1, filens$
 Text9.Text = filens$
 Line Input #1, filens$
 Text10.Text = filens$
 Line Input #1, filens$
 Text11.Text = filens$
 Line Input #1, filens$
 Text12.Text = filens$
 Line Input #1, filens$
 Text13.Text = filens$
 Line Input #1, filens$
 Text14.Text = filens$
 Line Input #1, filens$
 Text16.Text = filens$
 Line Input #1, filens$
 Text17.Text = filens$
 Line Input #1, filens$
 Text18.Text = filens$
 Line Input #1, filens$
 Text19.Text = filens$
 Line Input #1, filens$
 Text20.Text = filens$
 Line Input #1, filens$
 Text21.Text = filens$
 Line Input #1, filens$
 Text22.Text = filens$
 Line Input #1, filens$
 Text23.Text = filens$
 Line Input #1, filens$
   Text28.Text = filens$
 Line Input #1, filens$
 Text29.Text = filens$
 Line Input #1, filens$
 
   Close #1
 If Err <> 32755 Then
Picture4.Visible = False
CommonDialog1.FileName = Text14.Text
OpenFileName = Text14.Text
MousePointer = 0
Picture4.ScaleMode = 3
Picture4.Picture = LoadPicture()
Picture4.Picture = LoadPicture(OpenFileName)
Picture4.Visible = True
End If
Let Image1.Picture = Picture4.Picture
Picture4.Visible = False
End Sub

Private Sub Text27_Change()
Shape16.Visible = False
On Error Resume Next
CommonDialog1.FileName = Text26.Text
On Error Resume Next
Open Text26.Text For Input As #1
 Line Input #1, filens$
 Text1.Text = filens$
 Line Input #1, filens$
 Text2.Text = filens$
 Line Input #1, filens$
 Text3.Text = filens$
 Line Input #1, filens$
 Text4.Text = filens$
 Line Input #1, filens$
 Text5.Text = filens$
 Line Input #1, filens$
 Text6.Text = filens$
 Line Input #1, filens$
 Text7.Text = filens$
 Line Input #1, filens$
 Text8.Text = filens$
 Line Input #1, filens$
 Text9.Text = filens$
 Line Input #1, filens$
 Text10.Text = filens$
 Line Input #1, filens$
 Text11.Text = filens$
 Line Input #1, filens$
 Text12.Text = filens$
 Line Input #1, filens$
 Text13.Text = filens$
 Line Input #1, filens$
 Text14.Text = filens$
 Line Input #1, filens$
 Text16.Text = filens$
 Line Input #1, filens$
 Text17.Text = filens$
 Line Input #1, filens$
 Text18.Text = filens$
 Line Input #1, filens$
 Text19.Text = filens$
 Line Input #1, filens$
 Text20.Text = filens$
 Line Input #1, filens$
 Text21.Text = filens$
 Line Input #1, filens$
 Text22.Text = filens$
 Line Input #1, filens$
 Text23.Text = filens$
 Line Input #1, filens$
   Text28.Text = filens$
 Line Input #1, filens$
 Text29.Text = filens$
 Line Input #1, filens$
 
   Close #1
 If Err <> 32755 Then
Picture4.Visible = False
CommonDialog1.FileName = Text14.Text
OpenFileName = Text14.Text
MousePointer = 0
Picture4.ScaleMode = 3
Picture4.Picture = LoadPicture()
Picture4.Picture = LoadPicture(OpenFileName)
Picture4.Visible = True
End If
Let Image1.Picture = Picture4.Picture
Picture4.Visible = False
End Sub

Private Sub Timer1_Timer()
Picture3.Left = Picture3.Left + 50
If Picture3.Left >= 720 Then
Picture3.Left = Picture3.Left - 50
If Picture3.Left <= -270 Then
Picture3.Left = Picture3.Left + 50
If Picture3.Left >= 0 Then
Timer1.Enabled = False
End If
End If
End If
End Sub
