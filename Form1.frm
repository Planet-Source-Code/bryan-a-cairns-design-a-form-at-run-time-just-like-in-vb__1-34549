VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Design Form at Run Time"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000C&
      Height          =   3255
      Left            =   3960
      ScaleHeight     =   3195
      ScaleWidth      =   2475
      TabIndex        =   15
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "Command9"
         Height          =   375
         Index           =   8
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Command8"
         Height          =   375
         Index           =   7
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Command7"
         Height          =   375
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "Command6"
         Height          =   375
         Index           =   5
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Command5"
         Height          =   375
         Index           =   4
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Command4"
         Height          =   375
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command3"
         Height          =   375
         Index           =   2
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Command2"
         Height          =   375
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Command1"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   6255
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Height          =   540
         Index           =   5
         Left            =   5640
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   14
         Top             =   840
         Width           =   540
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Height          =   540
         Index           =   0
         Left            =   4440
         Picture         =   "Form1.frx":030A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         Top             =   240
         Width           =   540
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Height          =   540
         Index           =   1
         Left            =   4440
         Picture         =   "Form1.frx":0614
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   12
         Top             =   840
         Width           =   540
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Height          =   540
         Index           =   2
         Left            =   5040
         Picture         =   "Form1.frx":0A56
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   11
         Top             =   840
         Width           =   540
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Height          =   540
         Index           =   3
         Left            =   5040
         Picture         =   "Form1.frx":0D60
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         Top             =   240
         Width           =   540
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Height          =   540
         Index           =   4
         Left            =   5640
         Picture         =   "Form1.frx":106A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   240
         Width           =   540
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000016&
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000016&
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000016&
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   4680
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   3480
      Width           =   1335
   End
   Begin Project1.UserControl1 Resizer 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   375
      _ExtentX        =   2990
      _ExtentY        =   2355
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PLEASE VOTE!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   28
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":1374
      Height          =   975
      Index           =   2
      Left            =   120
      TabIndex        =   27
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":143B
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Design a Form at Run Time just like in VB!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PLEASE VOTE as this usercontrol is the result of a lot of work
'Bryan Cairns - cairnsb@html-helper.com
'
'This form is pretty boring, sorry but I wanted to make this usercontrol
'simple to use in your projects.
'All you need to do to "activate" the "grippers" is to set the "BoundControl" property
'the usercontrol will take case of the rest

Private Sub Form_Load()
'*** IMPORTANT ***
'Set the ZOrder of the control to 0
Resizer.ZOrder 0
Set Resizer.BoundControl = Text1(Index)
End Sub

Private Sub Check1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Resizer.BoundControl = Check1(Index)
End Sub

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Resizer.BoundControl = Command1(Index)
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Resizer.BoundControl = Frame1
End Sub

Private Sub Label1_Click(Index As Integer)
 Set Resizer.BoundControl = Label1(Index)
End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Resizer.BoundControl = Picture1(Index)
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Resizer.BoundControl = Picture2
End Sub

Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Resizer.BoundControl = Text1(Index)
End Sub

