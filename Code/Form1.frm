VERSION 5.00
Object = "*\AFlexUI.vbp"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   10650
   Begin TabDlg.SSTab SSTab1 
      Height          =   2670
      Left            =   180
      TabIndex        =   7
      Top             =   4605
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   4710
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Command2 
         Caption         =   "Resized Horizontally"
         Height          =   510
         Left            =   330
         TabIndex        =   8
         Top             =   915
         Width           =   1845
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Duplicate Me"
      Height          =   705
      Index           =   0
      Left            =   8745
      TabIndex        =   6
      Top             =   1635
      Width           =   1440
   End
   Begin VB.TextBox Text2 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Text            =   "Resized Horizontally + Vertically"
      Top             =   990
      Width           =   8175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   8640
      Picture         =   "Form1.frx":0054
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   960
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Text            =   "Resized Horizontally"
      Top             =   240
      Width           =   8805
   End
   Begin VB.Label Label2 
      Caption         =   "Anchored Right"
      Height          =   495
      Left            =   9180
      TabIndex        =   5
      Top             =   1005
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Anchored Right + Resized Vertically"
      Height          =   495
      Left            =   8610
      TabIndex        =   4
      Top             =   6765
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   8430
      X2              =   8430
      Y1              =   855
      Y2              =   7170
   End
   Begin FlexUIProject.FlexUI FlexUI1 
      Left            =   9885
      Top             =   6225
      _ExtentX        =   820
      _ExtentY        =   794
      ControlsSetupString=   "Text1|35*Picture1|34*Text2|47*Line1|46*Label3|42*Label2|34*Command1(0)|47*SSTab1|59*Command2|35*Label1|32*"
      MinFormWidth    =   718
      MinFormHeight   =   529
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim obj As Object
Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        Load Command1(1)
        With Command1(1)
            .Visible = True
            .Top = Command1(0).Top + Command1(0).Height
            .Left = Command1(0).Left
            .ZOrder vbBringToFront
        End With
        Set obj = Command1(1)
        
        FlexUI1.SetControlStyle Command1(1), ResizeHorizontally
    ElseIf Index = 1 Then
        Unload Command1(1)

    End If
End Sub

Private Sub Command2_Click()
MsgBox Command2.Container.Name

End Sub

Private Sub Text2_Click()
    FlexUI1.CenterFormScreen
End Sub
