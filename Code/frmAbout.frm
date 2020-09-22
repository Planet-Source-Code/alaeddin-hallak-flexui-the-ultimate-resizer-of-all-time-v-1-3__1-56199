VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About FlexUI"
   ClientHeight    =   3075
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2122.42
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   510
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1860
      TabIndex        =   0
      Top             =   2580
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "Send us your thoughts and suggestions at: mhdhallak@hotmail.com"
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   1
      Left            =   1035
      TabIndex        =   6
      Top             =   1785
      Width           =   3225
   End
   Begin VB.Label Label1 
      Caption         =   "Copyright © 2004 by Ahmed Al-Shaikh, Alaeddin Hallak"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   1035
      TabIndex        =   4
      Top             =   1470
      Width           =   4065
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FlexUI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1005
      TabIndex        =   3
      Top             =   195
      Width           =   3885
   End
   Begin VB.Label lblDesc 
      Caption         =   "Add resizing power automatically to your form with very minimal effort."
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1035
      TabIndex        =   2
      Top             =   945
      Width           =   3870
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FlexUI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   480
      Left            =   1005
      TabIndex        =   5
      Top             =   210
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    lblTitle.Caption = "FlexUI " & App.Major & "." & App.Minor
    Label2.Caption = lblTitle.Caption
End Sub

Private Sub lblTitle_Click()

End Sub
