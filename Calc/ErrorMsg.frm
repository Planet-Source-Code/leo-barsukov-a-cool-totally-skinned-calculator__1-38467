VERSION 5.00
Begin VB.Form ErrorMsg 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1275
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Error"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   10080
      Left            =   0
      Picture         =   "ErrorMsg.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22530
   End
   Begin VB.Menu menu 
      Caption         =   "system"
      Visible         =   0   'False
      Begin VB.Menu Quit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "ErrorMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_LostFocus()
Me.Hide
Calc.Syntax.Enabled = False
End Sub

Private Sub Quit_Click()
End
End Sub
