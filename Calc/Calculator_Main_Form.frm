VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Calc 
   BorderStyle     =   0  'None
   Caption         =   "Calculator"
   ClientHeight    =   5310
   ClientLeft      =   6165
   ClientTop       =   4530
   ClientWidth     =   4605
   Icon            =   "Calculator_Main_Form.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Calculator_Main_Form.frx":1042
   ScaleHeight     =   5310
   ScaleWidth      =   4605
   Begin VB.TextBox Dims 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   1935
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   21
      Left            =   2160
      Top             =   1680
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   8421504
      TX              =   "Sqr"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   310
      Left            =   4210
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   80
      Width           =   315
      Begin VB.Timer Syntax 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   240
         Top             =   240
      End
      Begin MSScriptControlCtl.ScriptControl Script 
         Left            =   240
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         AllowUI         =   -1  'True
      End
      Begin VB.Image Image3 
         Height          =   945
         Left            =   0
         Picture         =   "Calculator_Main_Form.frx":52EB8
         Top             =   0
         Width           =   315
      End
   End
   Begin Calculator.Button Solve 
      Height          =   1095
      Left            =   3960
      Top             =   4080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1931
      ForeColor       =   0
      TX              =   "="
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Clr 
      Height          =   495
      Left            =   2160
      Top             =   1080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   0
      TX              =   "CE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   0
      Left            =   2160
      Top             =   4680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ForeColor       =   16711680
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Scripto 
      Height          =   285
      Left            =   -720
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3240
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   -2400
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Code 
      Appearance      =   0  'Flat
      Height          =   3615
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Calculator_Main_Form.frx":53EBA
      Top             =   -3960
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   5
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture2 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   6
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture3 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   7
      Top             =   0
      Width           =   0
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   1
      Left            =   2160
      Top             =   4080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   16711680
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   2
      Left            =   2760
      Top             =   4080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   16711680
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   3
      Left            =   3360
      Top             =   4080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   16711680
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   4
      Left            =   2160
      Top             =   3480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   16711680
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   5
      Left            =   2760
      Top             =   3480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   16711680
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   6
      Left            =   3360
      Top             =   3480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   16711680
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   7
      Left            =   2160
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   16711680
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   8
      Left            =   2760
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   16711680
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   9
      Left            =   3360
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   16711680
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   10
      Left            =   3360
      Top             =   4680
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   0
      TX              =   "."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   1095
      Index           =   11
      Left            =   3960
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1931
      ForeColor       =   33023
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   12
      Left            =   3960
      Top             =   1080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   33023
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   13
      Left            =   3360
      Top             =   1080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   33023
      TX              =   "*"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   14
      Left            =   2760
      Top             =   1080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   33023
      TX              =   "/"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   15
      Left            =   2160
      Top             =   2280
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   255
      TX              =   "("
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   16
      Left            =   2760
      Top             =   2280
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   255
      TX              =   ")"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   17
      Left            =   3360
      Top             =   2280
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   255
      TX              =   "["
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   18
      Left            =   3960
      Top             =   2280
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   255
      TX              =   "]"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   19
      Left            =   2760
      Top             =   1680
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   8421504
      TX              =   "Cos"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   20
      Left            =   3360
      Top             =   1680
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   8421504
      TX              =   "Sin"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Calculator.Button Num 
      Height          =   495
      Index           =   22
      Left            =   3960
      Top             =   1680
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ForeColor       =   8421504
      TX              =   "Pi"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   80
      Top             =   75
      Width           =   195
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   4
      X1              =   0
      X2              =   4560
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   240
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   5445
      Left            =   4680
      Picture         =   "Calculator_Main_Form.frx":53F36
      Top             =   0
      Visible         =   0   'False
      Width           =   4605
   End
End
Attribute VB_Name = "Calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   This Project shows Many Things.
'   One of them, It shows how to make Skinnable
'   Programs. It also Shows how to move your program by
'   ANY object Visible on form (Excluding Shapes And Lines)
'   The main thing, it shows how to make a Copy Paste
'   Calculator.
'
'                 All Copyright Rights Reserved by:
'                                   The only person,
'                                       That only
'                                   You know better than
'                                       anyone, YOU!
'
'  NOT A COPYRIGHTED PROGRAM. CONTAINS CODE FROM OTHER PEOPLE's
'  PROJECTS INCLUDING (CODING GENIOUS)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'
'
'
'    cccccc         a       l           cccccc
'   c      c       a a      l          c      c
'   c             a   a     l          c
'   c            aaaaaaa    l          c
'   c      c    a       a   l      l   c      c
'    cccccc     a       a   llllllll    cccccc
'         __________
'        /_______  /\
'       //______/ / /
'      /__   __  / /
'     //_/__/_/ / /
'    /   /_/   / /
'   /_________/ /
'   \_________\/
'
'
'
'
'
'
'
'
'Declare All Variables
Dim Answ() As String, Nums, oX, oY
Public S As New clsShaped

Private Sub Dims_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Script.Reset
On Error GoTo Err:
Script.AddCode Dims
End If
Exit Sub
Err:
        ErrorMsg.Label1.Caption = Script.Error.Description
        ErrorMsg.Label1.AutoSize = True
        ErrorMsg.Height = ErrorMsg.Label1.Height + (ErrorMsg.Label1.Top * 2)
        ErrorMsg.Width = ErrorMsg.Label1.Width + (ErrorMsg.Label1.Left * 2)
        ErrorMsg.Top = Me.Top + Me.Height / 2
        ErrorMsg.Left = Me.Left + Me.Width / 2 - ErrorMsg.Width / 2
        ErrorMsg.Show
        Syntax.Enabled = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.Top = 0
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
oX = X
oY = Y
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Me.Top = Me.Top + Y - oY
Me.Left = Me.Left + X - oX
End If
Image3.Top = 0
End Sub

Private Sub Image3_Click()
End
End Sub
'Close Button Code
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Top = 0 - (Image3.Height / 3) * 2
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Image3.Top = 0 - (Image3.Height / 3) * 2 Then Image3.Top = 0 - (Image3.Height / 3)
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Top = 0
End Sub

Private Sub Image4_Click()
Me.PopupMenu ErrorMsg.menu
End Sub

''''''''''''''''''''''''''''''''''
Private Sub Num_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text2.SetFocus

    Text2.ZOrder 0
    
    
    Text2.SelText = Nums(Index)

End Sub

Private Sub Syntax_Timer()
    ''''''''''''''''''''''
    'Reset the "Syntax Error" messege
    Text3.Text = ""
    Text2.ZOrder 0
    Text2.SetFocus
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2)
    ErrorMsg.Hide
    Syntax.Enabled = False 'Disable Timer
End Sub

Private Sub Text3_GotFocus()
Text2.SelStart = Len(Text2)
End Sub

Private Sub Clr_Click()
    Syntax.Enabled = False
    'The CE Button Command
    Text2.Text = "" 'Reset Problem Field
    Text3.Text = "" 'Reset Answer Field
    Text2.ZOrder 0 'Set Problem Feild On Top
    Text2.SetFocus 'Set Focus to Problem Field
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'/'/'/''/'/'/'/'/''/'/'/'/'/''/'//////'/''/'/'/''/'/'/'/''/'/''/''/'/''/'/''/'/''/'/'/''/
'Solve Action (See For Your Self)
'/'/'/'/'/'/'/'/'/''/'/'/'/'////'/'//''/'/''/'/'/'/'/''/'/''/'/'/'/'/'/'/'/''/'/'/'/'/''/

Private Sub Solve_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Text2.Text = "" Then Exit Sub
    
    Answ = Split(Text2.Text & " ")
    
    Text2.Text = ""
    
    For i = 0 To UBound(Answ())
        Text2.Text = Text2.Text & Answ(i)
    Next
    Scripto = ""
    
    Scripto = Scripto & "Ans=" & Text2 & Chr(13) & Chr(10) & Code
        
       On Error GoTo Syntax:
     Text3.Text = Script.Eval(Text2.Text)
     
    Text3.ZOrder 0
    Text3.SetFocus
    
    Exit Sub

Syntax:
        ErrorMsg.Label1.Caption = Script.Error.Description
        ErrorMsg.Label1.AutoSize = True
        ErrorMsg.Height = ErrorMsg.Label1.Height + (ErrorMsg.Label1.Top * 2)
        ErrorMsg.Width = ErrorMsg.Label1.Width + (ErrorMsg.Label1.Left * 2)
        ErrorMsg.Top = Me.Top + Me.Height / 2
        ErrorMsg.Left = Me.Left + Me.Width / 2 - ErrorMsg.Width / 2
        ErrorMsg.Show
        Syntax.Enabled = True
End Sub
Private Sub Form_Load()
    Nums = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "+", "-", "*", "/", "(", ")", "[", "]", "Cos(", "Sin(", "SQR(", "*3.14")
    
    For i = 0 To 9
        Num(i).Caption = i
    Next
    
S.Shape Me.hwnd, Image1.Picture, vbMagenta

End Sub
