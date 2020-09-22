VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PassBox Control Test [Version 0.8]"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   282
   ScaleMode       =   2  'Point
   ScaleWidth      =   317.25
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   5160
      Width           =   1215
   End
   Begin Project1.PassBox PassBox4 
      Height          =   315
      Left            =   180
      TabIndex        =   14
      Top             =   3780
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   556
      Text            =   ""
      CueBanner       =   "Type your search here"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   315
      Left            =   5520
      TabIndex        =   13
      Top             =   3780
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   180
      PasswordChar    =   "*"
      TabIndex        =   10
      Text            =   "Password"
      Top             =   3000
      Width           =   2835
   End
   Begin Project1.PassBox PassBox3 
      Height          =   315
      Left            =   3300
      TabIndex        =   11
      Top             =   3000
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   556
      Text            =   "PassBox"
      PasswordChar    =   "*"
   End
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   1140
      Width           =   2835
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Text            =   "TextBox"
      Top             =   420
      Width           =   2835
   End
   Begin Project1.PassBox PassBox1 
      Height          =   315
      Left            =   3300
      TabIndex        =   3
      Top             =   420
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   556
      Text            =   "PassBox"
   End
   Begin Project1.PassBox PassBox2 
      Height          =   1455
      Left            =   3300
      TabIndex        =   7
      Top             =   1140
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2566
      Text            =   $"Form1.frx":0087
      ScrollBars      =   2
      MultiLine       =   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   9
      X2              =   306
      Y1              =   252.75
      Y2              =   252.75
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   9
      X2              =   306
      Y1              =   252
      Y2              =   252
   End
   Begin VB.Label Label7 
      Caption         =   $"Form1.frx":010F
      Height          =   615
      Left            =   180
      TabIndex        =   15
      Top             =   4260
      Width           =   5895
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "This PassBox has the XP avanced features:"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   3480
      Width           =   3120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Password PassBox Control"
      Height          =   195
      Left            =   3300
      TabIndex        =   9
      Top             =   2760
      Width           =   1905
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Password VB TextBox"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MultiLine PassBox Control"
      Height          =   195
      Left            =   3300
      TabIndex        =   5
      Top             =   900
      Width           =   1830
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "MultiLine VB TextBox"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   900
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Standard PassBox Control"
      Height          =   195
      Left            =   3300
      TabIndex        =   1
      Top             =   180
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Standard VB TextBox"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1530
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const UDS_ALIGNLEFT As Long = &H8
Private Const UDS_ALIGNRIGHT As Long = &H4
Private Const UDS_AUTOBUDDY As Long = &H10
Private Const UDS_ARROWKEYS As Long = &H20
Private Const UDS_HORZ As Long = &H40
Private Const UDS_HOTTRACK As Long = &H100
Private Const UDS_NOTHOUSANDS As Long = &H80
Private Const UDS_SETBUDDYINT As Long = &H2
Private Const UDS_WRAP As Long = &H1

Private Const WS_GROUP As Long = &H20000
Private Const WS_BORDER As Long = &H800000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000

Private Declare Function CreateUpDownControl Lib "comctl32.dll" (ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal hParent As Long, ByVal nID As Long, ByVal hInst As Long, ByVal hBuddy As Long, ByVal nUpper As Long, ByVal nLower As Long, ByVal nPos As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private lModule As Long

Private Sub Command3_Click()
  Call PassBox4.ShowBalloonTip("Enter any word in the search box, then press enter", "Search help", TTI_INFO)
End Sub

Private Sub Form_Initialize()
lModule = LoadLibrary("Shell32.dll")
InitCommonControls
End Sub
Private Sub Form_Load()
'Dim lStyle As Long
'  lStyle = WS_VISIBLE Or WS_CHILD Or WS_GROUP Or UDS_SETBUDDYINT Or UDS_ALIGNRIGHT
'  CreateUpDownControl lStyle, 0, 0, 0, 0, PassBox1.hWnd, &H100, App.hInstance, PassBox1.hWndEdit, 255, 0, 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call FreeLibrary(lModule)
End Sub

Private Sub Command1_Click()
  Unload Me
End Sub
