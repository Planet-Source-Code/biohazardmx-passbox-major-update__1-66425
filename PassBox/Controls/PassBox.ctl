VERSION 5.00
Begin VB.UserControl PassBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   36
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   164
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PassBox Control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   1365
   End
End
Attribute VB_Name = "PassBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************
' Control:   PassBox
' Date:      08/09/2006
' Author:    BioHazardMX
' Purpose:   PassBox UserControl module
' Version:   0.8
' Requires:  OleGuids.tlb (IDE only)
'***************************************************************************************
' ChangeLog:
' [Version 0.8]
'  * New CueBanner and BalloonTip support under WinXP!
'  * Fixed bug with MouseMove event and Button mask
'  * Fixed bug with BackColor and ForeColor not applying at startup
'  * Minimum Height is now restricted to 19 pixels (like VB TextBoxes)
' [Version 0.7]
'  * Code re-arranged and commented
'  * Added standard events
'  * Fixed bug with scrollbars showing in single line mode
'  * Added Back-Fore color properties
'  * Fixed bug with Shift mask in WM_KEYDOWN-UP
' [Version 0.6]
'  * This was the first public version (uploaded to planetsourcecode.com)
'  * Fixed bug with Enter key and IPAO
'  * Added "Can..." properties
'  * Fixed bug in Get/Set SelLength
' [Version 0.5]
'  * New IPAO for focus (fixed Tab key issues)
'  * New "Locked" property
' [Version 0.4]
'  * Fixed bug with "Text" property, no more crashes (using GetWindowTextLength)
' [Version 0.3]
'  * Fixed repeated AttachMessage causing an "Message Already Handled" message box
'  * Added scrollbars for multiline mode
'  * Fixed multiline & password styles, now can't be mixed
' [Version 0.2]
'  * Added Single line AutoHScroll mode
'  * Added Multiline mode
'  * Fixed subclassing error with WM_KEYUP and Tab key
' [Version 0.1]
'  * First version, basic Password edit control
'***************************************************************************************
 Option Explicit
'***************************************************************************************
' Constants
'***************************************************************************************
'---Window Messages
Private Const WM_MOUSEACTIVATE As Long = &H21
Private Const WM_CLEAR As Long = &H303
Private Const WM_CHAR As Long = &H102
Private Const WM_USER As Long = &H400
Private Const WM_SETFONT As Long = &H30
Private Const WM_SETTEXT As Long = &HC
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_COMMAND As Long = &H111
'---Window Styles
Private Const WS_CHILD As Long = &H40000000
Private Const WS_BORDER As Long = &H800000
Private Const WS_TABSTOP As Long = &H10000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILDWINDOW As Long = (WS_CHILD)
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_EX_NOPARENTNOTIFY As Long = &H4&
'---Edit Messages
Private Const EM_CANPASTE As Long = (WM_USER + 50)
Private Const EM_CANREDO As Long = (WM_USER + 85)
Private Const EM_CANUNDO As Long = &HC6
Private Const EM_GETLIMITTEXT As Long = (WM_USER + 37)
Private Const EM_GETSEL As Long = &HB0
Private Const EM_LIMITTEXT As Long = &HC5
Private Const EM_REPLACESEL As Long = &HC2
Private Const EM_SETLIMITTEXT As Long = EM_LIMITTEXT
Private Const EM_SETSEL As Long = &HB1
Private Const EM_UNDO As Long = &HC7
Private Const EM_GETPASSWORDCHAR As Long = &HD2
Private Const EM_SETPASSWORDCHAR As Long = &HCC
Private Const ECM_FIRST As Long = &H1500
Private Const EM_SETCUEBANNER As Long = (ECM_FIRST + 1)
Private Const EM_GETCUEBANNER As Long = (ECM_FIRST + 2)    '// Set the cue banner with the lParm = LPCWSTR
Private Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)     '// Show a balloon tip associated to the edit control
Private Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)
'---Edit Styles
Private Const ES_CENTER As Long = &H1&
Private Const ES_LEFT As Long = &H0&
Private Const ES_LOWERCASE As Long = &H10&
Private Const ES_MULTILINE As Long = &H4&
Private Const ES_NUMBER As Long = &H2000&
Private Const ES_READONLY As Long = &H800&
Private Const ES_RIGHT As Long = &H2&
Private Const ES_UPPERCASE As Long = &H8&
Private Const ES_PASSWORD As Long = &H20&
Private Const ES_AUTOHSCROLL As Long = &H80&
Private Const ES_AUTOVSCROLL As Long = &H40&
Private Const ES_WANTRETURN As Long = &H1000&
'---Edit Notification Messages
Private Const EN_CHANGE As Long = &H300
Private Const EN_ERRSPACE As Long = &H500
Private Const EN_HSCROLL As Long = &H601
Private Const EN_KILLFOCUS As Long = &H200
Private Const EN_SELCHANGE As Long = &H702
Private Const EN_SETFOCUS As Long = &H100
Private Const EN_VSCROLL As Long = &H602
'---Misc API Constants
Private Const MK_ALT As Long = &H20
Private Const MK_CONTROL As Long = &H8
Private Const MK_LBUTTON As Long = &H1
Private Const MK_MBUTTON As Long = &H10
Private Const MK_RBUTTON As Long = &H2
Private Const MK_SHIFT As Long = &H4
Private Const VK_TAB As Long = &H9
Private Const MA_NOACTIVATE As Long = 3
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE As Long = -16
'---LOGFONT Constants
Private Const LOGPIXELSX As Long = &H58
Private Const LOGPIXELSY As Long = &H5A
Private Const LF_FACESIZE As Long = &H20
Private Const FW_NORMAL As Long = &H190
Private Const FW_BOLD As Long = &H2BC
Private Const FF_DONTCARE As Long = &H0
Private Const DEFAULT_PITCH As Long = &H0
Private Const DEFAULT_CHARSET As Long = &H1
Private Const DEFAULT_QUALITY As Long = &H0
Private Const DRAFT_QUALITY As Long = &H1
Private Const PROOF_QUALITY As Long = &H2
Private Const NONANTIALIASED_QUALITY As Long = &H3
Private Const ANTIALIASED_QUALITY As Long = &H4
'---Autocomplete Flags
Private Const SHACF_DEFAULT As Long = &H0
Private Const SHACF_FILESYSTEM As Long = &H1
Private Const SHACF_URLHISTORY As Long = &H2
Private Const SHACF_URLMRU As Long = &H4
Private Const SHACF_USETAB As Long = &H8
Private Const SHACF_URLALL As Long = (SHACF_URLHISTORY Or SHACF_URLMRU)
Private Const SHACF_FILESYS_ONLY As Long = &H10
Private Const SHACF_FILESYS_DIRS As Long = &H20
Private Const SHACF_AUTOSUGGEST_FORCE_ON As Long = &H10000000
Private Const SHACF_AUTOSUGGEST_FORCE_OFF As Long = &H20000000
Private Const SHACF_AUTOAPPEND_FORCE_ON As Long = &H40000000
Private Const SHACF_AUTOAPPEND_FORCE_OFF As Long = &H80000000
Private Const S_OK = 0
'---Scrollbar Styles
Private Const SB_BOTH As Long = 3
Private Const SB_HORZ As Long = 0
Private Const SB_VERT As Long = 1
'***************************************************************************************
' User Defined Types
'***************************************************************************************
'---LOGFONT
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type
'---EDITBALLOONTIP
Private Type EDITBALLOONTIP
  cbStruct As Long
  pszTitle As Long
  pszText As Long
  ttiIcon As Long
End Type
'---OSVERSIONINFO
Private Type OSVERSIONINFO
  dwVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion(0 To 127) As Byte
End Type
'***************************************************************************************
' Enumerations
'***************************************************************************************
Public Enum BalloonTipIconConstants
  TTI_NONE = 0
  TTI_INFO = 1
  TTI_WARNING = 2
  TTI_ERROR = 3
End Enum
'***************************************************************************************
' API Declares
'***************************************************************************************
'---General API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreateFontIndirect Lib "GDI32.dll" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function CreateWindowEx Lib "User32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function DestroyWindow Lib "User32.dll" (ByVal hWnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Integer
Private Declare Function GetDC Lib "User32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetFocus Lib "User32.dll" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long
Private Declare Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "User32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "User32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function LockWindowUpdate Lib "User32.dll" (ByVal hWndLock As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "User32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageLong Lib "User32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageString Lib "User32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageLongW Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetBkColor Lib "GDI32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "GDI32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetFocus Lib "User32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "User32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowText Lib "User32.dll" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SHAutoComplete Lib "SHLWAPI.dll" (ByVal hWndEdit As Long, ByVal dwFlags As Long) As Long
Private Declare Function ShowScrollBar Lib "User32.dll" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
'***************************************************************************************
' Variables, Classes and Implements
'***************************************************************************************
'---Misc Variables
Private lTxtWnd As Long
Private lhWnd As Long
Private lhDC As Long
Private lWidth As Long
Private lHeight As Long
Private lPtr As Long
Private lStyle As Long
Private lFont As LOGFONT
Private hFont As Long
Private bRunning As Boolean
Private tIPAOHookStruct As IPAOHookStruct
'---Property Variables
Private lBackColor As Long
Private lForeColor As Long
Private lSBars As Long
Private sText As String
Private sPassChar As String
Private bEnabled As Boolean
Private bMultiLine As Boolean
Private bLocked As Boolean
Private bPassword As Boolean
Private sCueBanner As String
Private sTipTitle As String
Private sTipText As String
'---Implements
Implements ISubclass
'***************************************************************************************
' Events
'***************************************************************************************
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event Change()
'***************************************************************************************
' Subclassing
'***************************************************************************************
'---MsgResponse Let
Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
  '...'
End Property
'---MsgResponse Get
Private Property Get ISubclass_MsgResponse() As EMsgResponse
  ISubclass_MsgResponse = emrPreprocess
End Property
'---WindowProc
Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lButton As MouseButtonConstants
Dim lShift As ShiftConstants
Dim iKeyCode As Integer, lNotify As Long
  Select Case iMsg
   '------------------------------------------------------------------------------
   'Implement focus.  Code taken from vbAccelerator.com
    Case WM_SETFOCUS
      If (lTxtWnd = hWnd) Then
        'The control itself
         Dim pOleObject                  As IOleObject
         Dim pOleInPlaceSite             As IOleInPlaceSite
         Dim pOleInPlaceFrame            As IOleInPlaceFrame
         Dim pOleInPlaceUIWindow         As IOleInPlaceUIWindow
         Dim pOleInPlaceActiveObject     As IOleInPlaceActiveObject
         Dim PosRect                     As RECT
         Dim ClipRect                    As RECT
         Dim FrameInfo                   As OLEINPLACEFRAMEINFO
         Dim grfModifiers                As Long
         Dim AcceleratorMsg              As Msg
        'Get in-place frame and make sure it is set to our in-between
        'implementation of IOleInPlaceActiveObject in order to catch
        'TranslateAccelerator calls
         Set pOleObject = Me
         Set pOleInPlaceSite = pOleObject.GetClientSite
         pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
         CopyMemory pOleInPlaceActiveObject, tIPAOHookStruct.ThisPointer, 4
         pOleInPlaceFrame.SetActiveObject pOleInPlaceActiveObject, vbNullString
         If Not pOleInPlaceUIWindow Is Nothing Then
           pOleInPlaceUIWindow.SetActiveObject pOleInPlaceActiveObject, vbNullString
         End If
         CopyMemory pOleInPlaceActiveObject, 0&, 4
      Else
        'The user control:
         SetFocusAPI lhWnd
      End If
    Case WM_MOUSEACTIVATE
      If GetFocus() <> lhWnd And GetFocus() <> lTxtWnd Then
         SetFocusAPI UserControl.hWnd
         ISubclass_WindowProc = MA_NOACTIVATE
         Exit Function
      Else
         ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
      End If
   'End Implement focus.
   '------------------------------------------------------------------------------
   Case WM_LBUTTONDOWN, WM_RBUTTONDOWN
     If iMsg = WM_LBUTTONDOWN Then lButton = vbLeftButton
     If iMsg = WM_RBUTTONDOWN Then lButton = vbRightButton
     If wParam <> 0 Then
       If wParam = (wParam And MK_ALT) Then lShift = vbAltMask
       If wParam = (wParam And MK_SHIFT) Then lShift = vbShiftMask
       If wParam = (wParam And MK_CONTROL) Then lShift = vbCtrlMask
     End If
     RaiseEvent MouseDown(CInt(lButton), CInt(lShift), LoWord(lParam), HiWord(lParam))
     
   Case WM_LBUTTONUP, WM_RBUTTONUP
     If iMsg = WM_LBUTTONUP Then lButton = vbLeftButton
     If iMsg = WM_RBUTTONUP Then lButton = vbRightButton
     If wParam <> 0 Then
       If wParam = (wParam And MK_ALT) Then lShift = vbAltMask
       If wParam = (wParam And MK_SHIFT) Then lShift = vbShiftMask
       If wParam = (wParam And MK_CONTROL) Then lShift = vbCtrlMask
     End If
     RaiseEvent MouseUp(CInt(lButton), CInt(lShift), LoWord(lParam), HiWord(lParam))
     RaiseEvent Click
       
   Case WM_MOUSEMOVE
     If wParam <> 0 Then
       If wParam = (wParam And MK_LBUTTON) Then lButton = vbLeftButton
       If wParam = (wParam And MK_MBUTTON) Then lButton = vbMiddleButton
       If wParam = (wParam And MK_RBUTTON) Then lButton = vbRightButton
       If wParam = (wParam And MK_ALT) Then lShift = vbAltMask
       If wParam = (wParam And MK_SHIFT) Then lShift = vbShiftMask
       If wParam = (wParam And MK_CONTROL) Then lShift = vbCtrlMask
     End If
     RaiseEvent MouseMove(CInt(lButton), CInt(lShift), LoWord(lParam), HiWord(lParam))
     
   Case WM_KEYDOWN, WM_KEYUP
     iKeyCode = LoWord(wParam)
     If iMsg = WM_KEYDOWN Then RaiseEvent KeyDown(iKeyCode, pvGetShiftState)
     If iMsg = WM_KEYUP Then RaiseEvent KeyUp(iKeyCode, pvGetShiftState)
     
   Case WM_CHAR
     iKeyCode = LoWord(wParam)
     RaiseEvent KeyPress(iKeyCode)
     
   Case WM_COMMAND
     lNotify = HiWord(wParam)
     If lNotify = EN_CHANGE Then RaiseEvent Change
     
 End Select
End Function
'***************************************************************************************
' Public Properties
'***************************************************************************************
'---hWnds
Public Property Get hWnd() As Long
 'The handle of the USERCONTROL (to put a PassBox in a toolbar, etc)
  hWnd = lhWnd
End Property
Public Property Get hWndEdit() As Long
 'The handle of the EDIT WINDOW (to attach Up-Down controls, etc)
  hWndEdit = lTxtWnd
End Property
'---CanCut
Public Property Get CanCut() As Boolean
 'This can be used to update a toolbar or menu
  If SelLength > 0 Then CanCut = True
End Property
'---CanCopy
Public Property Get CanCopy() As Boolean
 'This can be used to update a toolbar or menu
  If SelLength > 0 Then CanCopy = True
End Property
'---CanPaste
Public Property Get CanPaste() As Boolean
 'This can be used to update a toolbar or menu
  CanPaste = CBool(SendMessageLong(lTxtWnd, EM_CANPASTE, 0, 0))
End Property
'---CanUndo
Public Property Get CanUndo() As Boolean
 'This can be used to update a toolbar or menu
  CanUndo = CBool(SendMessageLong(lTxtWnd, EM_CANUNDO, 0, 0))
End Property
'---CueBanner
Public Property Get CueBanner() As String
  CueBanner = sCueBanner
End Property
Public Property Let CueBanner(ByVal vData As String)
  sCueBanner = vData
  PropertyChanged ("CueBanner")
  Call pvUpdateText
End Property
'---Text
Public Property Get Text() As String
Dim lLen As Long, sBuffer As String
 'Retrieve the length of the window's text
  lLen = GetWindowTextLength(lTxtWnd) + 1
 'Allocate a buffer big enough to hold the string
  sBuffer = String(lLen, vbNullChar)
 'Fill the buffer with the window's text
  Call GetWindowText(lTxtWnd, sBuffer, lLen)
  sText = Left(sBuffer, Len(sBuffer) - 1)
  Text = sText
End Property
Public Property Let Text(ByVal vData As String)
  sText = vData
  PropertyChanged ("Text")
  Call pvUpdateText
End Property
'---PasswordChar
Public Property Get PasswordChar() As String
  PasswordChar = sPassChar
End Property
Public Property Let PasswordChar(ByVal vData As String)
  sPassChar = vData
  PropertyChanged ("PasswordChar")
  Call pvUpdateStyles
End Property
'---ScrollBars
Public Property Get ScrollBars() As ScrollBarConstants
  ScrollBars = lSBars
End Property
Public Property Let ScrollBars(ByVal vData As ScrollBarConstants)
  lSBars = vData
  PropertyChanged ("ScrollBars")
  Call pvUpdateStyles
End Property
'---Enabled
Public Property Get Enabled() As Boolean
  Enabled = bEnabled
End Property
Public Property Let Enabled(ByVal vData As Boolean)
  bEnabled = vData
  PropertyChanged ("Enabled")
  Call pvUpdateStyles
End Property
'---Locked
Public Property Get Locked() As Boolean
  Locked = bLocked
End Property
Public Property Let Locked(ByVal vData As Boolean)
  bLocked = vData
  PropertyChanged ("Locked")
  Call pvUpdateStyles
End Property
'---Multiline
Public Property Get MultiLine() As Boolean
  MultiLine = bMultiLine
End Property
Public Property Let MultiLine(ByVal vData As Boolean)
  bMultiLine = vData
  PropertyChanged ("MultiLine")
  Call pvUpdateStyles
End Property
'---BackColor
Public Property Get BackColor() As OLE_COLOR
  BackColor = lBackColor
End Property
Public Property Let BackColor(ByVal vData As OLE_COLOR)
  lBackColor = vData
  PropertyChanged ("BackColor")
  Call pvUpdateStyles
End Property
'---ForeColor
Public Property Get ForeColor() As OLE_COLOR
  ForeColor = lForeColor
End Property
Public Property Let ForeColor(ByVal vData As OLE_COLOR)
  lForeColor = vData
  PropertyChanged ("ForeColor")
  Call pvUpdateStyles
End Property
'---SelStart
Public Property Get SelStart() As Long
Dim lParam As Long
 'Get the starting position
  lParam = SendMessageLong(lTxtWnd, EM_GETSEL, 0, 0)
  SelStart = LoWord(lParam)
End Property
Public Property Let SelStart(ByVal vData As Long)
 'Set the starting position
  Call SendMessageLong(lTxtWnd, EM_SETSEL, vData, vData)
End Property
'---SelLength
Public Property Get SelLength() As Long
Dim lParam As Long
 'Get the starting and ending position
  lParam = SendMessageLong(lTxtWnd, EM_GETSEL, 0, 0)
 'SelLength = Ending position - Starting position
  SelLength = HiWord(lParam) - LoWord(lParam)
End Property
Public Property Let SelLength(ByVal vData As Long)
Dim lParam As Long, lStart As Long
 'Get the starting position
  lParam = SendMessageLong(lTxtWnd, EM_GETSEL, 0, 0)
  lStart = LoWord(lParam)
 'SelLength = Starting position + Length
  Call SendMessageLong(lTxtWnd, EM_SETSEL, lStart, lStart + vData)
End Property
'***************************************************************************************
' Private Properties
'***************************************************************************************
'---WindowStyle
Private Property Get WindowStyle() As Long
Dim lNewStyle As Long
'Is it a password field?
 If sPassChar <> "" Then bPassword = True
 'Create a "Template" style
  lNewStyle = WS_CHILD Or WS_VISIBLE Or WS_TABSTOP
 'Now set the specific styles
  If bMultiLine And Not bPassword Then
   'MultiLine Edit Control
    lNewStyle = lNewStyle Or ES_MULTILINE Or ES_WANTRETURN
    Select Case lSBars
      Case vbSBNone:
        lNewStyle = lNewStyle Or ES_AUTOVSCROLL
      Case vbHorizontal
        lNewStyle = lNewStyle Or ES_AUTOHSCROLL
      Case vbVertical
        lNewStyle = lNewStyle Or ES_AUTOVSCROLL
      Case vbBoth
        lNewStyle = lNewStyle Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL
    End Select
  ElseIf Not bMultiLine And Not bPassword Then
   'SingleLine Edit Control
    lNewStyle = lNewStyle Or ES_AUTOHSCROLL
  ElseIf bPassword Then
   'Password Control
    lNewStyle = lNewStyle Or ES_AUTOHSCROLL Or ES_PASSWORD
    Call SendMessage(lTxtWnd, EM_SETPASSWORDCHAR, Asc(Left(sPassChar, 1)), 0)
  End If
 'Is it locked?
  If bLocked Then lNewStyle = lNewStyle Or ES_READONLY
 'Return the proper style
  WindowStyle = lNewStyle
End Property
'---HiWord
Private Property Get HiWord(ByVal lValue As Long) As Long
  HiWord = lValue \ &H10000
End Property
'---LoWord
Private Property Get LoWord(ByVal Value As Long) As Long
  LoWord = (Value And &HFFFF&)
End Property
'---IsXPOrAbove
Private Property Get IsXPOrAbove() As Boolean
Dim OSVer As OSVERSIONINFO
  OSVer.dwVersionInfoSize = Len(OSVer)
  GetVersionEx OSVer
  If (OSVer.dwMajorVersion > 5) Then
    IsXPOrAbove = True
  ElseIf (OSVer.dwMajorVersion = 5) Then
    If (OSVer.dwMinorVersion >= 1) Then
      IsXPOrAbove = True
    End If
  End If
End Property
'***************************************************************************************
' UserControl Events
'***************************************************************************************
'---GotFocus
Private Sub UserControl_GotFocus()
  Call SetFocus(lTxtWnd)
End Sub
'---InitProperties
Private Sub UserControl_InitProperties()
  lSBars = 0
  sText = Ambient.DisplayName
  sPassChar = ""
  bMultiLine = False
  bEnabled = True
  bLocked = False
  lBackColor = vbWindowBackground
  lForeColor = vbWindowText
  Call pvUpdateText
End Sub
'---Initialize
Private Sub UserControl_Initialize()
Dim IPAO As IOleInPlaceActiveObject
 'Set our custom IPAO
  With tIPAOHookStruct
    Set IPAO = Me
    CopyMemory .IPAOReal, IPAO, 4
    CopyMemory .TBEx, Me, 4
    .lpVTable = IPAOVTable
    .ThisPointer = VarPtr(tIPAOHookStruct)
  End With
  lBackColor = vbWindowBackground
  lForeColor = vbWindowText
  UserControl.BackColor = lBackColor
  UserControl.ForeColor = lForeColor
  Call pvCreateTextBox
End Sub
'---Terminate
Private Sub UserControl_Terminate()
  'Reset the default IPAO
   With tIPAOHookStruct
      CopyMemory .IPAOReal, 0&, 4
      CopyMemory .TBEx, 0&, 4
   End With
  Call pvDestroyTextBox
  bRunning = False
End Sub
'---ReadProperties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    sText = .ReadProperty("Text", Ambient.DisplayName)
    sPassChar = .ReadProperty("PasswordChar", "")
    sCueBanner = .ReadProperty("CueBanner", "")
    lSBars = .ReadProperty("ScrollBars", vbSBNone)
    bMultiLine = .ReadProperty("MultiLine", False)
    bEnabled = .ReadProperty("Enabled", True)
    bLocked = .ReadProperty("Locked", False)
    lBackColor = .ReadProperty("BackColor", vbWindowBackground)
    lForeColor = .ReadProperty("ForeColor", vbWindowText)
  End With
  bRunning = Ambient.UserMode
  Call pvUpdateStyles
  Call pvUpdateText
End Sub
'---WriteProperties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty("Text", sText, Ambient.DisplayName)
    Call .WriteProperty("PasswordChar", sPassChar, "")
    Call .WriteProperty("CueBanner", sCueBanner, "")
    Call .WriteProperty("ScrollBars", lSBars, vbSBNone)
    Call .WriteProperty("MultiLine", bMultiLine, False)
    Call .WriteProperty("Enabled", bEnabled, True)
    Call .WriteProperty("Locked", bLocked, False)
    Call .WriteProperty("BackColor", lBackColor, vbWindowBackground)
    Call .WriteProperty("ForeColor", lForeColor, vbWindowText)
  End With
End Sub
'---Resize
Private Sub UserControl_Resize()
Dim CurHeight As Long
  
  CurHeight = ScaleY(Extender.Height, vbContainerSize, vbPixels)
  If CurHeight < 19 Then Extender.Height = ScaleY(19, vbPixels, vbContainerSize)
  
  lhWnd = UserControl.hWnd
  lHeight = UserControl.ScaleHeight
  lWidth = UserControl.ScaleWidth
  Call pvResizeTextBox
End Sub
'***************************************************************************************
' Public Procedures
'***************************************************************************************
Public Sub ShowBalloonTip(ByVal Text As String, ByVal Title As String, Optional ByVal Icon As BalloonTipIconConstants = 0)
Dim lResult As Long
Dim tBalloonTip As EDITBALLOONTIP
  If Not IsXPOrAbove Then Exit Sub
  tBalloonTip.cbStruct = LenB(tBalloonTip)
  tBalloonTip.pszText = StrPtr(Text)
  tBalloonTip.pszTitle = StrPtr(Title)
  tBalloonTip.ttiIcon = Icon
  lResult = SendMessageW(lTxtWnd, EM_SHOWBALLOONTIP, 0, tBalloonTip)
End Sub
Public Sub HideBalloonTip()
Dim lResult As Long
  If Not IsXPOrAbove Then Exit Sub
  lResult = SendMessageLongW(lTxtWnd, EM_HIDEBALLOONTIP, 0, 0)
End Sub
'***************************************************************************************
' Private Procedures
'***************************************************************************************
'---pvResizeTextBox
Private Sub pvResizeTextBox()
  Call SetWindowPos(lTxtWnd, 0, 0, 0, lWidth, lHeight, 0)
End Sub
'---pvUpdateText
Private Sub pvUpdateText()
  Call SetWindowText(lTxtWnd, sText)
  If Not IsXPOrAbove Then Exit Sub
  Call SendMessageLongW(lTxtWnd, EM_SETCUEBANNER, 0, StrPtr(" " & sCueBanner))
End Sub
'---pvUpdateStyles
Private Sub pvUpdateStyles()
  UserControl.BackColor = lBackColor
  UserControl.ForeColor = lForeColor
  Call LockWindowUpdate(hWnd)
  Call pvDestroyTextBox
  Call pvCreateTextBox
  Call LockWindowUpdate(0)
End Sub
'---pvCreateTextBox
Private Sub pvCreateTextBox()
 'Initialize Variables
  lhWnd = UserControl.hWnd
  lHeight = UserControl.ScaleHeight
  lWidth = UserControl.ScaleWidth
 'Retrieve the appropiate style
  lStyle = WindowStyle
 'Create an "Edit" window
  lTxtWnd = CreateWindowEx(WS_EX_CLIENTEDGE Or WS_EX_NOPARENTNOTIFY, "EDIT", "", lStyle, 0, 0, lWidth, lHeight, lhWnd, 0, App.hInstance, 0)
 'Remove the scrollbars
  Call ShowScrollBar(lTxtWnd, SB_BOTH, False)
 'If it is multiline then add the scrollbars
  If bMultiLine Then
    Select Case lSBars
      Case vbHorizontal
        Call ShowScrollBar(lTxtWnd, SB_HORZ, True)
      Case vbVertical
        Call ShowScrollBar(lTxtWnd, SB_VERT, True)
      Case vbBoth
        Call ShowScrollBar(lTxtWnd, SB_BOTH, True)
    End Select
  End If
 'Set Font
  pvOLEFontToLogFont UserControl.Font, lFont
  hFont = CreateFontIndirect(lFont)
  SendMessage lTxtWnd, WM_SETFONT, hFont, 1
 'Set Text
  Call pvUpdateText
 'Subclass edit window
  If bRunning Then
    AttachMessage Me, lTxtWnd, WM_SETFOCUS
    AttachMessage Me, lTxtWnd, WM_MOUSEACTIVATE
    AttachMessage Me, lTxtWnd, WM_MOUSEMOVE
    AttachMessage Me, lTxtWnd, WM_LBUTTONUP
    AttachMessage Me, lTxtWnd, WM_LBUTTONDOWN
    AttachMessage Me, lTxtWnd, WM_LBUTTONDBLCLK
    AttachMessage Me, lTxtWnd, WM_RBUTTONUP
    AttachMessage Me, lTxtWnd, WM_RBUTTONDOWN
    AttachMessage Me, lTxtWnd, WM_RBUTTONDBLCLK
    AttachMessage Me, lTxtWnd, WM_KEYDOWN
    AttachMessage Me, lTxtWnd, WM_KEYUP
    AttachMessage Me, lTxtWnd, WM_CHAR
    AttachMessage Me, lhWnd, WM_COMMAND
  End If
  'Uncomment for filename auto completion
  'pvSetAutoComplete lTxtWnd, SHACF_DEFAULT
End Sub
'---pvDestroyTextBox
Private Sub pvDestroyTextBox()
 'Destroy created windows
  If lTxtWnd <> 0 Then DestroyWindow lTxtWnd
  DeleteObject hFont
 'Unubclass edit window
  DetachMessage Me, lTxtWnd, WM_SETFOCUS
  DetachMessage Me, lTxtWnd, WM_MOUSEACTIVATE
  DetachMessage Me, lTxtWnd, WM_MOUSEMOVE
  DetachMessage Me, lTxtWnd, WM_LBUTTONUP
  DetachMessage Me, lTxtWnd, WM_LBUTTONDOWN
  DetachMessage Me, lTxtWnd, WM_LBUTTONDBLCLK
  DetachMessage Me, lTxtWnd, WM_RBUTTONUP
  DetachMessage Me, lTxtWnd, WM_RBUTTONDOWN
  DetachMessage Me, lTxtWnd, WM_RBUTTONDBLCLK
  DetachMessage Me, lTxtWnd, WM_KEYDOWN
  DetachMessage Me, lTxtWnd, WM_KEYUP
  DetachMessage Me, lTxtWnd, WM_CHAR
  DetachMessage Me, lhWnd, WM_COMMAND
End Sub
'---pvSetAutoComplete
Private Function pvSetAutoComplete(ByVal hWnd As Long, ByVal eFlags As Long)
Dim lR As Long
  lR = SHAutoComplete(hWnd, eFlags)
  pvSetAutoComplete = (lR <> S_OK)
End Function
'---pvOLEFontToLogFont
Private Sub pvOLEFontToLogFont(fntThis As StdFont, tLF As LOGFONT)
Dim sFont As String
Dim iChar As Integer
  With tLF
    sFont = fntThis.Name
    For iChar = 1 To Len(sFont)
      .lfFaceName(iChar - 1) = CByte(Asc(Mid(sFont, iChar, 1)))
    Next iChar
    .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
    .lfItalic = fntThis.Italic
    If (fntThis.Bold) Then
      .lfWeight = FW_BOLD
    Else
      .lfWeight = FW_NORMAL
    End If
    .lfUnderline = fntThis.Underline
    .lfStrikeOut = fntThis.Strikethrough
    .lfCharSet = fntThis.Charset
   'DEFAULT_QUALITY means that will support cleartype
   'with capable fonts (tahoma, verdana, etc)
    .lfQuality = DEFAULT_QUALITY
  End With
End Sub
'---pvGetShiftState
Private Function pvGetShiftState() As ShiftConstants
Dim iR As Integer
  iR = iR Or (-1 * pvKeyIsPressed(vbKeyShift))
  iR = iR Or (-2 * pvKeyIsPressed(vbKeyMenu))
  iR = iR Or (-4 * pvKeyIsPressed(vbKeyControl))
  pvGetShiftState = iR
End Function
'---pvKeyIsPressed
Private Function pvKeyIsPressed(ByVal nVirtKeyCode As KeyCodeConstants) As Boolean
Dim lR As Long
  lR = GetAsyncKeyState(nVirtKeyCode)
  If (lR And &H8000&) = &H8000& Then
    pvKeyIsPressed = True
  End If
End Function
'***************************************************************************************
' Other Procedures
'***************************************************************************************
'---TranslateAccelerator
Friend Function TranslateAccelerator(lpMsg As VBOleGuids.Msg) As Long
   TranslateAccelerator = S_FALSE
   If lpMsg.message = WM_KEYDOWN Then
      Select Case lpMsg.wParam And &HFFFF&
      Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageDown, vbKeyPageUp, vbKeyHome, vbKeyEnd
         SendMessageLong lTxtWnd, lpMsg.message, lpMsg.wParam, lpMsg.lParam
         TranslateAccelerator = S_OK
      End Select
   End If
End Function
'***************************************************************************************
