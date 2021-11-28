VERSION 5.00
Begin VB.UserControl SSTabEx 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ControlContainer=   -1  'True
   PropertyPages   =   "ctlSSTabEx.ctx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   ToolboxBitmap   =   "ctlSSTabEx.ctx":0059
   Begin VB.Timer tmrCheckDuplicationByIDEPaste 
      Interval        =   1
      Left            =   792
      Top             =   1548
   End
   Begin VB.Timer tmrTabHoverEffect 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   396
      Top             =   2268
   End
   Begin VB.Timer tmrSubclassControls 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   396
      Top             =   1908
   End
   Begin VB.Timer tmrCancelDoubleClick 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   396
      Top             =   1548
   End
   Begin VB.Timer tmrCheckContainedControlsAdditionDesignTime 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   36
      Top             =   2268
   End
   Begin VB.PictureBox picAux2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   1944
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   5
      Top             =   684
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.PictureBox picInactiveTabBodyThemed 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   972
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   4
      Top             =   684
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.PictureBox picTabBodyThemed 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   0
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   3
      Top             =   684
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.PictureBox picAux 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   1944
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.Timer tmrDraw 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   36
      Top             =   1908
   End
   Begin VB.Timer tmrTabMouseLeave 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   36
      Top             =   1548
   End
   Begin VB.PictureBox picRotate 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   972
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   624
      Left            =   0
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   912
   End
End
Attribute VB_Name = "SSTabEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Uncomment the line below for IDE protection when running uncompiled (some features will be lost in the IDE when it is uncommented)
' #Const NOSUBCLASSINIDE = True

Implements IBSSubclass

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type WINDOWPOS
   hWnd As Long
   hWndInsertAfter As Long
   X As Long
   Y As Long
   cx As Long
   cy As Long
   Flags As Long
End Type

'Bitmap type used to store Bitmap Data
Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Private Type PAINTSTRUCT
    hDC                     As Long
    fErase                  As Long
    rcPaint                 As RECT
    fRestore                As Long
    fIncUpdate              As Long
    rgbReserved(1 To 32)    As Byte
End Type

Private Type T_MSG
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Type tagHIGHCONTRAST
    cbSize As Long
    dwFlags As Long
    lpszDefaultScheme As Long
End Type

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As tagHIGHCONTRAST, ByVal fuWinIni As Long) As Long
      
Private Const SPI_GETHIGHCONTRAST As Long = 66
Private Const HCF_HIGHCONTRASTON As Long = 1

Private Declare Function SetLayout Lib "gdi32" (ByVal hDC As Long, ByVal dwLayout As Long) As Long
Private Const LAYOUT_RTL = &H1                               ' Right to left
Private Const LAYOUT_BITMAPORIENTATIONPRESERVED = &H8

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetMessageExtraInfo Lib "user32" () As Long
Private Const MOUSEEVENTF_LEFTDOWN = &H2 ' Left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 ' Left button up

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const SM_SWAPBUTTON = 23&

Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As T_MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Const PM_REMOVE = &H1

Private Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_INVALIDATE = &H1
Private Const RDW_UPDATENOW = &H100

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_THEMECHANGED As Long = &H31A
Private Const WM_PAINT As Long = &HF
Private Const WM_MOVE As Long = &H3&
Private Const WM_MOUSEACTIVATE As Long = &H21
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_SETREDRAW As Long = &HB&
Private Const WM_USER As Long = &H400
Private Const WM_DRAW As Long = WM_USER + 10 ' custom message
Private Const WM_INIT As Long = WM_USER + 11 ' custom message
Private Const WM_LBUTTONDBLCLK As Long = &H203&
Private Const WM_PRINTCLIENT As Long = &H318
Private Const WM_NCACTIVATE As Long = &H86&
Private Const WM_WINDOWPOSCHANGING = &H46&
Private Const WM_GETDPISCALEDSIZE As Long = &H2E4&

'Private Const MA_NOACTIVATEANDEAT As Long = &H4
Private Const WM_MOUSELEAVE As Long = &H2A3

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Const GA_ROOT = 2

Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObjectA Lib "gdi32.dll" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
Private Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'Draw Text Constants
Private Const DT_CALCRECT = &H400&
Private Const DT_CENTER = &H1&
Private Const DT_SINGLELINE = &H20&
Private Const DT_VCENTER = &H4&
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_WORDBREAK = &H10&
Private Const DT_RTLREADING As Long = &H20000

Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long

Private Const TABP_TABITEM = 1
Private Const TABP_TABITEMLEFTEDGE = 2
Private Const TABP_TABITEMRIGHTEDGE = 3
'Private Const TABP_TABITEMBOTHEDGE = 4
'Private Const TABP_TOPTABITEM = 5
'Private Const TABP_TOPTABITEMLEFTEDGE = 6
'Private Const TABP_TOPTABITEMRIGHTEDGE = 7
'Private Const TABP_TOPTABITEMBOTHEDGE = 8
Private Const TABP_PANE = 9
'Private Const TABP_BODY = 10

Private Const TIS_NORMAL = 1
Private Const TIS_HOT = 2
Private Const TIS_SELECTED = 3
Private Const TIS_DISABLED = 4
Private Const TIS_FOCUSED = 5

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long

Private Const HALFTONE = 4
Private Type COLORADJUSTMENT
        caSize As Integer
        caFlags As Integer
        caIlluminantIndex As Integer
        caRedGamma As Integer
        caGreenGamma As Integer
        caBlueGamma As Integer
        caReferenceBlack As Integer
        caReferenceWhite As Integer
        caContrast As Integer
        caBrightness As Integer
        caColorfulness As Integer
        caRedGreenTint As Integer
End Type

Private Declare Function GetColorAdjustment Lib "gdi32" (ByVal hDC As Long, lpca As COLORADJUSTMENT) As Long
Private Declare Function SetColorAdjustment Lib "gdi32" (ByVal hDC As Long, lpca As COLORADJUSTMENT) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type


Private Enum vbExMouseButtonsConstants
    vxMBLeft = 1&
    vxMBRight = 2&
End Enum

Private Declare Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
Private Declare Function IsAppThemed Lib "uxtheme" () As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Long
Private Declare Function GetThemeAppProperties Lib "uxtheme" () As Long

Private Const S_OK As Long = &H0
Private Const STAP_ALLOW_CONTROLS As Long = (1 * (2 ^ 1))

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszCaption As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Private Const cAuxTransparentColor As Long = &HFF01FF ' Not the MaskColor, but another transparent color for internal operations
Private Const cTabPictureDistanceToCaption As Long = 3

Private Enum efnRotatePicDirection
    efn90DegreesClockWise = 0
    efn90DegreesCounterClockWise = 1
    efnFlipVertical = 2
    efnFlipHorizontal = 3
End Enum

' Public Enums
Public Enum vbExTabOrientationConstants
    ssTabOrientationTop = 0
    ssTabOrientationBottom = 1
    ssTabOrientationLeft = 2
    ssTabOrientationRight = 3
End Enum

Public Enum vbExMousePointerConstants
    ssDefault = 0
    ssArrow = 1
    ssCross = 2
    ssIBeam = 3
    ssIcon = 4
    ssSize = 5
    ssSizeNESW = 6
    ssSizeNS = 7
    ssSizeNWSE = 8
    ssSizeEW = 9
    ssUpArrow = 10
    ssHourglass = 11
    ssNoDrop = 12
    ssArrowHourglass = 13
    ssArrowQuestion = 14
    ssSizeAll = 15
    ssCustom = 99
End Enum

Public Enum vbExOLEDropConstants
    ssOLEDropNone = 0
    ssOLEDropManual = 1
End Enum

Public Enum vbExStyleConstants
    ssStyleTabbedDialog = 0
    ssStylePropertyPage = 1
    ssStyleTabStrip = 2
End Enum

Public Enum vbExAutoYesNoConstants
    ssNo = 0
    ssYes = 1
    ssYNAuto = 2
End Enum

Public Enum vbExTabWidthStyleConstants
    ssTWSJustified = 0
    ssTWSNonJustified = 1
    ssTWSFixed = 2
    ssTWSAuto = 3
End Enum

Public Enum vbExTabAppearanceConstants
    ssTAAuto = 0
    ssTATabbedDialog = 1
    ssTATabbedDialogRounded = 2
    ssTAPropertyPage = 3
    ssTAPropertyPageRounded = 4
End Enum

Public Enum vbExTabPictureAlignmentConstants
    ssPicAlignBeforeCaption = 0
    ssPicAlignCenteredBeforeCaption = 1
    ssPicAlignAfterCaption = 2
    ssPicAlignCenteredAfterCaption = 3
End Enum

Public Enum vbExAutoRelocateControlsConstants
    ssRelocateNever = 0
    ssRelocateAlways = 1
    ssRelocateOnTabOrientationChange = 2
End Enum

Public Enum vbExTabHoverHighlightConstants
    ssTHHNo = 0
    ssTHHInstant = 1
    ssTHHEffect = 2
End Enum

Public Enum vbExBackStyleConstants
    ssTransparent = 0
    ssOpaque = 1
End Enum

' Events
' Original
Public Event Click(ByVal PreviousTab As Integer)
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Attribute DblClick.VB_UserMemId = -601
Public Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyPress(ByVal KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Public Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Public Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs when a source component is dropped onto a target component, informing the source component that a drag action was either performed or canceled."
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when a source component is dropped onto a target component  when the source component determines that a drop can occur."
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when one component is dragged over another."
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs after every OLEDragOver event."
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs on an source component when a target component performs the GetData method on the sources DataObject object, but the data for the specified format has not yet been loaded."
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when a component's OLEDrag method is performed, or when a component initiates an OLE drag/drop operation when the OLEDragMode property is set to Automatic."

' Added
Public Event BeforeClick(ByRef Cancel As Boolean)
Attribute BeforeClick.VB_Description = "Occurs when the current tab is about to change."
Public Event ChangeControlBackColor(ByVal ControlName As String, ByVal ControlTypeName As String, ByRef Cancel As Boolean)
Public Event RowsChange()
Attribute RowsChange.VB_Description = "Occurs when the Rows property changes its value."
Public Event TabBodyResize()
Attribute TabBodyResize.VB_Description = "Occurs when the tab body changes its size."
Public Event TabMouseEnter(ByVal nTab As Integer)
Public Event TabMouseLeave(ByVal nTab As Integer)
Public Event TabRightClick(ByVal nTab As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event TabSelChange()
Attribute TabSelChange.VB_Description = "Occurs after the current tab has already changed."
Public Event Resize()
Attribute Resize.VB_Description = "Occurs when the control is first drawn or when its size changes."


Private Type T_TabData
    ' Properties
    Caption As String
    Enabled As Boolean
    Visible As Boolean
    Picture As StdPicture
    Pic16 As StdPicture
    Pic20 As StdPicture
    Pic24 As StdPicture
    ToolTipText As String
    Controls As Collection
    ' Run time data
    TabRect As RECT
    PicToUse As StdPicture
    PicToUseSet As Boolean
    PicDisabled As StdPicture
    PicDisabledSet As Boolean
    Hovered As Boolean
    Selected As Boolean
    LeftTab As Boolean
    RightTab As Boolean
    TopTab As Boolean
    PicAndCaptionWidth As Long
    Row As Long
    RowPos As Long
    PosH As Long
    Width As Long
End Type

Private Const cDefaultTabHeight = 238 ' in Twips
Private Const cRowPerspectiveSpace = 150& ' in Twips

' Variables for properties
' Original
Private mBackColor As Long
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
Private mEnabled As Boolean
Private mForeColor As Long
Private mUserControlHwnd As Long ' read only
Private mTabSel As Integer
Private mTabsPerRow As Integer
Private mTabs As Integer
Private mRows As Integer ' read only
Private mTabOrientation As vbExTabOrientationConstants
Private mShowFocusRect As Boolean
Private mWordWrap As Boolean
Private mStyle As vbExStyleConstants
Private mTabHeight As Single ' internally in Himetric
Private mTabMaxWidth As Single ' internally in Himetric
Private mMousePointer As vbExMousePointerConstants
Private mMouseIcon As StdPicture
Private mOLEDropMode As vbExOLEDropConstants
Private mTabData() As T_TabData
Private mRightToLeft As Boolean

' Added
Private mMaskColor As Long
Private mUseMaskColor As Boolean
Private mTabSelExtraHeight As Single ' internally  in Himetric
Private mTabSelFontBold As vbExAutoYesNoConstants
Private mTabSelHighlight As Boolean
Private mTabHoverHighlight As vbExTabHoverHighlightConstants
Private mVisualStyles As Boolean
Private mTabBackColor As Long
Private mTabSelBackColor As Long
Private mTabBackColor_SavedWhileVisualStyles As Long
Private mTabSelBackColor_SavedWhileVisualStyles As Long
Private mTabBackColorSavedWhileVisualStyles As Boolean
Private mTabSelForeColor As Long
Private mShowDisabledState As Boolean
Private mTabBodyRect As RECT ' internally in Pixels, red only
Private mChangeControlsBackColor As Boolean
Private mTabMinWidth As Single ' internally in Himetric
Private mTabWidthStyle As vbExTabWidthStyleConstants
Private mShowRowsInPerspective As vbExAutoYesNoConstants
Private mTabSeparation As Integer
Private mForceVisualStyles As Boolean
Private mTabAppearance As vbExTabAppearanceConstants
Private mRedraw As Boolean
Private mTabPictureAlignment As vbExTabPictureAlignmentConstants
Private mAutoRelocateControls As vbExAutoRelocateControlsConstants
Private mEndOfTabs As Single
Private mSoftEdges As Boolean
Private mMinSpaceNeeded As Single
Private mHandleHighContrastTheme As Boolean
Private mBackStyle As vbExBackStyleConstants
Private mAutoTabHeight As Boolean

' Variables
Private mTabBodyStart As Long ' in Pixels
Private mTabBodyHeight As Long ' in Pixels
Private mTabBodyWidth As Long ' in Pixels
Private mScaleWidth As Long
Private mScaleHeight As Long
Private mHasFocus As Boolean
Private mFormIsActive As Boolean
Private mDrawing As Boolean
Private mTabUnderMouse As Integer
Private mAmbientUserMode As Boolean
Private mExtenderToolTipText As String
Private mLastTabToolTipTextSet As String
Private mThereAreTabsToolTipTexts As Boolean
Private mDefaultTabHeight As Single  ' in Himetric
Private mPropertiesReady As Boolean
Private mButtonFace_H As Integer
Private mButtonFace_L As Integer
Private mButtonFace_S As Integer
Private mTabBodyThemedReady As Boolean
Private mInactiveTabBodyThemedReady As Boolean
Private mTabBodyWidth_Prev As Long
Private mTabBodyHeight_Prev As Long
Private mTheme As Long
Private mControlIsThemed As Boolean
Private mTabSeparation2 As Long
Private mThemeExtraDataAlreadySet As Boolean
Private mParentControlsTabStop As Collection
Private mParentControlsUseMnemonic As Collection
Private mContainedControlsThatAreContainers As Collection
Private mSubclassedControlsForPaintingHwnds As Collection
Private mSubclassedFramesHwnds As Collection
Private mSubclassedControlsForMoveHwnds As Collection
Private mTabStopsInitialized As Boolean
Private mAccessKeys As String
Private mAccessKeysSet As Boolean
Private mBlendDisablePicWithTabBackColor_NotThemed As Boolean
Private mBlendDisablePicWithTabBackColor_Themed As Boolean
Private mSubclassControlsPaintingPending As Boolean
Private mRepaintSubclassedControls As Boolean
Private mFormHwnd As Long
Private mBtnDown As Boolean
Private mTabAppearance2 As vbExTabAppearanceConstants
Private mAppearanceIsPP As Boolean
Private mNoActivate As Boolean
Private mCanPostDrawMessage As Boolean
Private mDrawMessagePosted As Boolean
Private mNeedToDraw As Boolean
Private mRows_Prev As Integer
Private mChangedControlsBackColor As Boolean
Private mLastContainedControlsString As String
Private mLastContainedControlsCount As Long
Private mLastContainedControlsPositionsStr As String
Private mTabBodyReset As Boolean
Private mSubclassed As Boolean
Private mTabBodyStart_Prev As Long
Private mTabOrientation_Prev As vbExTabOrientationConstants
Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private mFirstDraw As Boolean
Private mUserControlShown As Boolean
Private mTabBodyRect_Prev As RECT
Private mEnsureDrawn As Boolean
Private mDPIX As Long
Private mDPIY As Long
Private mXCorrection As Single
Private mYCorrection As Single
Private mHoverEffectColors(5) As Long
Private mTabHoverEffect_Step As Long
Private mGlowColor_Bk As Long
Private mGlowColor_Sel_Bk As Long
Private mHighContrastThemeOn As Boolean
Private mHandleHighContrastTheme_OrigForeColor As Long
Private mHandleHighContrastTheme_OrigTabBackColor As Long
Private mHandleHighContrastTheme_OrigTabSelForeColor As Long
Private mHandleHighContrastTheme_OrigTabSelBackColor As Long
Private mBackColorIsfromAmbient As Boolean
Private mForeColorIsfromAmbient As Boolean
Private mTabBackColorIsfromAmbient As Boolean
Private mLeftShiftToHide As Long
Private mLeftThresholdHided As Long
Private mPendingLeftShift As Long
Private mUserControlTerminated As Boolean

' Colors
Private m3DDKShadow As Long
Private m3DHighlight As Long
Private m3DShadow As Long
Private m3DDKShadow_Sel As Long
Private m3DHighlight_Sel As Long
Private m3DShadow_Sel As Long
Private mTabBackColorDisabled As Long
Private mTabSelBackColorDisabled As Long
Private mGrayText As Long
Private mGrayText_Sel As Long
Private mGlowColor As Long
Private mGlowColor_Sel As Long
Private mTabBackColor_R As Long
Private mTabBackColor_G As Long
Private mTabBackColor_B As Long
Private mTabSelBackColor_R As Long
Private mTabSelBackColor_G As Long
Private mTabSelBackColor_B As Long

Private m3DShadowH As Long
Private m3DShadowV As Long
Private m3DShadowH_Sel As Long
Private m3DShadowV_Sel As Long
Private m3DHighlightH As Long
Private m3DHighlightV As Long
Private m3DHighlightH_Sel As Long
Private m3DHighlightV_Sel As Long
Private mTabBackColor2 As Long
Private mTabSelBackColor2 As Long

' Themed extra data
Private mThemedInactiveReferenceTabBackColor As Long
Private mThemedInactiveReferenceTabBackColor_H As Integer
Private mThemedInactiveReferenceTabBackColor_L As Integer
Private mThemedInactiveReferenceTabBackColor_S As Integer
Private mThemedTabBodyReferenceTopBackColor As Long
Private mTABITEM_TopLeftCornerTransparencyMask(5) As Long
Private mTABITEM_TopRightCornerTransparencyMask(5) As Long
Private mTABITEMRIGHTEDGE_RightSideTransparencyMask(5) As Long
Private mThemedTabBodyBottomShadowPixels As Long
Private mThemedTabBodyRightShadowPixels As Long
Private mThemedTabBodyBackColor_R As Long
Private mThemedTabBodyBackColor_G As Long
Private mThemedTabBodyBackColor_B As Long


' Properties

' Returns/sets the background color.
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
Attribute BackColor.VB_MemberFlags = "c"
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        mBackColor = nValue
        PropertyChanged "BackColor"
        ResetCachedThemeImages
        Draw
    End If
End Property


' Returns a Font object.
Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = mFont
End Property

Public Property Let Font(ByVal nValue As StdFont)
    Set Font = nValue
End Property

Public Property Set Font(ByVal nValue As StdFont)
    If Not nValue Is mFont Then
        Set mFont = nValue
        PropertyChanged "Font"
        SetFont
        SetAutoTabHeight
        Draw
    End If
End Property

Private Sub SetFont()
    On Error Resume Next
    If mFont Is Nothing Then
        Set mFont = Ambient.Font
    End If
    If mFont Is Nothing Then
        Set mFont = UserControl.Font
    End If
    Set UserControl.Font = mFont
    Set picDraw.Font = mFont
    Set picAux.Font = mFont
    Err.Clear
End Sub

' Determines if the control is enabled.
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns or sets a value that determines whether a form or control can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal nValue As Boolean)
    Dim iRedraw As Boolean
    Dim iWv As Boolean
    
    If nValue <> mEnabled Then
        mEnabled = nValue
        UserControl.Enabled = mEnabled Or (Not mAmbientUserMode)
        PropertyChanged "Enabled"
        If mChangeControlsBackColor Then
            If mShowDisabledState Then
                mTabBodyReset = True
                iWv = IsWindowVisible(mUserControlHwnd) <> 0
                If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
                If mEnabled Then
                    SetControlsBackColor mTabSelBackColor, mTabSelBackColorDisabled
                Else
                    SetControlsBackColor mTabSelBackColorDisabled, mTabSelBackColor
                End If
                If iWv Then
                    SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
                    iRedraw = True
                End If
            End If
        End If
        mSubclassControlsPaintingPending = True
        mRepaintSubclassedControls = True
        Draw
        If iRedraw Then
            RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
        End If
    End If
End Property

            
' Returns/sets the text color.
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
Attribute ForeColor.VB_MemberFlags = "c"
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        ForeColor = mHandleHighContrastTheme_OrigForeColor
    Else
        ForeColor = mForeColor
    End If
End Property

Public Property Let ForeColor(ByVal nValue As OLE_COLOR)
    Dim iPrev As Long
    
    If nValue <> mForeColor Then
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigForeColor = nValue
        Else
            iPrev = mForeColor
            mForeColor = nValue
            PropertyChanged "ForeColor"
            If mTabSelForeColor = iPrev Then
                TabSelForeColor = nValue
            Else
                Draw
            End If
        End If
    End If
End Property


Public Property Get TabSelForeColor() As OLE_COLOR
Attribute TabSelForeColor.VB_Description = "Returns/sets the caption color of the active tab."
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        TabSelForeColor = mHandleHighContrastTheme_OrigTabSelForeColor
    Else
        TabSelForeColor = mTabSelForeColor
    End If
End Property

Public Property Let TabSelForeColor(ByVal nValue As OLE_COLOR)
    If nValue <> mTabSelForeColor Then
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigTabSelForeColor = nValue
        Else
            mTabSelForeColor = nValue
            PropertyChanged "TabSelForeColor"
            Draw
        End If
    End If
End Property

' Returns/sets the text displayed in the active tab.
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in the active tab."
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "c"
    Caption = mTabData(mTabSel).Caption
End Property

Public Property Let Caption(ByVal nValue As String)
    TabCaption(mTabSel) = nValue
End Property


' Returns the Windows handle of the control.
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the Windows handle of the control."
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = mUserControlHwnd
End Property


' Returns/sets the number of tabs to appear on each row.
Public Property Get TabsPerRow() As Integer
Attribute TabsPerRow.VB_Description = "Returns/sets the number of tabs to appear on each row."
    TabsPerRow = mTabsPerRow
End Property

Public Property Let TabsPerRow(ByVal nValue As Integer)
    If nValue <> mTabsPerRow Then
        mTabsPerRow = nValue
        PropertyChanged "TabsPerRow"
        Draw
    End If
End Property


' Returns/sets the number of tabs.
Public Property Get Tabs() As Integer
Attribute Tabs.VB_Description = "Returns/sets the number of tabs."
    Tabs = mTabs
End Property

Public Property Let Tabs(ByVal nValue As Integer)
    Dim c As Long
    
    If nValue <> mTabs Then
        PropertyChanged "Tabs"
        If mTabs > nValue Then
            For c = nValue To mTabs - 1
                If mTabData(c).Controls.Count > 0 Then
                    On Error Resume Next
                    Err.Clear
                    Err.Raise 380  '  invalid property value
                    Dim iStr As String
                    iStr = Err.Description
                    On Error GoTo 0
                    RaiseError 380, TypeName(Me), iStr & ". Tab " & CStr(c) & " has controls, can't remove tabs with controls. Remove the contained controls first."
                    Exit Property
                End If
            Next c
        End If
        ReDim Preserve mTabData(nValue - 1)
        If mTabs < nValue Then
            For c = mTabs To nValue - 1
                Set mTabData(c).Controls = New Collection
                mTabData(c).Enabled = True
                mTabData(c).Visible = True
                mTabData(c).Caption = "Tab " & CStr(c)
            Next
        End If
        mTabs = nValue
        If mTabSel > (mTabs - 1) Then
            mTabSel = (mTabs - 1)
        End If
        Draw
    End If
End Property


' Returns the number of rows of tabs.
Public Property Get Rows() As Integer
Attribute Rows.VB_Description = "Returns the number of rows of tabs."
Attribute Rows.VB_MemberFlags = "400"
    Rows = mRows
End Property

' Returns/sets the active tab number.
Public Property Get TabSel() As Integer
Attribute TabSel.VB_Description = "Returns/sets the active tab number."
    TabSel = mTabSel
End Property

Public Property Let TabSel(ByVal nValue As Integer)
    Dim iPrev As Integer
    Dim iPrev2 As Integer
    Dim iCancel As Boolean
    Dim iWv As Boolean
    
    If (nValue < 0) Or (nValue >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If Not mTabData(nValue).Visible Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    
    If nValue <> mTabSel Then
        RaiseEvent BeforeClick(iCancel)
        If iCancel = 0 Then
            iPrev = mTabSel
            mTabSel = nValue
            PropertyChanged "TabSel"
            If (iPrev >= 0) And (iPrev <= UBound(mTabData)) Then
                mTabData(iPrev).Selected = False
            End If
            If (mTabSel > -1) And (mTabSel < mTabs) Then
                mTabData(mTabSel).Selected = True
            End If
            iPrev2 = iPrev
            RaiseEvent Click(iPrev)
            iWv = IsWindowVisible(mUserControlHwnd) <> 0
            If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
            SetVisibleControls iPrev2
            If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
            mSubclassControlsPaintingPending = True
            If tmrTabHoverEffect.Enabled Then
                tmrTabHoverEffect.Enabled = False
                mGlowColor = mGlowColor_Sel_Bk
            End If
            Draw
            If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
            RaiseEvent TabSelChange
        End If
    End If
End Property


' Returns/sets a value that determines which side of the control the tabs will appear.
Public Property Get TabOrientation() As vbExTabOrientationConstants
Attribute TabOrientation.VB_Description = "Returns/sets a value that determines which side of the control the tabs will appear."
    TabOrientation = mTabOrientation
End Property

Public Property Let TabOrientation(ByVal nValue As vbExTabOrientationConstants)
    If nValue < 0 Or nValue > 3 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mTabOrientation Then
        mTabOrientation = nValue
        PropertyChanged "TabOrientation"
        ResetCachedThemeImages
        Draw
    End If
End Property


Public Property Get TabPictureAlignment() As vbExTabPictureAlignmentConstants
Attribute TabPictureAlignment.VB_Description = "Returns/sets the alignment of the tab picture with respect of the tab caption."
    TabPictureAlignment = mTabPictureAlignment
End Property

Public Property Let TabPictureAlignment(ByVal nValue As vbExTabPictureAlignmentConstants)
    If nValue < 0 Or nValue > 3 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mTabPictureAlignment Then
        mTabPictureAlignment = nValue
        PropertyChanged "TabPictureAlignment"
        DrawDelayed
    End If
End Property


' Specifies a bitmap to display on the current tab.
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Specifies a bitmap or icon to display on the current tab."
    Set Picture = TabPicture(mTabSel)
End Property

Public Property Let Picture(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPicture(mTabSel) = nValue
End Property

Public Property Set Picture(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPicture(mTabSel) = nValue
End Property


Public Property Get Pic16() As Picture
Attribute Pic16.VB_Description = "Specifies a bitmap to display on the current tab at 96 DPI, when the application is DPI aware."
    Set Pic16 = TabPic16(mTabSel)
End Property

Public Property Let Pic16(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic16(mTabSel) = nValue
End Property

Public Property Set Pic16(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic16(mTabSel) = nValue
End Property


Public Property Get Pic20() As Picture
Attribute Pic20.VB_Description = "Specifies a bitmap to display on the current tab at 120 DPI, when the application is DPI aware."
    Set Pic20 = TabPic20(mTabSel)
End Property

Public Property Let Pic20(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic20(mTabSel) = nValue
End Property

Public Property Set Pic20(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic20(mTabSel) = nValue
End Property


Public Property Get Pic24() As Picture
Attribute Pic24.VB_Description = "Specifies a bitmap to display on the current tab at 144 DPI, when the application is DPI aware."
    Set Pic24 = TabPic24(mTabSel)
End Property

Public Property Let Pic24(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic24(mTabSel) = nValue
End Property

Public Property Set Pic24(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    Set TabPic24(mTabSel) = nValue
End Property


' Determines whether a focus rectangle will be drawn in the caption when the control has the focus.
Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_Description = "Determines whether a focus rectangle will be drawn in the caption when the control has the focus."
    ShowFocusRect = mShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal nValue As Boolean)
    If nValue <> mShowFocusRect Then
        mShowFocusRect = nValue
        PropertyChanged "ShowFocusRect"
        Draw
    End If
End Property


' Determines whether text in the caption of each tab will wrap to the next line if it is too long.
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Determines whether text in the caption of each tab will wrap to the next line if it is too long."
    WordWrap = mWordWrap
End Property

Public Property Let WordWrap(ByVal nValue As Boolean)
    If nValue <> mWordWrap Then
        mWordWrap = nValue
        PropertyChanged "WordWrap"
        Draw
    End If
End Property


' Returns/sets the style of the tabs.
Public Property Get Style() As vbExStyleConstants
Attribute Style.VB_Description = "Returns/sets the style of the tabs."
    Style = mStyle
End Property

Public Property Let Style(ByVal nValue As vbExStyleConstants)
    Dim iStyle As vbExStyleConstants
    
    If nValue < 0 Or nValue > 2 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    iStyle = nValue
    If iStyle <> mStyle Then
        mStyle = iStyle
        PropertyChanged "Style"
        Draw
    End If
End Property


' Returns/sets the height of the tabs.
Public Property Get TabHeight() As Single
Attribute TabHeight.VB_Description = "Returns/sets the height of tabs."
    TabHeight = FixRoundingError(ToContainerSizeY(mTabHeight, vbHimetric))
End Property

Public Property Let TabHeight(ByVal nValue As Single)
    Dim iValue As Single
    
    iValue = FromContainerSizeY(nValue, vbHimetric)
    If (iValue < 1) Or (pScaleY(iValue, vbHimetric, vbTwips) > IIf(UserControl.Height > 2000, UserControl.Height, 2000)) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If pScaleY(iValue, vbHimetric, vbPixels) < 1 Then iValue = pScaleY(1, vbPixels, vbHimetric)
    If Round(iValue * 10000) <> Round(mTabHeight * 10000) Then
        If Abs(Round(iValue) - Round(mTabHeight)) > 1 Then
            If Round(iValue) <> Round(mDefaultTabHeight) Then
                mAutoTabHeight = False
                PropertyChanged "AutoTabHeight"
            End If
        End If
        mTabHeight = iValue
        If mTabSelExtraHeight > mTabHeight Then
            TabSelExtraHeight = mTabHeight
        End If
        PropertyChanged "TabHeight"
        ResetCachedThemeImages
        Draw
    End If
End Property


' Returns/sets the maximum width of each tab.
Public Property Get TabMaxWidth() As Single
Attribute TabMaxWidth.VB_Description = "Returns/sets the maximum width of each tab."
    TabMaxWidth = FixRoundingError(ToContainerSizeX(mTabMaxWidth, vbHimetric))
End Property

Public Property Let TabMaxWidth(ByVal nValue As Single)
    Dim iValue As Single
    
    iValue = FromContainerSizeX(nValue, vbHimetric)
    If ((iValue < pScaleX(10, vbPixels, vbHimetric)) And Not iValue = 0) Or (pScaleX(iValue, vbHimetric, vbTwips) > IIf(UserControl.Width > 3000, UserControl.Width, 3000)) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If Round(iValue * 10000) <> Round(mTabMaxWidth * 10000) Then
        mTabMaxWidth = iValue
        If mTabMaxWidth <> 0 Then
            If mTabMaxWidth < mTabMinWidth Then
                TabMinWidth = ToContainerSizeY(mTabMaxWidth, vbHimetric)
            End If
        End If
        PropertyChanged "TabMaxWidth"
        Draw
    End If
End Property


' Returns/sets the minimun width of each tab.
Public Property Get TabMinWidth() As Single
Attribute TabMinWidth.VB_Description = "Returns/sets the minimun width of each tab."
    TabMinWidth = FixRoundingError(ToContainerSizeX(mTabMinWidth, vbHimetric))
End Property

Public Property Let TabMinWidth(ByVal nValue As Single)
    Dim iValue As Single
    
    iValue = FromContainerSizeX(nValue, vbHimetric)
    If (iValue < 0) Or (pScaleX(iValue, vbHimetric, vbTwips) > IIf(UserControl.Width > 3000, UserControl.Width, 3000)) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If Round(iValue * 10000) <> Round(mTabMinWidth * 10000) Then
        mTabMinWidth = iValue
        If (mTabMinWidth > mTabMaxWidth) And (mTabMaxWidth <> 0) Then
            TabMaxWidth = ToContainerSizeY(mTabMinWidth, vbHimetric)
        End If
        PropertyChanged "TabMinWidth"
        Draw
    End If
End Property


Public Property Get TabWidthStyle() As vbExTabWidthStyleConstants
Attribute TabWidthStyle.VB_Description = "Returns/sets a value that determines whether the color assigned in the MaskColor property is used as a mask for setting transparent regions in the tab pictures."
Attribute TabWidthStyle.VB_MemberFlags = "400"
    TabWidthStyle = mTabWidthStyle
End Property

Public Property Let TabWidthStyle(ByVal nValue As vbExTabWidthStyleConstants)
    If nValue < 0 Or nValue > 3 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If mTabWidthStyle <> nValue Then
        mTabWidthStyle = nValue
        PropertyChanged "TabWidthStyle"
        Draw
    End If
End Property


Public Property Get TabAppearance() As vbExTabAppearanceConstants
Attribute TabAppearance.VB_Description = "Returns/sets a value that determines the appearance of the tabs."
Attribute TabAppearance.VB_MemberFlags = "400"
    TabAppearance = mTabAppearance
End Property

Public Property Let TabAppearance(ByVal nValue As vbExTabAppearanceConstants)
    If nValue < 0 Or nValue > 4 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If mTabAppearance <> nValue Then
        mTabAppearance = nValue
        PropertyChanged "TabAppearance"
        ResetCachedThemeImages
        Draw
    End If
End Property



' Returns/sets the type of mouse pointer displayed when over the control.
Public Property Get MousePointer() As vbExMousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over the control."
    MousePointer = mMousePointer
End Property

Public Property Let MousePointer(ByVal nValue As vbExMousePointerConstants)
    Select Case nValue
        Case Is < 0, 16 To 98, Is > 99
            RaiseError 380, TypeName(Me) ' invalid property value
            Exit Property
    End Select
    If nValue <> mMousePointer Then
        mMousePointer = nValue
        UserControl.MousePointer = mMousePointer
        PropertyChanged "MousePointer"
    End If
End Property


' Returns/sets the icon used as the mouse pointer when the MousePointer property is set to 99 (custom).
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Returns/sets the icon used as the mouse pointer when the MousePointer property is set to 99 (custom)."
    Set MouseIcon = mMouseIcon
End Property

Public Property Let MouseIcon(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If Not nValue Is mMouseIcon Then
        Set mMouseIcon = nValue
        PropertyChanged "MouseIcon"
    End If
End Property

Public Property Set MouseIcon(ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If Not nValue Is mMouseIcon Then
        Set mMouseIcon = nValue
        Set UserControl.MouseIcon = mMouseIcon
        PropertyChanged "MouseIcon"
    End If
End Property


' Returns/Sets whether this control can act as an OLE drop target.
Public Property Get OLEDropMode() As vbExOLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/sets how a target component handles drop operations."
    OLEDropMode = mOLEDropMode
End Property

Public Property Let OLEDropMode(ByVal nValue As vbExOLEDropConstants)
    If nValue < 0 Or nValue > 1 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mOLEDropMode Then
        mOLEDropMode = nValue
        UserControl.OLEDropMode = mOLEDropMode
        PropertyChanged "OLEDropMode"
    End If
End Property


' Returns the picture displayed on the specified tab.
Public Property Get TabPicture(ByVal Index As Integer) As Picture
Attribute TabPicture.VB_Description = "Returns/sets the picture to be displayed on the specified tab."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    Set TabPicture = mTabData(Index).Picture
End Property

Public Property Let TabPicture(ByVal Index As Integer, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    Set TabPicture(Index) = nValue
End Property

Public Property Set TabPicture(ByVal Index As Integer, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If Not nValue Is mTabData(Index).Picture Then
        Set mTabData(Index).Picture = nValue
        mTabData(Index).PicToUseSet = False
        mTabData(Index).PicDisabledSet = False
        PropertyChanged "TabPicture"
        SetAutoTabHeight
        DrawDelayed
    End If
End Property


Public Property Get TabPic16(ByVal Index) As Picture
Attribute TabPic16.VB_Description = "Specifies a bitmap to display on the specified tab at 96 DPI, when the application is DPI aware."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    Set TabPic16 = mTabData(Index).Pic16
End Property

Public Property Let TabPic16(ByVal Index, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    Set TabPic16(Index) = nValue
End Property

Public Property Set TabPic16(ByVal Index, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If Not nValue Is mTabData(Index).Pic16 Then
        Set mTabData(Index).Pic16 = nValue
        mTabData(Index).PicToUseSet = False
        mTabData(Index).PicDisabledSet = False
        PropertyChanged "TabPic16"
        SetAutoTabHeight
        DrawDelayed
    End If
End Property


Public Property Get TabPic20(ByVal Index) As Picture
Attribute TabPic20.VB_Description = "Specifies a bitmap to display on the specified tab at 120 DPI, when the application is DPI aware."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    Set TabPic20 = mTabData(Index).Pic20
End Property

Public Property Let TabPic20(ByVal Index, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    Set TabPic20(Index) = nValue
End Property

Public Property Set TabPic20(ByVal Index, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If Not nValue Is mTabData(Index).Pic20 Then
        Set mTabData(Index).Pic20 = nValue
        mTabData(Index).PicToUseSet = False
        mTabData(Index).PicDisabledSet = False
        PropertyChanged "TabPic20"
        SetAutoTabHeight
        DrawDelayed
    End If
End Property


Public Property Get TabPic24(ByVal Index) As Picture
Attribute TabPic24.VB_Description = "Specifies a bitmap to display on the specified tab at 144 DPI, when the application is DPI aware."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    Set TabPic24 = mTabData(Index).Pic24
End Property

Public Property Let TabPic24(ByVal Index, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    Set TabPic20(Index) = nValue
End Property

Public Property Set TabPic24(ByVal Index, ByVal nValue As Picture)
    If Not nValue Is Nothing Then If nValue.Handle = 0 Then Set nValue = Nothing
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If Not nValue Is mTabData(Index).Pic24 Then
        Set mTabData(Index).Pic24 = nValue
        mTabData(Index).PicToUseSet = False
        mTabData(Index).PicDisabledSet = False
        PropertyChanged "TabPic24"
        SetAutoTabHeight
        DrawDelayed
    End If
End Property


' Determines if the specified tab is visible.
Public Property Get TabVisible(ByVal Index As Integer) As Boolean
Attribute TabVisible.VB_Description = "Determines if the specified tab is visible."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    TabVisible = mTabData(Index).Visible
End Property

Public Property Let TabVisible(ByVal Index As Integer, ByVal nValue As Boolean)
    Dim c As Long
    
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mTabData(Index).Visible Then
        If mTabSel = Index Then
            c = mTabSel - 1
            Do Until c < 0
                If mTabData(c).Visible And mTabData(c).Enabled Then
                    Exit Do
                End If
                c = c - 1
            Loop
            If c = -1 Then
                c = mTabSel + 1
                Do Until c = mTabs
                    If mTabData(c).Visible And mTabData(c).Enabled Then
                        Exit Do
                    End If
                    c = c + 1
                Loop
            End If
            If (c < 0) Or (c > (mTabs - 1)) Then
                c = mTabSel - 1
                Do Until c < 0
                    If mTabData(c).Visible Then
                        Exit Do
                    End If
                    c = c - 1
                Loop
                If c = -1 Then
                    c = mTabSel + 1
                    Do Until c = mTabs
                        If mTabData(c).Visible Then
                            Exit Do
                        End If
                        c = c + 1
                    Loop
                End If
            End If
            If (c > -1) And (c < mTabs) Then
                TabSel = c
                If mTabSel = c Then ' the change could had been canceled through the BeforeClick event, in that case TabSel woudn't change
                    mTabData(Index).Visible = nValue
                    mTabData(Index).Selected = False
                End If
            Else
                mTabSel = -1
                mTabData(Index).Visible = nValue
                mTabData(Index).Selected = False
                HideAllContainedControls
            End If
        Else
            mTabData(Index).Visible = nValue
            If (mTabSel < 0) Or (mTabSel > (mTabs - 1)) Then
                TabSel = Index
                mTabData(Index).Selected = True
            End If
        End If
        mAccessKeysSet = False
        PropertyChanged "TabVisible"
        mTabBodyReset = True
        DrawDelayed
    End If
End Property


' Determines if the specified tab is enabled.
Public Property Get TabEnabled(ByVal Index As Integer) As Boolean
Attribute TabEnabled.VB_Description = "Determines if the specified tab is enabled."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    TabEnabled = mTabData(Index).Enabled
End Property

Public Property Let TabEnabled(ByVal Index As Integer, ByVal nValue As Boolean)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mTabData(Index).Enabled Then
        mTabData(Index).Enabled = nValue
        PropertyChanged "TabEnabled"
        mAccessKeysSet = False
        DrawDelayed
    End If
End Property


' Returns the text displayed on the specified tab.
Public Property Get TabCaption(ByVal Index As Integer) As String
Attribute TabCaption.VB_Description = "Returns the text displayed on the specified tab."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    TabCaption = mTabData(Index).Caption
End Property

Public Property Let TabCaption(ByVal Index As Integer, ByVal nValue As String)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mTabData(Index).Caption Then
        mTabData(Index).Caption = nValue
        PropertyChanged "TabCaption"
        mAccessKeysSet = False
        DrawDelayed
    End If
End Property


Public Property Get TabToolTipText(ByVal Index) As String
Attribute TabToolTipText.VB_Description = "Returns/sets the text that will be shown as tooltip text when the mouse pointer is over the specified tab."
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    TabToolTipText = mTabData(Index).ToolTipText
End Property

Public Property Let TabToolTipText(ByVal Index, ByVal nValue As String)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mTabData(Index).ToolTipText Then
        If mThereAreTabsToolTipTexts And mAmbientUserMode Then RestoreExtenderTTT
        mTabData(Index).ToolTipText = nValue
        CheckIfThereAreTabsToolTipTexts
        If mTabUnderMouse > -1 Then
            If mTabData(mTabUnderMouse).ToolTipText <> "" Then
                ShowTabTTT mTabUnderMouse
            End If
        End If
        PropertyChanged "TabToolTipText"
    End If
End Property


Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets a color in the tabs pictures to be a mask (that is, transparent)."
    MaskColor = mMaskColor
End Property

Public Property Let MaskColor(ByVal nValue As OLE_COLOR)
    If nValue <> mMaskColor Then
        mMaskColor = nValue
        PropertyChanged "MaskColor"
        Draw
    End If
End Property


Public Property Get TabSelExtraHeight() As Single
Attribute TabSelExtraHeight.VB_Description = "Returns/sets a value that determines if the active tab will be higher than the others."
Attribute TabSelExtraHeight.VB_MemberFlags = "400"
    TabSelExtraHeight = FixRoundingError(ToContainerSizeY(mTabSelExtraHeight, vbHimetric))
End Property

Public Property Let TabSelExtraHeight(ByVal nValue As Single)
    Dim iValue As Single
    
    iValue = FromContainerSizeY(nValue, vbHimetric)
    If iValue < 0 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If iValue > mTabHeight Then
        iValue = mTabHeight 'limit
    End If
    If Round(iValue * 10000) <> Round(mTabSelExtraHeight * 10000) Then
        mTabSelExtraHeight = iValue
        PropertyChanged "TabSelExtraHeight"
        Draw
    End If
End Property


Public Property Get TabSelHighlight() As Boolean
Attribute TabSelHighlight.VB_Description = "Returns/sets a value that determines if the selected tab will be highlighted."
Attribute TabSelHighlight.VB_MemberFlags = "400"
    TabSelHighlight = mTabSelHighlight
End Property

Public Property Let TabSelHighlight(ByVal nValue As Boolean)
    If nValue <> mTabSelHighlight Then
        mTabSelHighlight = nValue
        PropertyChanged "TabSelHighlight"
        Draw
    End If
End Property


Public Property Get TabHoverHighlight() As vbExTabHoverHighlightConstants
Attribute TabHoverHighlight.VB_Description = "Returns/sets a value that determines if the tabs will appear highlighted when the mouse is over them."
Attribute TabHoverHighlight.VB_MemberFlags = "400"
    TabHoverHighlight = mTabHoverHighlight
End Property

Public Property Let TabHoverHighlight(ByVal nValue As vbExTabHoverHighlightConstants)
    If nValue <> mTabHoverHighlight Then
        mTabHoverHighlight = nValue
        If (mTabHoverHighlight < ssTHHNo) Or (mTabHoverHighlight > ssTHHEffect) Then
            mTabHoverHighlight = ssTHHInstant
        End If
        PropertyChanged "TabHoverHighlight"
        Draw
    End If
End Property


Public Property Get VisualStyles() As Boolean
Attribute VisualStyles.VB_Description = "Returns/sets a value that determines whether the appearance of the control will use Windows visual styles-"
    VisualStyles = mVisualStyles
End Property

Public Property Let VisualStyles(ByVal nValue As Boolean)
    Dim iWv As Boolean
    
    If nValue <> mVisualStyles Then
        If nValue And (IsAppThemeEnabled Or mForceVisualStyles) Then
            mTabBackColor_SavedWhileVisualStyles = mTabBackColor
            mTabSelBackColor_SavedWhileVisualStyles = mTabSelBackColor
            TabBackColor = vbButtonFace
            TabSelBackColor = vbButtonFace
            mTabBackColorSavedWhileVisualStyles = True
        Else
            If mTabBackColorSavedWhileVisualStyles Then
                mTabBackColorSavedWhileVisualStyles = False
                TabBackColor = mTabBackColor_SavedWhileVisualStyles
                TabSelBackColor = mTabSelBackColor_SavedWhileVisualStyles
            End If
        End If
        mVisualStyles = nValue
        PropertyChanged "VisualStyles"
        mSubclassControlsPaintingPending = True
        mRepaintSubclassedControls = True
        iWv = IsWindowVisible(mUserControlHwnd) <> 0
        If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
        Draw
        If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
        If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
    End If
End Property


Public Property Get TabBackColor() As OLE_COLOR
Attribute TabBackColor.VB_Description = "Returns/sets the background color of the tabs."
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        TabBackColor = mHandleHighContrastTheme_OrigTabBackColor
    Else
        If mTabBackColorSavedWhileVisualStyles Then
            TabBackColor = mTabBackColor_SavedWhileVisualStyles
        Else
            TabBackColor = mTabBackColor
        End If
    End If
End Property

Public Property Let TabBackColor(ByVal nValue As OLE_COLOR)
    Dim iWv As Boolean
    Dim iPrev As Long
    
    If nValue <> mTabBackColor Then
        If mTabBackColorSavedWhileVisualStyles Then
            mTabBackColor_SavedWhileVisualStyles = nValue
        Else
            iPrev = mTabBackColor
            If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
                mHandleHighContrastTheme_OrigTabBackColor = nValue
                If (mTabSelBackColor = iPrev) And (mTabSelBackColor <> nValue) Then
                    TabSelBackColor = nValue
                End If
            Else
                mTabBackColor = nValue
                PropertyChanged "TabBackColor"
                SetColors
                iWv = IsWindowVisible(mUserControlHwnd) <> 0
                If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
                If (mTabSelBackColor = iPrev) And (mTabSelBackColor <> nValue) And (mBackStyle = ssOpaque) Then
                    TabSelBackColor = nValue
                Else
                    Draw
                End If
                If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
                If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
            End If
        End If
    End If
End Property


Public Property Get TabSelBackColor() As OLE_COLOR
Attribute TabSelBackColor.VB_Description = "Returns /sets the color of the active tab including the tab body."
    If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
        TabSelBackColor = mHandleHighContrastTheme_OrigTabSelBackColor
    Else
        If mTabBackColorSavedWhileVisualStyles Then
            TabSelBackColor = mTabSelBackColor_SavedWhileVisualStyles
        Else
            TabSelBackColor = mTabSelBackColor
        End If
    End If
End Property

Public Property Let TabSelBackColor(ByVal nValue As OLE_COLOR)
    Dim iPrev As Long
    Dim iWv As Boolean
    
    If nValue <> mTabSelBackColor Then
        If mAmbientUserMode And mHandleHighContrastTheme And mHighContrastThemeOn Then
            mHandleHighContrastTheme_OrigTabSelBackColor = nValue
        Else
            If mTabBackColorSavedWhileVisualStyles Then
                mTabSelBackColor_SavedWhileVisualStyles = nValue
            Else
                If Enabled Or Not mShowDisabledState Then
                    iPrev = mTabSelBackColor
                Else
                    iPrev = mTabSelBackColorDisabled
                End If
                mTabSelBackColor = nValue
                PropertyChanged "TabSelBackColor"
                SetColors
                iWv = IsWindowVisible(mUserControlHwnd) <> 0
                If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
                If mChangeControlsBackColor Then
                    SetControlsBackColor IIf((Not Enabled) And mShowDisabledState, mTabSelBackColorDisabled, mTabSelBackColor), iPrev
                End If
                mSubclassControlsPaintingPending = True
                mRepaintSubclassedControls = True
                mTabBodyReset = True
                SubclassControlsPainting
                Draw
                If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
                If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
            End If
        End If
    End If
End Property


Public Property Get ShowDisabledState() As Boolean
Attribute ShowDisabledState.VB_Description = "Returns/sets a value that determines if the tabs color will be darkened when the control is disabled."
Attribute ShowDisabledState.VB_MemberFlags = "400"
    ShowDisabledState = mShowDisabledState
End Property

Public Property Let ShowDisabledState(ByVal nValue As Boolean)
    If nValue <> mShowDisabledState Then
        mShowDisabledState = nValue
        PropertyChanged "ShowDisabledState"
        mTabBodyReset = True
        Draw
        If mChangeControlsBackColor Then
            If mEnabled Or Not mShowDisabledState Then
                SetControlsBackColor mTabSelBackColor, mTabSelBackColorDisabled
            Else
                SetControlsBackColor mTabSelBackColorDisabled, mTabSelBackColor
            End If
        End If
    End If
End Property


Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Returns/sets a value that determines if the drawing of the control is enabled."
Attribute Redraw.VB_MemberFlags = "400"
    Redraw = mRedraw
End Property

Public Property Let Redraw(ByVal nValue As Boolean)
    If nValue <> mRedraw Then
        mRedraw = nValue
        If mRedraw Then
            If mNeedToDraw Then
                Draw
            End If
        End If
    End If
End Property


Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets a value that determines whether the color assigned in the MaskColor property is used as a mask. (That is, used to create transparent regions.)"
    UseMaskColor = mUseMaskColor
End Property

Public Property Let UseMaskColor(ByVal nValue As Boolean)
    If nValue <> mUseMaskColor Then
        mUseMaskColor = nValue
        PropertyChanged "UseMaskColor"
        Draw
    End If
End Property

Public Property Get TabSelFontBold() As vbExAutoYesNoConstants
Attribute TabSelFontBold.VB_Description = "Returns/sets a value that determines if the font of the caption in currently selected tab will be bold."
Attribute TabSelFontBold.VB_MemberFlags = "400"
    TabSelFontBold = mTabSelFontBold
End Property

Public Property Let TabSelFontBold(ByVal nValue As vbExAutoYesNoConstants)
    Dim iValue As vbExAutoYesNoConstants
    
    iValue = nValue
    If (iValue <> ssNo) And (iValue <> ssYNAuto) Then
        iValue = ssYes
    End If
    If iValue <> mTabSelFontBold Then
        mTabSelFontBold = iValue
        PropertyChanged "TabSelFontBold"
        Draw
    End If
End Property


Public Property Get ShowRowsInPerspective() As vbExAutoYesNoConstants
Attribute ShowRowsInPerspective.VB_Description = "Returns/sets a value that determines when the control has more that one row of tabs, if they will be drawn changing the horizontal position on each row."
Attribute ShowRowsInPerspective.VB_MemberFlags = "400"
    ShowRowsInPerspective = mShowRowsInPerspective
End Property

Public Property Let ShowRowsInPerspective(ByVal nValue As vbExAutoYesNoConstants)
    Dim iValue As vbExAutoYesNoConstants
    
    iValue = nValue
    If (iValue <> ssNo) And (iValue <> ssYNAuto) Then
        iValue = ssYes
    End If
    If iValue <> mShowRowsInPerspective Then
        mShowRowsInPerspective = iValue
        PropertyChanged "ShowRowsInPerspective"
        ResetCachedThemeImages
        Draw
    End If
End Property


Public Property Get TabSeparation() As Integer
Attribute TabSeparation.VB_Description = "Returns/sets the number of pixels of separation between tabs."
Attribute TabSeparation.VB_MemberFlags = "400"
    TabSeparation = mTabSeparation
End Property

Public Property Let TabSeparation(ByVal nValue As Integer)
    If nValue < 0 Or nValue > 20 Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mTabSeparation Then
        mTabSeparation = nValue
        PropertyChanged "TabSeparation"
        ResetCachedThemeImages
        Draw
    End If
End Property

Public Property Get ChangeControlsBackColor() As Boolean
Attribute ChangeControlsBackColor.VB_Description = "Returns/sets a value that determines if the background color of the contained controls will be changed according to TabBackColor."
    ChangeControlsBackColor = mChangeControlsBackColor
End Property

Public Property Let ChangeControlsBackColor(ByVal nValue As Boolean)
    Dim iWv As Boolean
    
    If nValue <> mChangeControlsBackColor Then
        mChangeControlsBackColor = nValue
        PropertyChanged "ChangeControlsBackColor"
        iWv = IsWindowVisible(mUserControlHwnd) <> 0
        If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
        If Not mChangeControlsBackColor Then
            SetControlsBackColor vbButtonFace, IIf(mEnabled Or Not mShowDisabledState, mTabSelBackColor, mTabSelBackColorDisabled)
        Else
            SetControlsBackColor IIf(mEnabled Or Not mShowDisabledState, mTabSelBackColor, mTabSelBackColorDisabled)
        End If
        mSubclassControlsPaintingPending = True
        mRepaintSubclassedControls = True
        SubclassControlsPainting
        Draw
        If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
        If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
    End If
End Property


Public Property Get AutoRelocateControls() As vbExAutoRelocateControlsConstants
Attribute AutoRelocateControls.VB_Description = "Returns/sets a value that determines if the contained controls will be automatically relocated when the tab body changes."
    AutoRelocateControls = mAutoRelocateControls
End Property

Public Property Let AutoRelocateControls(ByVal nValue As vbExAutoRelocateControlsConstants)
    If (nValue < 0) Or (nValue > 2) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    If nValue <> mAutoRelocateControls Then
        mAutoRelocateControls = nValue
        PropertyChanged "AutoRelocateControls"
    End If
End Property


Public Property Get SoftEdges() As Boolean
Attribute SoftEdges.VB_Description = "Returns/sets a value that determines if the edges will be displayed with less contrast."
    SoftEdges = mSoftEdges
End Property

Public Property Let SoftEdges(ByVal nValue As Boolean)
    If nValue <> mSoftEdges Then
        mSoftEdges = nValue
        PropertyChanged "SoftEdges"
        SetColors
        PostDrawMessage
    End If
End Property


Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Returns/Sets the text display direction and control visual appearance on a bidirectional system."
    RightToLeft = mRightToLeft
End Property

Public Property Let RightToLeft(ByVal nValue As Boolean)
    If nValue <> mRightToLeft Then
        mRightToLeft = nValue
        If mRightToLeft Then
            SetLayout GetDC(picDraw.hWnd), LAYOUT_RTL Or LAYOUT_BITMAPORIENTATIONPRESERVED
        Else
            SetLayout GetDC(picDraw.hWnd), 0
        End If
        PropertyChanged "RightToLeft"
        PostDrawMessage
    End If
End Property


Public Property Get BackStyle() As vbExBackStyleConstants
Attribute BackStyle.VB_Description = "Returns/sets the background style, opaque or transparent."
    BackStyle = mBackStyle
End Property

Public Property Let BackStyle(ByVal nValue As vbExBackStyleConstants)
    If nValue <> mBackStyle Then
        mBackStyle = nValue
        PropertyChanged "BackStyle"
        'ResetCachedThemeImages
        Draw
    End If
End Property


Public Property Get AutoTabHeight() As Boolean
Attribute AutoTabHeight.VB_Description = "Returns/sets a value that determines if the tab height is set automatically according to the font (and pictures)."
    AutoTabHeight = mAutoTabHeight
End Property

Public Property Let AutoTabHeight(ByVal nValue As Boolean)
    If nValue <> mAutoTabHeight Then
        mAutoTabHeight = nValue
        PropertyChanged "AutoTabHeight"
        SetAutoTabHeight
        Draw
    End If
End Property


Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As Long
    Select Case iMsg
        Case WM_PAINT, WM_PRINTCLIENT, WM_MOUSELEAVE
            IBSSubclass_MsgResponse = emrConsume
        Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEACTIVATE, WM_SETFOCUS, WM_LBUTTONDBLCLK, WM_MOVE, WM_WINDOWPOSCHANGING
            IBSSubclass_MsgResponse = emrPreprocess
        Case Else
            IBSSubclass_MsgResponse = emrPostProcess
    End Select
End Function

Private Sub IBSSubclass_UnsubclassIt()
    If mSubclassed Then
        ' The IDE protection was fired
        DoTerminate
           
        'If (Not mAmbientUserMode) Then
            ' The following emulates the zombie state (UserControl hatched/disabled), in case it didn't actually happened by VB.
            ' Because the control anyway will be unclickable on the IDE any more without the subclassing.
            ' The developer needs to close the form and open it again to restore the functionality.
            UserControl.FillStyle = 5
'            UserControl.DrawWidth = 30
'            UserControl.FillColor = vbRed
            UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), , B
            UserControl.FillStyle = 1
            UserControl.Enabled = False
        'End If
    End If
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    Dim iTab As Long
    
    Select Case iMsg
        Case WM_WINDOWPOSCHANGING ' invisible controls, to prevent being moved to the visible space if they are moved by code. Unfortunately the same can't be done to Labels and other windowless controls. But at least the protection acts on windowed controls.
            Dim iwp As WINDOWPOS
            
            CopyMemory iwp, ByVal lParam, Len(iwp)
            If iwp.X > -mLeftThresholdHided \ Screen.TwipsPerPixelX Then
                iwp.X = iwp.X - mLeftShiftToHide \ Screen.TwipsPerPixelX
                CopyMemory ByVal lParam, iwp, Len(iwp)
            End If
            
        Case WM_NCACTIVATE ' need to update the focus rect
            mFormIsActive = (wParam <> 0)
            If mHasFocus Then
                PostDrawMessage
            End If
        Case WM_PRINTCLIENT, WM_MOUSELEAVE ' fixes frames paint bug in XP
            IBSSubclass_WindowProc = DefWindowProc(hWnd, iMsg, wParam, lParam)
        Case WM_SYSCOLORCHANGE, WM_THEMECHANGED ' they are form's messages
            SetButtonFaceColor
            SetColors
            mThemeExtraDataAlreadySet = False
            SetThemeExtraData
            ResetCachedThemeImages
            If mHandleHighContrastTheme Then CheckHighContrastTheme
            Draw
        Case WM_SETFOCUS
            If mNoActivate Then
                bConsume = True
                IBSSubclass_WindowProc = 0
                SetFocusAPI wParam
                mNoActivate = False
            End If
        Case WM_DRAW
            Draw
        Case WM_INIT
            If Not mTabStopsInitialized Then
                StoreControlsTabStop True
                mTabStopsInitialized = True
            End If
        Case WM_MOUSEACTIVATE ' UserControl message, only at run time (Ambient.UserMode),, to avoid taking the focus when the tab control is clicked in a non-clickable part (outside a tab).
            If mTabUnderMouse = -1 Then
                Dim iPt2 As POINTAPI
                Dim iHwnd As Long
                
                GetCursorPos iPt2
                iHwnd = WindowFromPoint(iPt2.X, iPt2.Y)
                If iHwnd = mUserControlHwnd Then
                    mNoActivate = True
                End If
            End If
        Case WM_LBUTTONDBLCLK
            If Not MouseIsOverAContainedControl Then
                iTab = mTabSel
                Call ProcessMouseMove(vbLeftButton, 0, (lParam And &HFFFF&) * Screen_TwipsPerPixelX, (lParam \ &H10000 And &HFFFF&) * Screen_TwipsPerPixelX)
                Call UserControl_MouseDown(vbLeftButton, 0, (lParam And &HFFFF&) * Screen_TwipsPerPixelX, (lParam \ &H10000 And &HFFFF&) * Screen_TwipsPerPixelX)
                If mTabSel <> iTab Then
                    bConsume = True
                    IBSSubclass_WindowProc = 0
                    tmrCancelDoubleClick.Enabled = True
                End If
            End If
            If tmrCancelDoubleClick.Enabled Then
                bConsume = True
                mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, GetMessageExtraInfo()
                mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, GetMessageExtraInfo()
            End If
        Case WM_LBUTTONDOWN ' UserControl message, only in design mode (Not Ambient.UserMode), to provide change of selected tab by clicking at design time
            If Not MouseIsOverAContainedControl Then
                iTab = mTabSel
                Call ProcessMouseMove(vbLeftButton, 0, (lParam And &HFFFF&) * Screen_TwipsPerPixelX, (lParam \ &H10000 And &HFFFF&) * Screen_TwipsPerPixelX)
                Call UserControl_MouseDown(vbLeftButton, 0, (lParam And &HFFFF&) * Screen_TwipsPerPixelX, (lParam \ &H10000 And &HFFFF&) * Screen_TwipsPerPixelX)
                If mTabSel <> iTab Then
                    bConsume = True
                    IBSSubclass_WindowProc = 0
                    mBtnDown = True
                    'tmrCancelDoubleClick.Enabled = True
                End If
            End If
            If mChangeControlsBackColor And ((mTabBackColor <> vbButtonFace) Or mControlIsThemed) Then
                mLastContainedControlsCount = UserControl.ContainedControls.Count
                tmrCheckContainedControlsAdditionDesignTime.Enabled = True
            End If
        Case WM_LBUTTONUP ' UserControl message, only in design mode (Not Ambient.UserMode). To avoid the IDE to start dragging the control on mouse down when the developer clicks to change the selected tab
            If mBtnDown Then
                mBtnDown = False
                SendMessage hWnd, WM_LBUTTONDOWN, wParam, lParam
            End If
        Case WM_MOVE
            RedrawWindow hWnd, ByVal 0, 0, RDW_INVALIDATE Or RDW_ALLCHILDREN
        Case WM_PAINT ' contained controls paint messages, when the control is themed and ChangeControlsBackColor = True (only at run time, Ambient.UserMode)
            
            Dim iUpdateRect As RECT
            Dim iControlRect As RECT
            Dim iDestDC As Long
            Dim iWidth As Long
            Dim iHeight As Long
            Dim iTempDC As Long
            Dim iTempBmp As Long
            Dim iPs As PAINTSTRUCT
            Dim iBKColor As Long
            Dim iPt As POINTAPI
            Dim iBrush As Long
            Dim iTop As Long
            Dim iLeft As Long
            Dim iColor As Long
            Dim iFillRect As RECT
            
            If GetUpdateRect(hWnd, iUpdateRect, 0&) <> 0& Then
                Call BeginPaint(hWnd, iPs)
                
                iDestDC = iPs.hDC
                GetWindowRect hWnd, iControlRect
                
                iPt.X = iControlRect.Left + iPs.rcPaint.Left
                iPt.Y = iControlRect.Top + iPs.rcPaint.Top
                ScreenToClient hWnd, iPt
                iControlRect.Left = iControlRect.Left - iPt.X
                iControlRect.Top = iControlRect.Top - iPt.Y
                
                iTempDC = CreateCompatibleDC(iDestDC)
                iTempBmp = CreateCompatibleBitmap(iDestDC, iControlRect.Right - iControlRect.Left, iControlRect.Bottom - iControlRect.Top)
                DeleteObject SelectObject(iTempDC, iTempBmp)
                
                CallOldWindowProc hWnd, iMsg, iTempDC, lParam
                
                iWidth = iControlRect.Right - iControlRect.Left
                iHeight = iControlRect.Bottom - iControlRect.Top
                
                iPt.X = iControlRect.Left + iPs.rcPaint.Left
                iPt.Y = iControlRect.Top + iPs.rcPaint.Top
                ScreenToClient mUserControlHwnd, iPt
                
                
                If mChangeControlsBackColor Then
                    If mShowDisabledState And (Not mEnabled) Then
                        iColor = mTabSelBackColorDisabled
                    Else
                        iColor = mTabSelBackColor
                    End If
                Else
                    iColor = vbButtonFace
                End If
                TranslateColor iColor, 0&, iBKColor
                
                ' set the part of the update rect of the control that must be painted with the backgroung bitmap because is inside the tab body
                If iPt.Y < mTabBodyRect.Top Then
                    iHeight = iHeight - (mTabBodyRect.Top - 1 - iPt.Y)
                    iTop = (mTabBodyRect.Top - 1 - iPt.Y)
                    iPt.Y = mTabBodyRect.Top - 1
                    If (mTabBodyRect.Top + iHeight - 2) > mTabBodyRect.Bottom Then
                        iHeight = mTabBodyRect.Bottom - mTabBodyRect.Top + 2
                    End If
                ElseIf iPt.Y + iHeight > mTabBodyRect.Bottom Then
                    iHeight = mTabBodyRect.Bottom - iPt.Y
                    iTop = 0
                End If
                
                If iPt.X < mTabBodyRect.Left Then
                    iWidth = iWidth - (mTabBodyRect.Left - iPt.X)
                    iLeft = (mTabBodyRect.Left - 1 - iPt.X)
                    iPt.X = mTabBodyRect.Left - 1
                    If (mTabBodyRect.Left + iWidth - 2) > mTabBodyRect.Right Then
                        iWidth = mTabBodyRect.Right - mTabBodyRect.Left + 2
                    End If
                ElseIf iPt.X + iWidth > mTabBodyRect.Right Then
                    iWidth = mTabBodyRect.Right - iPt.X
                    iLeft = 0
                End If
                
                ' iLeft and iTop: from where to paint into the control in coordinates of the control
                ' iWidth and iHeight: the size of the image to be painted into the control
                ' iPt.X and iPt.Y: the position in the UserControl from where to take the image to be painted, in coordinales of the UserControl
                
                'the rest of the update rect that was not painted must be filled with the tab backcolor (if there are parts that are outside the tab body)
                
                If iTop > iPs.rcPaint.Top Then  ' there is a space over the painted region that must be filled
                    iFillRect = iPs.rcPaint
                    iFillRect.Bottom = iTop + 1
                    If iFillRect.Bottom > iFillRect.Top Then
                        iBrush = CreateSolidBrush(iBKColor)
                        FillRect iDestDC, iFillRect, iBrush
                        DeleteObject iBrush
                    End If
                End If
                If iLeft > iPs.rcPaint.Left Then   ' there is a space over the painted region that must be filled
                    iFillRect = iPs.rcPaint
                    iFillRect.Right = iLeft + 1
                    If iFillRect.Right > iFillRect.Left Then
                        iBrush = CreateSolidBrush(iBKColor)
                        FillRect iDestDC, iFillRect, iBrush
                        DeleteObject iBrush
                    End If
                End If
                If (iTop + iHeight) < iPs.rcPaint.Bottom Then
                    iFillRect = iPs.rcPaint
                    iFillRect.Top = (iTop + iHeight)
                    If iFillRect.Bottom > iFillRect.Top Then
                        iBrush = CreateSolidBrush(iBKColor)
                        FillRect iDestDC, iFillRect, iBrush
                        DeleteObject iBrush
                    End If
                End If
                If (iLeft + iWidth) < iPs.rcPaint.Right Then
                    iFillRect = iPs.rcPaint
                    iFillRect.Left = (iLeft + iWidth)
                    If iFillRect.Right > iFillRect.Left Then
                        iBrush = CreateSolidBrush(iBKColor)
                        FillRect iDestDC, iFillRect, iBrush
                        DeleteObject iBrush
                    End If
                End If

                If (iHeight > 0) And (iWidth > 0) Then
                    BitBlt iDestDC, iLeft, iTop, iWidth, iHeight, UserControl.hDC, iPt.X, iPt.Y, vbSrcCopy
                End If
                TransparentBlt iDestDC, iPs.rcPaint.Left, iPs.rcPaint.Top, iPs.rcPaint.Right - iPs.rcPaint.Left, iPs.rcPaint.Bottom - iPs.rcPaint.Top, iTempDC, iPs.rcPaint.Left, iPs.rcPaint.Top, iPs.rcPaint.Right - iPs.rcPaint.Left, iPs.rcPaint.Bottom - iPs.rcPaint.Top, iBKColor
                DeleteDC iTempDC
                DeleteObject iTempBmp
                Call EndPaint(hWnd, iPs)
                IBSSubclass_WindowProc = 0
            Else
                IBSSubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            End If
        Case WM_GETDPISCALEDSIZE
            Dim iPrev As Long
            
            iPrev = mLeftShiftToHide
            SetLeftShiftToHide Int(1440 / wParam)
            If mLeftShiftToHide <> iPrev Then
                mPendingLeftShift = iPrev - mLeftShiftToHide
                DoPendingLeftShift
            End If
    End Select
End Function

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    If Not mDrawing Then
        SetAutoTabHeight
        Draw
    End If
End Sub

Private Sub mForm_Load()
    UserControl_Show
End Sub

Private Sub tmrCancelDoubleClick_Timer()
    tmrCancelDoubleClick.Enabled = False
End Sub

Private Sub tmrCheckContainedControlsAdditionDesignTime_Timer()
    If IsMouseButtonPressed(vxMBLeft) Then Exit Sub
    If mBackStyle = ssOpaque Then tmrCheckContainedControlsAdditionDesignTime.Enabled = False
    
    If UserControl.ContainedControls.Count <> mLastContainedControlsCount Then
        mLastContainedControlsCount = UserControl.ContainedControls.Count
        SetControlsBackColor mTabSelBackColor
        If mControlIsThemed Or (mBackStyle = ssTransparent) Then
            mSubclassControlsPaintingPending = True
            RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
            Draw
        End If
    ElseIf (Not Ambient.UserMode) And (mBackStyle = ssTransparent) Then
        Dim iStr As String
        
        iStr = GetContainedControlsPositionsStr
        If iStr <> mLastContainedControlsPositionsStr Then
            mLastContainedControlsPositionsStr = iStr
            mSubclassControlsPaintingPending = True
            RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
            Draw
        End If
    End If
End Sub

Private Function GetContainedControlsPositionsStr() As String
    Dim iCtl As Control
    Dim iLeft As Long
    Dim iWidth As Long
    
    On Error Resume Next
    For Each iCtl In UserControl.ContainedControls
        iLeft = -mLeftShiftToHide
        iLeft = iCtl.Left
        If iLeft > -mLeftShiftToHide Then
            iWidth = -1
            iWidth = iCtl.Width
            If iWidth <> -1 Then
                GetContainedControlsPositionsStr = GetContainedControlsPositionsStr & CStr(iLeft) & "," & CStr(iCtl.Top) & "," & CStr(iWidth) & "," & CStr(iCtl.Height) & "|"
            End If
        End If
    Next
    'On Error GoTo 0
    
End Function
    

Private Function IsMouseButtonPressed(nButton As vbExMouseButtonsConstants) As Boolean
    Dim iButton As Long
    
    iButton = nButton
    If GetSystemMetrics(SM_SWAPBUTTON) <> 0 Then
        If nButton = vxMBLeft Then
            iButton = VK_RBUTTON
        ElseIf nButton = vxMBRight Then
            iButton = VK_LBUTTON
        End If
    End If
    IsMouseButtonPressed = GetAsyncKeyState(iButton) <> 0
End Function

Private Sub tmrCheckDuplicationByIDEPaste_Timer()
    If (Not Ambient.UserMode) Then
        If Not IsMsgBoxShown Then
            tmrCheckDuplicationByIDEPaste.Enabled = False
            CheckContainedControlsConsistency
        End If
    Else
        tmrCheckDuplicationByIDEPaste.Enabled = False
    End If
End Sub

Private Sub tmrDraw_Timer()
    Draw
End Sub

Private Sub tmrSubclassControls_Timer()
    tmrSubclassControls.Enabled = False
    SubclassControlsPainting
End Sub

Private Sub tmrTabHoverEffect_Timer()
    mTabHoverEffect_Step = mTabHoverEffect_Step + 1
    mGlowColor = mHoverEffectColors(mTabHoverEffect_Step)
    Draw
    If mTabHoverEffect_Step = 5 Then
        tmrTabHoverEffect.Enabled = False
        mGlowColor = mGlowColor_Bk
        mGlowColor_Sel = mGlowColor_Sel_Bk
    End If
End Sub

Private Sub tmrTabMouseLeave_Timer()
    Dim iPt As POINTAPI
    Dim iHwnd As Long
    
    GetCursorPos iPt
    iHwnd = WindowFromPoint(iPt.X, iPt.Y)
    If iHwnd <> mUserControlHwnd Then
        tmrTabMouseLeave.Enabled = False
        RaiseEvent_TabMouseLeave (mTabUnderMouse)
        mTabUnderMouse = -1
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Dim iChr As String
    Dim iPos As Long
    
    iChr = LCase(Chr(KeyAscii))
    iPos = InStr(mTabSel + 2, mAccessKeys, iChr)
    If iPos = 0 Then
        iPos = InStr(mAccessKeys, iChr)
    End If
    If iPos > 0 Then
        TabSel = iPos - 1
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "ScaleUnits" Then
        PropertyChanged "TabHeight"
        PropertyChanged "TabMaxWidth"
        PropertyChanged "TabMinWidth"
        PropertyChanged "TabSelExtraHeight"
    ElseIf PropertyName = "BackColor" Then
        If mBackColorIsfromAmbient Then BackColor = Ambient.BackColor
        If mTabBackColorIsfromAmbient Then TabBackColor = Ambient.BackColor
    ElseIf PropertyName = "ForeColor" Then
        If mForeColorIsfromAmbient Then ForeColor = Ambient.ForeColor
    End If
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
    If Not mHasFocus Then
        mHasFocus = True
        'PostDrawMessage
        tmrDraw.Enabled = True
    End If
End Sub

Friend Sub StoreVisibleControlsInSelectedTab()
    Dim iCtl As Control
    Dim iCtlName As String
    
    On Error Resume Next
    Set mTabData(mTabSel).Controls = New Collection
    For Each iCtl In UserControl.ContainedControls
        If TypeName(iCtl) = "Line" Then
            Err.Clear
            If iCtl.X1 > -mLeftThresholdHided Then
                If Err.Number = 0 Then
                    iCtlName = ControlName(iCtl)
                    mTabData(mTabSel).Controls.Add iCtlName, iCtlName
                End If
            End If
        Else
            Err.Clear
            If iCtl.Left > -mLeftThresholdHided Then
                If Err.Number = 0 Then
                    iCtlName = ControlName(iCtl)
                    mTabData(mTabSel).Controls.Add iCtlName, iCtlName
                End If
            End If
        End If
    Next
    Err.Clear
End Sub

Private Sub UserControl_Hide()
    mTabBodyReset = True
End Sub

Private Sub UserControl_Initialize()
    mTabUnderMouse = -1
    Set mParentControlsTabStop = New Collection
    Set mParentControlsUseMnemonic = New Collection
    Set mContainedControlsThatAreContainers = New Collection
    Set mSubclassedControlsForPaintingHwnds = New Collection
    Set mSubclassedFramesHwnds = New Collection
    Set mSubclassedControlsForMoveHwnds = New Collection
    mRedraw = True
    mTabOrientation_Prev = -1
    SetDPI
End Sub

' Control code

Private Sub UserControl_InitProperties()
    Dim c As Long
    
    On Error Resume Next
    mUserControlHwnd = UserControl.hWnd
    mAmbientUserMode = Ambient.UserMode
    mDefaultTabHeight = pScaleY(cDefaultTabHeight, vbTwips, vbHimetric)
    If mDefaultTabHeight = 0 Then
        mDefaultTabHeight = 419.8055
    End If
    If mAmbientUserMode Then
        If TypeOf UserControl.Parent Is Form Then
           Set mForm = UserControl.Parent
        End If
    End If
    On Error GoTo 0
    
    mTabSel = 0
    Set mFont = Ambient.Font
    mBackColor = Ambient.BackColor
    mForeColor = Ambient.ForeColor
    mTabSelForeColor = Ambient.ForeColor
    mBackColorIsfromAmbient = True
    mForeColorIsfromAmbient = True
    mEnabled = True
    mTabs = 3
    ReDim mTabData(mTabs - 1)
    For c = 0 To mTabs - 1
        Set mTabData(c).Controls = New Collection
        mTabData(c).Enabled = True
        mTabData(c).Visible = True
        mTabData(c).Caption = "Tab " & CStr(c)
    Next c
    mTabData(mTabSel).Selected = True
    mStyle = ssStyleTabbedDialog
    mWordWrap = True
    mMaskColor = &HFF00FF
    mUseMaskColor = True
    mShowFocusRect = False
    mTabsPerRow = 3
    mTabHeight = mDefaultTabHeight
    mVisualStyles = True
    mTabBackColor = Ambient.BackColor
    mTabSelBackColor = Ambient.BackColor
    mTabBackColorIsfromAmbient = True
    mShowDisabledState = False
    mTabSelFontBold = ssYNAuto
    mChangeControlsBackColor = True
    mTabSelHighlight = False
    mTabHoverHighlight = ssTHHEffect
    mTabWidthStyle = ssTWSAuto
    mShowRowsInPerspective = ssYNAuto
    mTabSeparation = 0
    mTabAppearance = ssTAAuto
    mTabPictureAlignment = ssPicAlignBeforeCaption
    mAutoRelocateControls = ssRelocateAlways
    mSoftEdges = True
    mHandleHighContrastTheme = True
    mRightToLeft = Ambient.RightToLeft
    mBackStyle = ssOpaque
    mAutoTabHeight = True
    
    SetFont
    SetAutoTabHeight
    SetButtonFaceColor
    SetColors
    
    mPropertiesReady = True
    
    mSubclassed = True
#If NOSUBCLASSINIDE Then
    If InIDE Then
        mSubclassed = False
    End If
#End If
    
    If mSubclassed Then
        If mAmbientUserMode Then
            AttachMessage Me, mUserControlHwnd, WM_MOUSEACTIVATE
            AttachMessage Me, mUserControlHwnd, WM_SETFOCUS
            AttachMessage Me, mUserControlHwnd, WM_DRAW
            AttachMessage Me, mUserControlHwnd, WM_INIT
            PostMessage mUserControlHwnd, WM_INIT, 0&, 0&
            mCanPostDrawMessage = True
        Else
            AttachMessage Me, mUserControlHwnd, WM_LBUTTONDOWN
            AttachMessage Me, mUserControlHwnd, WM_LBUTTONUP
            AttachMessage Me, mUserControlHwnd, WM_LBUTTONDBLCLK
        End If
    Else
        mFormIsActive = True
    End If
    UserControl.Size 2500, 1700
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim t As Long
    Dim iAgain As Boolean
    
    RaiseEvent KeyDown(KeyCode, Shift)
    If (KeyCode = vbKeyPageDown And ((Shift And vbCtrlMask) > 0)) Or (KeyCode = vbKeyRight) Or KeyCode = vbKeyTab And ((Shift And vbCtrlMask) > 0) And ((Shift And vbShiftMask) = 0) Then
        t = mTabSel + 1
        If t = mTabs Then t = 0
        Do Until mTabData(t).Enabled And mTabData(t).Visible
            t = t + 1
            If t = mTabs Then
                If iAgain Then Exit Sub
                t = 0
                iAgain = True
            End If
        Loop
        TabSel = t
    ElseIf KeyCode = vbKeyPageUp And ((Shift And vbCtrlMask) > 0) Or (KeyCode = vbKeyLeft) Or KeyCode = vbKeyTab And ((Shift And vbCtrlMask) > 0) And ((Shift And vbShiftMask) > 0) Then
        t = mTabSel - 1
        If t = -1 Then t = mTabs - 1
        Do Until mTabData(t).Enabled And mTabData(t).Visible
            t = t - 1
            If t = -1 Then
                If iAgain Then Exit Sub
                t = mTabs - 1
                iAgain = True
            End If
        Loop
        TabSel = t
    ElseIf (KeyCode = vbKeyDown And ((Shift And vbCtrlMask) = 0)) Then
        SetFocusToNextControlInSameContainer True
    ElseIf (KeyCode = vbKeyUp And ((Shift And vbCtrlMask) = 0)) Then
        SetFocusToNextControlInSameContainer False
    End If
End Sub

Private Sub SetFocusToNextControlInSameContainer(nForward As Boolean)
    Dim iContainerUsr As Object
    Dim iContainerCtl As Object
    Dim iControls As Object
    Dim iHwnds() As Long
    Dim iTabIndexes() As Long
    Dim iCtl As Control
    Dim iTi As Long
    Dim iHwnd As Long
    Dim iEnabled As Boolean
    Dim iVisible As Boolean
    Dim iCount As Long
    Dim iUb As Long
    Dim iTiUsr As Long
    Dim c As Long
    
    On Error Resume Next
    Set iContainerUsr = UserControl.Extender.Container
    If iContainerUsr Is Nothing Then GoTo Exit_Sub
    
    Set iControls = UserControl.Parent.Controls
    If iControls Is Nothing Then GoTo Exit_Sub
    
    iTiUsr = -1
    iTiUsr = UserControl.Extender.TabIndex
    If iTiUsr = -1 Then GoTo Exit_Sub
    
    ReDim iHwnds(100)
    ReDim iTabIndexes(100)
    iUb = 100
    iCount = 0
    
    For Each iCtl In iControls
        Set iContainerCtl = Nothing
        Set iContainerCtl = iCtl.Container
        If iContainerCtl Is iContainerUsr Then
            iTi = -1
            iHwnd = 0
            iEnabled = False
            iVisible = False
            
            iTi = iCtl.TabIndex
            If iTi > -1 Then
                iHwnd = iCtl.hWnd
                If iHwnd > 0 Then
                    iEnabled = iCtl.Enabled
                    iVisible = iCtl.Visible
                    If iEnabled And iVisible Then
                        iCount = iCount + 1
                        If (iCount - 1) > iUb Then
                            iUb = iUb + 100
                            ReDim Preserve iHwnds(iUb)
                            ReDim Preserve iTabIndexes(iUb)
                        End If
                        iHwnds(iCount - 1) = iHwnd
                        iTabIndexes(iCount - 1) = iTi
                    End If
                End If
            End If
        End If
    Next
    
    If iCount > 1 Then ' 1 means that the UserControl is the only control in the container, so there is no other control to focus
        ReDim Preserve iHwnds(iCount - 1)
        ReDim Preserve iTabIndexes(iCount - 1)
        
        ' Bubble sort
        Dim s As Long
        Dim iChanged As Boolean

        s = UBound(iTabIndexes)
        Do
            iChanged = False
            For c = 0 To s - 1
                If iTabIndexes(c) > iTabIndexes(c + 1) Then
                    iTi = iTabIndexes(c)
                    iHwnd = iHwnds(c)
                    iTabIndexes(c) = iTabIndexes(c + 1)
                    iHwnds(c) = iHwnds(c + 1)
                    iTabIndexes(c + 1) = iTi
                    iHwnds(c + 1) = iHwnd
                    iChanged = True
                End If
            Next c
            s = s - 1
        Loop While iChanged
        
        For c = 0 To UBound(iTabIndexes)
            If iTabIndexes(c) = iTiUsr Then
                If nForward Then
                    If c = UBound(iTabIndexes) Then
                        iHwnd = iHwnds(0)
                    Else
                        iHwnd = iHwnds(c + 1)
                    End If
                Else
                    If c = 0 Then
                        iHwnd = iHwnds(UBound(iTabIndexes))
                    Else
                        iHwnd = iHwnds(c - 1)
                    End If
                End If
                SetFocusAPI iHwnd
            End If
        Next c
    End If
    
Exit_Sub:
    Err.Clear
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    mHasFocus = False
    PostDrawMessage
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iX As Single
    Dim iY As Single
    
    iX = X * mXCorrection
    iY = Y * mYCorrection
    
    RaiseEvent MouseDown(Button, Shift, iX, iY)
    
    If Button = 1 Then
        If mTabUnderMouse > -1 Then
            If mTabData(mTabUnderMouse).Enabled Then
                If mTabSel <> mTabUnderMouse Then
                    mHasFocus = True
                    TabSel = mTabUnderMouse
                End If
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iX As Single
    Dim iY As Single
    
    iX = X * mXCorrection
    iY = Y * mYCorrection
    
    RaiseEvent MouseMove(Button, Shift, iX, iY)
    ProcessMouseMove Button, Shift, iX, iY
End Sub

Private Sub ProcessMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim t As Integer
    Dim iX As Long
    Dim iY As Long
    
    iX = pScaleX(X, vbTwips, vbPixels)
    If mRightToLeft Then
        iX = mScaleWidth - iX
    End If
    iY = pScaleX(Y, vbTwips, vbPixels)
    
    ' first check for the active tab, because in some cases it is bigger and can overlap surrounding tabs
    If (mTabSel > -1) And (mTabSel < mTabs) Then
        With mTabData(mTabSel).TabRect
            If iX >= .Left Then
                If iX <= .Right Then
                    If iY >= .Top Then
                        If iY <= .Bottom Then
                            If mTabSel <> mTabUnderMouse Then
                                If mTabUnderMouse > -1 Then
                                    RaiseEvent_TabMouseLeave (mTabUnderMouse)
                                End If
                                RaiseEvent_TabMouseEnter (mTabSel)
                                mTabUnderMouse = mTabSel
                                tmrTabMouseLeave.Enabled = False
                                tmrTabMouseLeave.Enabled = True
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End With
    End If
    
    For t = 0 To mTabs - 1
        If t <> mTabSel Then
            If mTabData(t).Visible And mTabData(t).Enabled Then
                With mTabData(t).TabRect
                    If iX >= .Left Then
                        If iX <= .Right Then
                            If iY >= .Top Then
                                If iY <= .Bottom Then
                                    If t <> mTabUnderMouse Then
                                        If mTabUnderMouse > -1 Then
                                            RaiseEvent_TabMouseLeave (mTabUnderMouse)
                                        End If
                                        RaiseEvent_TabMouseEnter (t)
                                        mTabUnderMouse = t
                                        tmrTabMouseLeave.Enabled = False
                                        tmrTabMouseLeave.Enabled = True
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End With
            End If
        End If
    Next t
    If mTabUnderMouse > -1 Then
        tmrTabMouseLeave.Enabled = False
        RaiseEvent_TabMouseLeave (mTabUnderMouse)
    End If
    mTabUnderMouse = -1
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iX As Single
    Dim iY As Single
    
    iX = X * mXCorrection
    iY = Y * mYCorrection
    
    RaiseEvent MouseUp(Button, Shift, iX, iY)
    If mTabUnderMouse > -1 Then
        If Button = 2 Then
            RaiseEvent TabRightClick(mTabUnderMouse, Shift, iX, iY)
        End If
    End If
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim c As Long
    Dim c2 As Long
    Dim iStr As String
    Dim iStr2 As String
    Dim iAllCtlNames As Collection
    Dim iLeftShiftToHideWhenSaved As Long
    
    On Error Resume Next
    mUserControlHwnd = UserControl.hWnd
    mAmbientUserMode = Ambient.UserMode
    mDefaultTabHeight = pScaleY(cDefaultTabHeight, vbTwips, vbHimetric)
    If mDefaultTabHeight = 0 Then
        mDefaultTabHeight = 419.8055
    End If
    If mAmbientUserMode Then
        If TypeOf UserControl.Parent Is Form Then
            Set mForm = UserControl.Parent
        End If
    End If
    On Error GoTo 0
    
    iLeftShiftToHideWhenSaved = PropBag.ReadProperty("LeftShiftToHideWhenSaved", 75000)
    If iLeftShiftToHideWhenSaved <> mLeftShiftToHide Then
        mPendingLeftShift = iLeftShiftToHideWhenSaved - mLeftShiftToHide
    End If
    mTabs = PropBag.ReadProperty("Tabs", 3)
    mBackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    If mBackColor = Ambient.BackColor Then mBackColorIsfromAmbient = True
    mForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
    mTabSelForeColor = PropBag.ReadProperty("TabSelForeColor", mForeColor)
    If mForeColor = Ambient.ForeColor Then mForeColorIsfromAmbient = True
    Set mFont = PropBag.ReadProperty("Font", Nothing)
    mEnabled = PropBag.ReadProperty("Enabled", True)
    mTabsPerRow = PropBag.ReadProperty("TabsPerRow", 3)
    If mTabsPerRow < 1 Then mTabsPerRow = 3
    mTabSel = PropBag.ReadProperty("Tab", 0)
    mTabOrientation = PropBag.ReadProperty("TabOrientation", ssTabOrientationTop)
    mShowFocusRect = PropBag.ReadProperty("ShowFocusRect", False)
    mWordWrap = PropBag.ReadProperty("WordWrap", True)
    mStyle = PropBag.ReadProperty("Style", ssStyleTabbedDialog)
    mTabHeight = PropBag.ReadProperty("TabHeight", mDefaultTabHeight)    ' in Himetric, for compatibility with the original SSTab
    If pScaleY(mTabHeight, vbHimetric, vbPixels) < 1 Then mTabHeight = pScaleY(1, vbPixels, vbHimetric)
    mTabMaxWidth = PropBag.ReadProperty("TabMaxWidth", 0)  ' in Himetric, for compatibility with the original SSTab
    mTabMinWidth = PropBag.ReadProperty("TabMinWidth", 0)  ' in Himetric
    mMousePointer = PropBag.ReadProperty("MousePointer", ssDefault)
    Set mMouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    mOLEDropMode = PropBag.ReadProperty("OLEDropMode", ssOLEDropNone)
    mMaskColor = PropBag.ReadProperty("MaskColor", &HFF00FF)
    mUseMaskColor = PropBag.ReadProperty("UseMaskColor", True)
    mTabSelExtraHeight = PropBag.ReadProperty("TabSelExtraHeight", 0)
    If mTabSelExtraHeight < 0 Then mTabSelExtraHeight = 0
    mTabSelHighlight = PropBag.ReadProperty("TabSelHighlight", False)
    mTabHoverHighlight = PropBag.ReadProperty("TabHoverHighlight", ssTHHEffect)
    mTabSelFontBold = PropBag.ReadProperty("TabSelFontBold", ssYNAuto)
    mVisualStyles = PropBag.ReadProperty("Themed", True)
    mTabBackColor = PropBag.ReadProperty("TabBackColor", Ambient.BackColor)
    mTabSelBackColor = PropBag.ReadProperty("TabSelBackColor", mTabBackColor)
    If mTabBackColor = Ambient.BackColor Then mTabBackColorIsfromAmbient = True
    mShowDisabledState = PropBag.ReadProperty("ShowDisabledState", False)
    mChangeControlsBackColor = PropBag.ReadProperty("ChangeControlsBackColor", True)
    mTabWidthStyle = PropBag.ReadProperty("TabWidthStyle", ssTWSAuto)
    mShowRowsInPerspective = PropBag.ReadProperty("ShowRowsInPerspective", ssYNAuto)
    mTabSeparation = PropBag.ReadProperty("TabSeparation", 0)
    mTabAppearance = PropBag.ReadProperty("TabAppearance", ssTAAuto)
    mTabPictureAlignment = PropBag.ReadProperty("TabPictureAlignment", ssPicAlignBeforeCaption)
    mAutoRelocateControls = PropBag.ReadProperty("AutoRelocateControls", ssRelocateAlways)
    mSoftEdges = PropBag.ReadProperty("SoftEdges", True)
    mRightToLeft = PropBag.ReadProperty("RightToLeft", Ambient.RightToLeft)
    If mRightToLeft Then
        SetLayout GetDC(picDraw.hWnd), LAYOUT_RTL Or LAYOUT_BITMAPORIENTATIONPRESERVED
    End If
    mHandleHighContrastTheme = PropBag.ReadProperty("HandleHighContrastTheme", True)
    mBackStyle = PropBag.ReadProperty("BackStyle", ssOpaque)
    mAutoTabHeight = PropBag.ReadProperty("AutoTabHeight", False)
    
    Set UserControl.MouseIcon = mMouseIcon
    UserControl.MousePointer = mMousePointer
    
    If mFont Is Nothing Then
        Set mFont = Ambient.Font
    End If
    If mFont Is Nothing Then
        Set mFont = UserControl.Font
    End If
    UserControl.Enabled = mEnabled Or (Not mAmbientUserMode)
    
    ReDim mTabData(mTabs - 1)
    Set iAllCtlNames = New Collection
    For c = 0 To mTabs - 1
        Set mTabData(c).Controls = New Collection
        Set mTabData(c).Picture = PropBag.ReadProperty("TabPicture(" & CStr(c) & ")", Nothing)
        If Not mTabData(c).Picture Is Nothing Then
            If mTabData(c).Picture.Handle = 0 Then Set mTabData(c).Picture = Nothing
        End If
        Set mTabData(c).Pic16 = PropBag.ReadProperty("TabPic16(" & CStr(c) & ")", Nothing)
        If Not mTabData(c).Pic16 Is Nothing Then
            If mTabData(c).Pic16.Handle = 0 Then Set mTabData(c).Pic16 = Nothing
        End If
        Set mTabData(c).Pic20 = PropBag.ReadProperty("TabPic20(" & CStr(c) & ")", Nothing)
        If Not mTabData(c).Pic20 Is Nothing Then
            If mTabData(c).Pic20.Handle = 0 Then Set mTabData(c).Pic20 = Nothing
        End If
        Set mTabData(c).Pic24 = PropBag.ReadProperty("TabPic24(" & CStr(c) & ")", Nothing)
        If Not mTabData(c).Pic24 Is Nothing Then
            If mTabData(c).Pic24.Handle = 0 Then Set mTabData(c).Pic24 = Nothing
        End If
        mTabData(c).Caption = PropBag.ReadProperty("TabCaption(" & CStr(c) & ")", "")
        mTabData(c).ToolTipText = PropBag.ReadProperty("TabToolTipText(" & CStr(c) & ")", "")
        For c2 = 0 To PropBag.ReadProperty("Tab(" & c & ").ControlCount", 0) - 1
            iStr = PropBag.ReadProperty("Tab(" & c & ").Control(" & c2 & ")", "")
            If iStr <> "" Then
                iStr2 = ""
                On Error Resume Next
                iStr2 = iAllCtlNames(iStr)
                On Error GoTo 0
                If iStr2 = "" Then
                    mTabData(c).Controls.Add iStr, iStr
                    iAllCtlNames.Add iStr, iStr
                End If
            End If
        Next
        mTabData(c).Enabled = True
        mTabData(c).Visible = True
    Next c
    mTabData(mTabSel).Selected = True
    
    SetFont
    SetAutoTabHeight
    SetButtonFaceColor
    SetColors
    CheckIfThereAreTabsToolTipTexts
    
    mSubclassed = True
#If NOSUBCLASSINIDE Then
    If InIDE Then
        mSubclassed = False
    End If
#End If
    
    If mSubclassed Then
        If mAmbientUserMode Then
            AttachMessage Me, mUserControlHwnd, WM_MOUSEACTIVATE
            AttachMessage Me, mUserControlHwnd, WM_SETFOCUS
            AttachMessage Me, mUserControlHwnd, WM_DRAW
            AttachMessage Me, mUserControlHwnd, WM_INIT
            PostMessage mUserControlHwnd, WM_INIT, 0&, 0&
            mCanPostDrawMessage = True
        Else
            AttachMessage Me, mUserControlHwnd, WM_LBUTTONDOWN
            AttachMessage Me, mUserControlHwnd, WM_LBUTTONUP
            AttachMessage Me, mUserControlHwnd, WM_LBUTTONDBLCLK
        End If
    Else
        mFormIsActive = True
    End If
    mPropertiesReady = True
    
    PostDrawMessage
    If tmrDraw.Enabled Then
        Draw
    End If
End Sub

Private Sub UserControl_Resize()
    ResetCachedThemeImages
    If mAmbientUserMode Then
        PostDrawMessage
    Else
        tmrDraw.Enabled = True
    End If
    RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
    If mUserControlTerminated Then Exit Sub
    If mUserControlShown Then
        Exit Sub
    End If
    If mHandleHighContrastTheme Then CheckHighContrastTheme
    
    If mPendingLeftShift <> 0 Then
        DoPendingLeftShift
    End If
    
    If mAmbientUserMode And mSubclassed Then
        If (mFormHwnd = 0) Then
            mFormHwnd = GetAncestor(UserControl.ContainerHwnd, GA_ROOT)
            mFormIsActive = GetForegroundWindow = mFormHwnd
            If (mFormHwnd <> 0) And mAmbientUserMode Then
                AttachMessage Me, mFormHwnd, WM_SYSCOLORCHANGE
                AttachMessage Me, mFormHwnd, WM_THEMECHANGED
                AttachMessage Me, mFormHwnd, WM_NCACTIVATE
                AttachMessage Me, mFormHwnd, WM_GETDPISCALEDSIZE
            End If
        End If
        
        Dim iAuxLeft As Long
        Dim iHwnd As Long
        Dim c As Long
        Dim iCtlName As String
        Dim iCtl As Control
        Dim iIsLine As Boolean
        
        On Error Resume Next
        If mSubclassedControlsForMoveHwnds.Count > 0 Then
            For c = 1 To mSubclassedControlsForMoveHwnds.Count
                iHwnd = mSubclassedControlsForMoveHwnds(c)
                DetachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
            Next c
            Set mSubclassedControlsForMoveHwnds = New Collection
        End If
    
        For Each iCtl In UserControl.ContainedControls
            iAuxLeft = 0
            iIsLine = False
            If TypeName(iCtl) = "Line" Then
                iAuxLeft = iCtl.X1
                iIsLine = True
            Else
                iAuxLeft = iCtl.Left
            End If
            If iAuxLeft >= -mLeftThresholdHided Then
                iCtlName = ControlName(iCtl)
                If Not ControlIsInTab(iCtlName, mTabSel) Then
                    If iIsLine Then
                        iCtl.X1 = iCtl.X1 - mLeftShiftToHide
                        iCtl.X2 = iCtl.X2 - mLeftShiftToHide
                    Else
                        iCtl.Left = iCtl.Left - mLeftShiftToHide
                    End If
                    iAuxLeft = iAuxLeft - mLeftShiftToHide
                End If
            End If
            If iAuxLeft < -mLeftThresholdHided Then
                iHwnd = 0
                iHwnd = iCtl.hWnd
                If iHwnd <> 0 Then
                    mSubclassedControlsForMoveHwnds.Add iHwnd
                    AttachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
                End If
            End If
        Next
        On Error GoTo 0
    End If
    
    If mChangeControlsBackColor Then
        If Not mChangedControlsBackColor Then
            SetControlsBackColor mTabSelBackColor
            mChangedControlsBackColor = True
        End If
    End If
    
    If mAmbientUserMode Then
        If Not mTabStopsInitialized Then
            StoreControlsTabStop True
            mTabStopsInitialized = True
        End If
        If mForm Is Nothing Then SubclassControlsPainting
    Else
        HideAllContainedControls
        MakeContainedControlsInSelTabVisible
        If Not IsMsgBoxShown Then CheckContainedControlsConsistency
    End If
    mUserControlShown = True
    SubclassControlsPainting
    If (Not mFirstDraw) Or mDrawMessagePosted Then
        Draw
        mFirstDraw = True
    End If
    RaiseEvent TabSelChange
End Sub

Private Sub DoPendingLeftShift()
    Dim iCtl As Control
    Dim iIsLine As Boolean
    Dim iAuxLeft As Long
    
    If mPendingLeftShift <> 0 Then
        For Each iCtl In UserControl.ContainedControls
            iAuxLeft = 0
            iIsLine = False
            On Error Resume Next
            If TypeName(iCtl) = "Line" Then
                iAuxLeft = iCtl.X1
                iIsLine = True
            Else
                iAuxLeft = iCtl.Left
            End If
            On Error GoTo 0
            If iAuxLeft < -mLeftThresholdHided Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 + mPendingLeftShift
                    iCtl.X2 = iCtl.X2 + mPendingLeftShift
                Else
                    iCtl.Left = iCtl.Left + mPendingLeftShift
                End If
            End If
        Next
        mPendingLeftShift = 0
    End If

End Sub

Private Function ControlIsInTab(nCtlName As String, nTab As Integer) As Boolean
    Dim c As Long
    
    For c = 1 To mTabData(nTab).Controls.Count
        If mTabData(nTab).Controls(c) = nCtlName Then
            ControlIsInTab = True
            Exit Function
        End If
    Next c
End Function
    
Private Sub UserControl_Terminate()
    DoTerminate
End Sub

Private Sub DoTerminate()
    Dim c As Long
    Dim iHwnd As Long
    
    If mUserControlTerminated Then Exit Sub
    mUserControlTerminated = True
    
    If (mFormHwnd <> 0) And mAmbientUserMode Then
        DetachMessage Me, mFormHwnd, WM_SYSCOLORCHANGE
        DetachMessage Me, mFormHwnd, WM_THEMECHANGED
        DetachMessage Me, mFormHwnd, WM_NCACTIVATE
        DetachMessage Me, mFormHwnd, WM_GETDPISCALEDSIZE
    End If
    If mSubclassed Then
        If mAmbientUserMode Then
            DetachMessage Me, mUserControlHwnd, WM_MOUSEACTIVATE
            DetachMessage Me, mUserControlHwnd, WM_SETFOCUS
            DetachMessage Me, mUserControlHwnd, WM_DRAW
            DetachMessage Me, mUserControlHwnd, WM_INIT
            mCanPostDrawMessage = False
        Else
            DetachMessage Me, mUserControlHwnd, WM_LBUTTONDOWN
            DetachMessage Me, mUserControlHwnd, WM_LBUTTONUP
            DetachMessage Me, mUserControlHwnd, WM_LBUTTONDBLCLK
        End If
    End If
    mSubclassed = False
    
    tmrTabMouseLeave.Enabled = False
    tmrDraw.Enabled = False
    tmrCancelDoubleClick.Enabled = False
    tmrCheckContainedControlsAdditionDesignTime.Enabled = False
    tmrTabHoverEffect.Enabled = False
    
    Set mParentControlsTabStop = Nothing
    Set mParentControlsUseMnemonic = Nothing
    Set mContainedControlsThatAreContainers = Nothing
    
    For c = 1 To mSubclassedControlsForPaintingHwnds.Count
        iHwnd = mSubclassedControlsForPaintingHwnds(c)
        DetachMessage Me, iHwnd, WM_PAINT
        DetachMessage Me, iHwnd, WM_MOVE
    Next c
    Set mSubclassedControlsForPaintingHwnds = Nothing
    
    For c = 1 To mSubclassedFramesHwnds.Count
        iHwnd = mSubclassedFramesHwnds(c)
        DetachMessage Me, iHwnd, WM_PRINTCLIENT
        DetachMessage Me, iHwnd, WM_MOUSELEAVE
    Next c
    Set mSubclassedFramesHwnds = Nothing
    
    For c = 1 To mSubclassedControlsForMoveHwnds.Count
        iHwnd = mSubclassedControlsForMoveHwnds(c)
        DetachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
    Next c
    Set mSubclassedControlsForMoveHwnds = Nothing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim c As Long
    Dim c2 As Long
    
    StoreVisibleControlsInSelectedTab
    
    PropBag.WriteProperty "Tabs", mTabs, 3
    PropBag.WriteProperty "BackColor", mBackColor, Ambient.BackColor
    PropBag.WriteProperty "ForeColor", mForeColor, Ambient.ForeColor
    PropBag.WriteProperty "TabSelForeColor", mTabSelForeColor, mForeColor
    PropBag.WriteProperty "Font", mFont, Nothing
    PropBag.WriteProperty "Enabled", mEnabled, True
    PropBag.WriteProperty "TabsPerRow", mTabsPerRow, 3
    PropBag.WriteProperty "Tab", mTabSel, 0
    PropBag.WriteProperty "TabOrientation", mTabOrientation, ssTabOrientationTop
    PropBag.WriteProperty "ShowFocusRect", mShowFocusRect, False
    PropBag.WriteProperty "WordWrap", mWordWrap, True
    PropBag.WriteProperty "Style", mStyle, ssStyleTabbedDialog
    PropBag.WriteProperty "TabHeight", Round(mTabHeight), Round(mDefaultTabHeight)  ' in Himetric, for compatibility with the original SSTab
    PropBag.WriteProperty "TabMaxWidth", Round(mTabMaxWidth), 0  ' in Himetric, for compatibility with the original SSTab
    PropBag.WriteProperty "TabMinWidth", Round(mTabMinWidth), 0 ' in Himetric
    PropBag.WriteProperty "MousePointer", mMousePointer, ssDefault
    PropBag.WriteProperty "MouseIcon", mMouseIcon, Nothing
    PropBag.WriteProperty "OLEDropMode", mOLEDropMode, ssOLEDropNone
    PropBag.WriteProperty "MaskColor", mMaskColor, &HFF00FF
    PropBag.WriteProperty "UseMaskColor", mUseMaskColor, True
    PropBag.WriteProperty "TabSelExtraHeight", Round(mTabSelExtraHeight), 0
    PropBag.WriteProperty "TabSelHighlight", mTabSelHighlight, False
    PropBag.WriteProperty "TabHoverHighlight", mTabHoverHighlight, ssTHHEffect
    PropBag.WriteProperty "TabSelFontBold", mTabSelFontBold, ssYNAuto
    PropBag.WriteProperty "Themed", mVisualStyles, True
    PropBag.WriteProperty "TabBackColor", mTabBackColor, Ambient.BackColor
    PropBag.WriteProperty "TabSelBackColor", mTabSelBackColor, mTabBackColor
    PropBag.WriteProperty "ShowDisabledState", mShowDisabledState, False
    PropBag.WriteProperty "ChangeControlsBackColor", mChangeControlsBackColor, True
    PropBag.WriteProperty "TabWidthStyle", mTabWidthStyle, ssTWSAuto
    PropBag.WriteProperty "ShowRowsInPerspective", mShowRowsInPerspective, ssYNAuto
    PropBag.WriteProperty "TabSeparation", mTabSeparation, 0
    PropBag.WriteProperty "TabAppearance", mTabAppearance, ssTAAuto
    PropBag.WriteProperty "TabPictureAlignment", mTabPictureAlignment, ssPicAlignBeforeCaption
    PropBag.WriteProperty "AutoRelocateControls", mAutoRelocateControls, ssRelocateAlways
    PropBag.WriteProperty "SoftEdges", mSoftEdges, True
    PropBag.WriteProperty "RightToLeft", mRightToLeft, Ambient.RightToLeft
    PropBag.WriteProperty "HandleHighContrastTheme", mHandleHighContrastTheme, True
    PropBag.WriteProperty "LeftShiftToHideWhenSaved", mLeftShiftToHide + mPendingLeftShift, 75000
    PropBag.WriteProperty "LeftThresholdHidedWhenSaved", mLeftThresholdHided, 15000
    PropBag.WriteProperty "BackStyle", mBackStyle, ssOpaque
    PropBag.WriteProperty "AutoTabHeight", mAutoTabHeight, False
    
    For c = 0 To mTabs - 1
        PropBag.WriteProperty "TabPicture(" & CStr(c) & ")", mTabData(c).Picture, Nothing
        PropBag.WriteProperty "TabPic16(" & CStr(c) & ")", mTabData(c).Pic16, Nothing
        PropBag.WriteProperty "TabPic20(" & CStr(c) & ")", mTabData(c).Pic20, Nothing
        PropBag.WriteProperty "TabPic24(" & CStr(c) & ")", mTabData(c).Pic24, Nothing
        PropBag.WriteProperty "TabCaption(" & CStr(c) & ")", mTabData(c).Caption, ""
        PropBag.WriteProperty "TabToolTipText(" & CStr(c) & ")", mTabData(c).ToolTipText, ""
        PropBag.WriteProperty "Tab(" & c & ").ControlCount", mTabData(c).Controls.Count
        For c2 = 1 To mTabData(c).Controls.Count
            PropBag.WriteProperty "Tab(" & c & ").Control(" & c2 - 1 & ")", mTabData(c).Controls(c2), ""
        Next
    Next c
End Sub

Private Sub Draw()
    Dim iTabWidth As Single
    Dim iTabData As T_TabData
    Dim iTabSelExtraHeight As Long
    Dim iLng As Long
    Dim t As Long
    Dim ctv As Long
    Dim iVisibleTabs As Long
    Dim iPosH As Long
    Dim iRow As Long ' this variable is reused and not always means the same thing
    Dim iRowPerspectiveSpace As Long
    Dim iAllRowsPerspectiveSpace As Long
    Dim iTabHeight As Long
    Dim iTmpRect As RECT
    Dim iLastVisibleTab As Long
    Dim iLastVisibleTab_Prev As Long
    Dim iScaleWidth As Long
    Dim iScaleHeight As Long
    Dim iTabMaxWidth As Long
    Dim iTabMinWidth As Long
    Dim iTabLeft As Long
    Dim iShowsRowsPerspective As Boolean
    Dim iRowTabCount As Long
    Dim iAccumulatedTabWith As Long
    Dim iTotalTabWidth As Long
    Dim iTabStretchRatio As Single
    Dim iTabWidthStyle As vbExTabWidthStyleConstants
    Dim iARPSTmp As Long
    Dim iAvailableSpaceForTabs As Long
    Dim iRowsStretchRatio() As Single
    Dim iRowsStretchRatio_StartingRow As Long
    Dim iRowsStretchRatio_AccumulatedTabWidth As Long
    Dim R As Long
    Dim iAccumulatedAdditionalFixedTabSpace As Long
    Dim iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth As Long
    Dim iSng As Single
    Dim iDecreaseStretchRatio As Boolean
    Dim iIncreaseStretchRatio As Boolean
    Dim iDoNotDecreaseStretchRatio As Boolean
    Dim iStyle2 As vbExStyleConstants
    Dim iMessage As T_MSG
    Dim iAlreadyNeedToBePainted As Boolean
    
    If mUserControlTerminated Then Exit Sub
    
    If Not mRedraw Then
        mNeedToDraw = True
        If Not mEnsureDrawn Then
            Exit Sub
        End If
    End If
    If Not mPropertiesReady Then
        PostDrawMessage
        Exit Sub
    End If
    tmrDraw.Enabled = False
    PeekMessage iMessage, mUserControlHwnd, WM_DRAW, WM_DRAW, PM_REMOVE ' remove posted message, if any
    mDrawMessagePosted = False
    If Not mAccessKeysSet Then
        If mAmbientUserMode Then SetAccessKeys
    End If
    UserControl.ScaleMode = vbPixels
    mScaleWidth = UserControl.ScaleWidth
    mScaleHeight = UserControl.ScaleHeight
    If Not ((mScaleWidth > 0) And (mScaleHeight > 0)) Then
        UserControl.ScaleMode = vbTwips
        Exit Sub
    End If
    If Not mFirstDraw Then mFirstDraw = True
    
    mDrawing = True
    mControlIsThemed = mVisualStyles And (IsAppThemeEnabled Or mForceVisualStyles) And (mBackStyle <> ssTransparent)
    If mControlIsThemed Then
        If mTheme <> 0 Then
            CloseThemeData mTheme
            mTheme = 0
        End If
        mTheme = OpenThemeData(mUserControlHwnd, StrPtr("Tab"))
        If mTheme = 0 Then
            mControlIsThemed = False
        End If
        If mControlIsThemed Then
            SetThemeExtraData
        End If
    End If
    If mControlIsThemed Then
        iStyle2 = ssStylePropertyPage
    ElseIf mStyle = ssStyleTabStrip Then
        iStyle2 = ssStylePropertyPage
    Else
        iStyle2 = mStyle
    End If
    
    iLng = mTabAppearance2
    If mTabAppearance = ssTAAuto Then
        If (iStyle2 <> ssStyleTabbedDialog) Then
            mTabAppearance2 = ssTAPropertyPage
        Else
            mTabAppearance2 = ssTATabbedDialog
        End If
    Else
        mTabAppearance2 = mTabAppearance
    End If
    mAppearanceIsPP = (mTabAppearance2 = ssTAPropertyPage) Or (mTabAppearance2 = ssTAPropertyPageRounded) Or mControlIsThemed
    If mTabAppearance2 <> iLng Then ResetCachedThemeImages
    
    iTabHeight = pScaleY(mTabHeight, vbHimetric, vbPixels)
    iTabSelExtraHeight = pScaleY(mTabSelExtraHeight, vbHimetric, vbPixels)
    iTabMaxWidth = pScaleX(mTabMaxWidth, vbHimetric, vbPixels)
    iTabMinWidth = pScaleX(mTabMinWidth, vbHimetric, vbPixels)
    If mTabWidthStyle = ssTWSAuto Then
        If mStyle = ssStyleTabStrip Then
            iTabWidthStyle = ssTWSJustified
        ElseIf mStyle = ssStylePropertyPage Then
            iTabWidthStyle = ssTWSNonJustified
        Else
            iTabWidthStyle = ssTWSFixed
        End If
    Else
        iTabWidthStyle = mTabWidthStyle
    End If
    If mShowRowsInPerspective = ssYNAuto Then
        If mStyle = ssStyleTabStrip Then
            iShowsRowsPerspective = False
        Else
            iShowsRowsPerspective = (iTabWidthStyle <> ssTWSJustified)
        End If
    Else
        iShowsRowsPerspective = CBool(mShowRowsInPerspective)
    End If
    If iShowsRowsPerspective Then
        iRowPerspectiveSpace = pScaleX(cRowPerspectiveSpace, vbTwips, vbPixels)
    End If
    
    mTabSeparation2 = mTabSeparation
    If mControlIsThemed Then
        mTabSeparation2 = mTabSeparation2 - 2
        If mTabSeparation2 < 0 Then mTabSeparation2 = 0
    End If
    
    If (mTabOrientation = ssTabOrientationTop) Then
        m3DShadowH = m3DShadow
        m3DShadowV = m3DShadow
        m3DHighlightH = m3DHighlight
        m3DHighlightV = m3DHighlight
        m3DShadowH_Sel = m3DShadow_Sel
        m3DShadowV_Sel = m3DShadow_Sel
        m3DHighlightH_Sel = m3DHighlight_Sel
        m3DHighlightV_Sel = m3DHighlight_Sel
    ElseIf mTabOrientation = ssTabOrientationBottom Then
        m3DShadowH = m3DHighlight
        m3DShadowV = m3DShadow
        m3DHighlightH = m3DShadow
        m3DHighlightV = m3DHighlight
        m3DShadowH_Sel = m3DHighlight_Sel
        m3DShadowV_Sel = m3DShadow_Sel
        m3DHighlightH_Sel = m3DShadow_Sel
        m3DHighlightV_Sel = m3DHighlight_Sel
    ElseIf mTabOrientation = ssTabOrientationLeft Then
        m3DShadowH = m3DShadow
        m3DShadowV = m3DHighlight
        m3DHighlightH = m3DHighlight
        m3DHighlightV = m3DShadow
        m3DShadowH_Sel = m3DShadow_Sel
        m3DShadowV_Sel = m3DHighlight_Sel
        m3DHighlightH_Sel = m3DHighlight_Sel
        m3DHighlightV_Sel = m3DShadow_Sel
    ElseIf mTabOrientation = ssTabOrientationRight Then
        m3DShadowH = m3DHighlight
        m3DShadowV = m3DShadow
        m3DHighlightH = m3DShadow
        m3DHighlightV = m3DHighlight
        m3DShadowH_Sel = m3DHighlight_Sel
        m3DShadowV_Sel = m3DShadow_Sel
        m3DHighlightH_Sel = m3DShadow_Sel
        m3DHighlightV_Sel = m3DHighlight_Sel
    End If
    
    If mBackStyle = ssOpaque Then
        If mEnabled Or (Not mAmbientUserMode) Or (Not mShowDisabledState) Then
            mTabBackColor2 = mTabBackColor
            mTabSelBackColor2 = mTabSelBackColor
        Else
            mTabBackColor2 = mTabBackColorDisabled
            mTabSelBackColor2 = mTabSelBackColorDisabled
        End If
    Else
        mTabBackColor2 = mTabBackColor
        TranslateColor mTabBackColor2, 0, mTabBackColor2
        TranslateColor mTabSelBackColor, 0, iLng
        If mTabBackColor2 = iLng Then
            mTabBackColor2 = mTabBackColor2 Xor &H1
        End If
        mTabSelBackColor2 = mTabSelBackColor
        UserControl.MaskColor = mTabSelBackColor
    End If
    UserControl.BackStyle = mBackStyle
    
    If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
        iScaleWidth = mScaleWidth
        iScaleHeight = mScaleHeight
    Else
        iScaleWidth = mScaleHeight
        iScaleHeight = mScaleWidth
    End If
    
    ' measure tab captions and pic width
    ctv = -1
    If iTabWidthStyle = ssTWSJustified Then
        iTotalTabWidth = 0
        mRows = 1
    End If
    iVisibleTabs = 0
    For t = 0 To mTabs - 1
        If mTabData(t).Visible Then
            iVisibleTabs = iVisibleTabs + 1
            ctv = ctv + 1
            If (iTabWidthStyle = ssTWSNonJustified) Or (iTabWidthStyle = ssTWSJustified) Then
                iLng = MeasureTabPictureAndCaption(t)
                If iTabMinWidth > 0 Then
                    If (iLng + 10) < iTabMinWidth Then
                        iLng = iTabMinWidth - 10
                    End If
                End If
                If iTabMaxWidth > 0 Then
                    If (iLng + 10) > iTabMaxWidth Then
                        iLng = iTabMaxWidth - 10
                    End If
                End If
                mTabData(t).PicAndCaptionWidth = iLng
            End If
        End If
    Next t
    
    If (iVisibleTabs = 0) Or (mTabSel = -1) Then
        Set UserControl.Picture = Nothing
        mTabBodyReset = True
        GoTo TheExit:
    End If
    
    ' set data about tabs placement on rows
    iLastVisibleTab = 0
    If iTabWidthStyle <> ssTWSJustified Then
        iRow = 0
        iPosH = 0
        ctv = 0
        For t = 0 To mTabs - 1
            If mTabData(t).Visible Then
                mTabData(t).TopTab = False
                ctv = ctv + 1
                iLastVisibleTab = t
                mTabData(t).LeftTab = False
                mTabData(t).RightTab = False
                iPosH = iPosH + 1
                If iPosH > mTabsPerRow Then
                    iPosH = 1
                    iRow = iRow + 1
                End If
                mTabData(t).PosH = iPosH
                If iPosH = 1 Then
                    mTabData(t).LeftTab = True
                End If
                If (iPosH = mTabsPerRow) Or (ctv = iVisibleTabs) Then
                    mTabData(t).RightTab = True
                End If
                mTabData(t).Row = iRow
            Else
                mTabData(t).Row = -1
            End If
        Next t
        mRows = iRow + 1
    Else
        ' define what tabs to place on each row when tabs are justified
        ' define what tabs to place on each row when tabs are justified
        ' step 1: calculate the number of rows that will be needed and the iRowsStretchRatio for each row (that will be needed in the step 2)
        iARPSTmp = 0
        Do
            iAllRowsPerspectiveSpace = iARPSTmp
            iAvailableSpaceForTabs = (iScaleWidth - iAllRowsPerspectiveSpace - IIf(mAppearanceIsPP, 4, 0))
            iAccumulatedTabWith = 0
            iAccumulatedAdditionalFixedTabSpace = 0
            iRow = 0
            ReDim iRowsStretchRatio(0)
            iRowsStretchRatio_StartingRow = 0
            iRowsStretchRatio_AccumulatedTabWidth = 0
            iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth = 0
            iRowTabCount = 0
            For t = 0 To mTabs - 1
                If mTabData(t).Visible Then
                    If (iAccumulatedTabWith + iAccumulatedAdditionalFixedTabSpace + mTabData(t).PicAndCaptionWidth + 10) > iAvailableSpaceForTabs Then
                        If iRowTabCount = 0 Then ' this only tab alone passes the available space in the row (and it is the first one or all the previous tabs also entered here)
                            If t < (mTabs - 1) Then
                                iRowsStretchRatio(iRow) = 1
                                iRow = iRow + 1
                                iRowsStretchRatio_StartingRow = iRow
                                iRowsStretchRatio_AccumulatedTabWidth = 0
                                iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth = 0
                                ReDim Preserve iRowsStretchRatio(iRow)
                            End If
                            iRowTabCount = 0
                            iAccumulatedTabWith = 0
                            iAccumulatedAdditionalFixedTabSpace = 0
                        ElseIf iRowTabCount = 1 Then ' this only tab alone passes the available space in the row (and it comes from a previus attempt to put it in the previous row)
                            iSng = ((iRow - iRowsStretchRatio_StartingRow) * iAvailableSpaceForTabs - iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth) / iRowsStretchRatio_AccumulatedTabWidth
                            If iSng < 1 Then
                                iDoNotDecreaseStretchRatio = True
                                iSng = 1
                            End If
                            For R = iRowsStretchRatio_StartingRow To iRow - 1
                                iRowsStretchRatio(R) = iSng
                            Next R
                            iRowsStretchRatio(iRow) = 1
                            iRow = iRow + 1
                            iRowsStretchRatio_StartingRow = iRow
                            ReDim Preserve iRowsStretchRatio(iRow)
                            iRowTabCount = 1
                            iAccumulatedTabWith = mTabData(t).PicAndCaptionWidth
                            iAccumulatedAdditionalFixedTabSpace = 10 + mTabSeparation2
                            iRowsStretchRatio_AccumulatedTabWidth = iAccumulatedTabWith
                            iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth = iAccumulatedAdditionalFixedTabSpace
                        Else
                            iRow = iRow + 1
                            ReDim Preserve iRowsStretchRatio(iRow)
                            iRowTabCount = 1
                            iAccumulatedTabWith = mTabData(t).PicAndCaptionWidth
                            iAccumulatedAdditionalFixedTabSpace = 10 + mTabSeparation2
                            iRowsStretchRatio_AccumulatedTabWidth = iRowsStretchRatio_AccumulatedTabWidth + iAccumulatedTabWith
                            iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth = iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth + iAccumulatedAdditionalFixedTabSpace
                        End If
                    Else
                        iAccumulatedTabWith = iAccumulatedTabWith + mTabData(t).PicAndCaptionWidth
                        iAccumulatedAdditionalFixedTabSpace = iAccumulatedAdditionalFixedTabSpace + 10 + mTabSeparation2
                        iRowsStretchRatio_AccumulatedTabWidth = iRowsStretchRatio_AccumulatedTabWidth + mTabData(t).PicAndCaptionWidth
                        iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth = iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth + 10 + mTabSeparation2
                        iRowTabCount = iRowTabCount + 1
                    End If
                End If
            Next t
            If iRowsStretchRatio_AccumulatedTabWidth > 0 Then
                iSng = ((iRow - iRowsStretchRatio_StartingRow + 1) * iAvailableSpaceForTabs - iRowsStretchRatio_AccumulatedAdditionalFixedTabWidth) / iRowsStretchRatio_AccumulatedTabWidth
                If iSng < 1 Then
                    iDoNotDecreaseStretchRatio = True
                    iSng = 1
                End If
                For R = iRowsStretchRatio_StartingRow To iRow
                    iRowsStretchRatio(R) = iSng
                Next R
            End If
            mRows = iRow + 1
            iARPSTmp = (mRows - 1) * iRowPerspectiveSpace
            If Not iShowsRowsPerspective Then
                iAllRowsPerspectiveSpace = iARPSTmp
                Exit Do
            End If
        Loop Until iARPSTmp = iAllRowsPerspectiveSpace ' until it did not add another row
        
        ' step 2: set in what row goes each tab
        iDecreaseStretchRatio = False
        iIncreaseStretchRatio = False
        Do
            iRowTabCount = 0
            iAccumulatedTabWith = 0
            iRow = 0
            ctv = 0
            If iDecreaseStretchRatio Then
                For R = 0 To mRows - 1
                    iRowsStretchRatio(R) = iRowsStretchRatio(R) * 0.95
                    If iRowsStretchRatio(R) < 1 Then iRowsStretchRatio(R) = 1
                Next R
                iDecreaseStretchRatio = False
            ElseIf iIncreaseStretchRatio Then
                For R = 0 To mRows - 1
                    iRowsStretchRatio(R) = iRowsStretchRatio(R) * 1.05
                Next R
                iIncreaseStretchRatio = False
            End If
            iLastVisibleTab_Prev = -1
            For t = 0 To mTabs - 1
                If mTabData(t).Visible Then
                    mTabData(t).TopTab = False
                    ctv = ctv + 1
                    iLastVisibleTab_Prev = iLastVisibleTab
                    iLastVisibleTab = t
                    mTabData(t).LeftTab = False
                    mTabData(t).RightTab = False
                    If ctv = iVisibleTabs Then
                        mTabData(t).RightTab = True
                    End If
                    iLng = mTabData(t).PicAndCaptionWidth * iRowsStretchRatio(iRow) + 10
                    If iAccumulatedTabWith + iLng > (iAvailableSpaceForTabs + mTabData(t).PicAndCaptionWidth * 0.38) Then ' 0.38 is an add-hoc value, the right thing to do would be to make another step and recalculate everything several times changing the stretch ratio until an equilibrium point is found (or something like that). But with a couple of examples it seems too work acceptable with this value of 0.38. If there are too many tabs or too few tabs in the top row, here is the problem (probably).
                        If iRowTabCount = 0 Then ' this only tab alone passes the available space in the row (and it is the first one or all the previous tabs also entered here)
                            mTabData(t).Row = iRow
                            mTabData(t).PosH = 1
                            mTabData(t).LeftTab = True
                            mTabData(t).RightTab = True
                            If (iRow + 1) < mRows Then
                                iRow = iRow + 1
                            End If
                            iRowTabCount = 0
                            iAccumulatedTabWith = 0
                        Else
                            If (iRow + 1) < mRows Then
                                If iLastVisibleTab_Prev <> t Then
                                    mTabData(iLastVisibleTab_Prev).RightTab = True
                                End If
                                iRow = iRow + 1
                                iRowTabCount = 1
                                iAccumulatedTabWith = iLng + mTabSeparation2
                            Else
                                iRowTabCount = iRowTabCount + 1
                            End If
                            mTabData(t).PosH = iRowTabCount
                            If iRowTabCount = 1 Then
                                mTabData(t).LeftTab = True
                            End If
                            mTabData(t).Row = iRow
                        End If
                    Else
                        iAccumulatedTabWith = iAccumulatedTabWith + iLng + mTabSeparation2
                        iRowTabCount = iRowTabCount + 1
                        mTabData(t).PosH = iRowTabCount
                        If iRowTabCount = 1 Then
                            mTabData(t).LeftTab = True
                        End If
                        mTabData(t).Row = iRow
                    End If
                Else
                    mTabData(t).Row = -1
                End If
            Next t
            mTabData(iLastVisibleTab).PosH = iRowTabCount
            If iRowTabCount = 1 Then
                mTabData(iLastVisibleTab).LeftTab = True
            End If
            mTabData(iLastVisibleTab).RightTab = True
            
            If mRows = 1 Then
                iTabWidthStyle = ssTWSNonJustified
            End If
            If iTabWidthStyle = ssTWSJustified Then
                ' step 3: set the widths of the tabs for each row
                For iRow = 0 To mRows - 1
                    iAccumulatedTabWith = 0
                    iAccumulatedAdditionalFixedTabSpace = 0
                    iRowTabCount = 0
                    For t = 0 To mTabs - 1
                        If mTabData(t).Row = iRow Then
                            iAccumulatedTabWith = iAccumulatedTabWith + mTabData(t).PicAndCaptionWidth
                            iAccumulatedAdditionalFixedTabSpace = iAccumulatedAdditionalFixedTabSpace + 10 + mTabSeparation2
                            iRowTabCount = iRowTabCount + 1
                        End If
                    Next t
                    If iRowTabCount > 1 Then
                        iAccumulatedAdditionalFixedTabSpace = iAccumulatedAdditionalFixedTabSpace - mTabSeparation2
                        iSng = (iAvailableSpaceForTabs - iAccumulatedAdditionalFixedTabSpace) / iAccumulatedTabWith
                        If iSng < 1 Then
                            If Not iDoNotDecreaseStretchRatio Then
                                iDecreaseStretchRatio = True
                                Exit For
                            End If
                        End If
                    Else
                        If iAccumulatedTabWith = 0 Then
                            iSng = 1
                        Else
                            iSng = (iAvailableSpaceForTabs - iAccumulatedAdditionalFixedTabSpace) / iAccumulatedTabWith
                        End If
                    End If
                    For t = 0 To mTabs - 1
                        If mTabData(t).Row = iRow Then
                            mTabData(t).Width = mTabData(t).PicAndCaptionWidth * iSng
                        End If
                    Next t
                Next iRow
            End If
            
            For iRow = mRows - 1 To 1 Step -1
                iLng = 0
                For t = 0 To mTabs - 1
                    If mTabData(t).Row = iRow Then
                        iLng = iLng + 1
                        Exit For
                    End If
                Next t
                If iLng = 0 Then
                    iIncreaseStretchRatio = True
                End If
            Next iRow
        Loop While (iDecreaseStretchRatio Or iIncreaseStretchRatio) And (iTabWidthStyle = ssTWSJustified)
    End If
    
    If mRows = 1 Then
        If iTabSelExtraHeight > 0 Then
            mTabBodyStart = iTabHeight + iTabSelExtraHeight + 2
        Else
            mTabBodyStart = iTabHeight + 2
        End If
    Else
        mTabBodyStart = mRows * iTabHeight + 2
    End If
    mTabBodyHeight = iScaleHeight - mTabBodyStart + 2 '+ 1
    
    If mRows > 1 Then
        iAllRowsPerspectiveSpace = iRowPerspectiveSpace * (mRows - 1)
    End If
    mTabBodyWidth = iScaleWidth - iAllRowsPerspectiveSpace '- 1
    If mControlIsThemed Then
        mTabBodyWidth = mTabBodyWidth + mThemedTabBodyRightShadowPixels - 2
    End If
    
    If iTabWidthStyle = ssTWSNonJustified Then
        iAvailableSpaceForTabs = (iScaleWidth - iAllRowsPerspectiveSpace - IIf(mAppearanceIsPP, 4, 0))
        For iRow = 0 To mRows - 1
            iAccumulatedTabWith = 0
            iAccumulatedAdditionalFixedTabSpace = 0
            For t = 0 To mTabs - 1
                If mTabData(t).Row = iRow Then
                    iAccumulatedTabWith = iAccumulatedTabWith + mTabData(t).PicAndCaptionWidth
                    iAccumulatedAdditionalFixedTabSpace = iAccumulatedAdditionalFixedTabSpace + 10
                    If Not mTabData(t).RightTab Then
                        iAccumulatedAdditionalFixedTabSpace = iAccumulatedAdditionalFixedTabSpace + mTabSeparation2
                    End If
                End If
            Next t
            If mAmbientUserMode Then
                mMinSpaceNeeded = (iScaleWidth - iAvailableSpaceForTabs) + iAccumulatedTabWith + iAccumulatedAdditionalFixedTabSpace
                If iAccumulatedTabWith + iAccumulatedAdditionalFixedTabSpace > iAvailableSpaceForTabs Then
                    iSng = (iAvailableSpaceForTabs - iAccumulatedAdditionalFixedTabSpace) / iAccumulatedTabWith
                    For t = 0 To mTabs - 1
                        If mTabData(t).Row = iRow Then
                            mTabData(t).PicAndCaptionWidth = mTabData(t).PicAndCaptionWidth * iSng
                        End If
                    Next t
                End If
            End If
        Next iRow
    End If
    
    ' minimun size
    If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
        If mTabBodyHeight < 3 Then
            UserControl.Height = UserControl.Height + pScaleY(3 - mTabBodyHeight, vbPixels, vbTwips)
            GoTo TheExit:
        End If
        If iTabWidthStyle <> ssTWSJustified Then
            If UserControl.Width < mTabsPerRow * 500 + pScaleX(iAllRowsPerspectiveSpace, vbPixels, vbTwips) Then
                UserControl.Width = mTabsPerRow * 500 + pScaleX(iAllRowsPerspectiveSpace, vbPixels, vbTwips) + Screen_TwipsPerPixelX
                GoTo TheExit:
            End If
        End If
    Else
        If mTabBodyHeight < 3 Then
            iLng = UserControl.Width + pScaleX(3 - mTabBodyHeight, vbPixels, vbTwips)
            UserControl.Width = iLng
            GoTo TheExit:
        End If
        If iTabWidthStyle <> ssTWSJustified Then
            If UserControl.Height < mTabsPerRow * 500 + pScaleX(iAllRowsPerspectiveSpace, vbPixels, vbTwips) Then ' we are drawing horizontally, so ScaleX
                UserControl.Height = mTabsPerRow * 500 + pScaleX(iAllRowsPerspectiveSpace, vbPixels, vbTwips) + Screen_TwipsPerPixely
                GoTo TheExit:
            End If
        End If
    End If
    If (iTabMaxWidth > 0) And (iTabWidthStyle = ssTWSFixed) Then
        iLng = iTabMaxWidth * mTabsPerRow
        If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
            If pScaleX(iLng, vbPixels, vbTwips) > UserControl.Width Then
                UserControl.Width = pScaleX(iLng, vbPixels, vbTwips)
                GoTo TheExit:
            End If
        Else
            If pScaleY(iLng, vbPixels, vbTwips) > UserControl.Height Then
                UserControl.Height = pScaleY(iLng, vbPixels, vbTwips)
                GoTo TheExit:
            End If
        End If
        If mAppearanceIsPP Then
            iTabWidth = (iScaleWidth - 5 - iAllRowsPerspectiveSpace - 1 - IIf(mControlIsThemed, 2 - mThemedTabBodyRightShadowPixels, 0) - mTabSeparation2 * (mTabsPerRow - 1)) / mTabsPerRow
        Else
            iTabWidth = (iScaleWidth - 1 - iAllRowsPerspectiveSpace - 1 - IIf(mControlIsThemed, 2 - mThemedTabBodyRightShadowPixels, 0) - mTabSeparation2 * (mTabsPerRow - 1)) / mTabsPerRow
        End If
        If iTabWidth > iTabMaxWidth Then
            iTabWidth = iTabMaxWidth
        End If
    Else
        iTabWidth = (iScaleWidth - iAllRowsPerspectiveSpace - 1 - mTabSeparation2 * (mTabsPerRow - 1)) / mTabsPerRow
    End If
    
    If (mTabBodyWidth_Prev <> mTabBodyWidth) And (mTabBodyWidth_Prev <> 0) Or (mTabBodyHeight_Prev <> mTabBodyHeight) And (mTabBodyHeight_Prev <> 0) Then
        ResetCachedThemeImages
    End If
    mTabBodyWidth_Prev = mTabBodyWidth
    mTabBodyHeight_Prev = mTabBodyHeight
    
    If iTabWidthStyle <> ssTWSJustified Then
        iTabStretchRatio = 1
    End If
    
    ' Rows positions
    For t = 0 To mTabs - 1
        mTabData(t).RowPos = (mRows - mTabData(t).Row - 1) + mTabData(mTabSel).Row
        If mTabData(t).RowPos > (mRows - 1) Then mTabData(t).RowPos = mTabData(t).RowPos - mRows
    Next t
    
    ' set the tab rects
    For iRow = 0 To mRows - 1
        For t = 0 To mTabs - 1
            If mTabData(t).Visible Then
                If mTabData(t).RowPos = iRow Then
                    iTabData = mTabData(t)
                    With iTabData.TabRect
                        If t = mTabSel Then
                            If iTabSelExtraHeight > 0 Then
                                If mRows = 1 Then
                                    .Top = (mRows - 1) * iTabHeight
                                Else
                                    .Top = (mRows - 1) * iTabHeight - iTabSelExtraHeight
                                End If
                            Else
                                .Top = (mRows - 1) * iTabHeight
                            End If
                            .Bottom = .Top + iTabHeight + iTabSelExtraHeight
                        Else
                            If mRows = 1 Then
                                .Top = mTabData(t).RowPos * iTabHeight + iTabSelExtraHeight
                            Else
                                .Top = mTabData(t).RowPos * iTabHeight
                            End If
                            .Bottom = .Top + iTabHeight
                        End If
                        If (iTabWidthStyle = ssTWSFixed) Then
                            .Left = (iTabData.PosH - 1) * IIf(mControlIsThemed, iTabWidth, Round(iTabWidth)) + iRowPerspectiveSpace * (mRows - mTabData(t).RowPos - 1) + 1 + (iTabData.PosH - 1) * mTabSeparation2
                            If mAppearanceIsPP Then
                                .Left = .Left + 1
                            End If
                            .Right = .Left + iTabWidth - 1 '- mTabSeparation2 ' no volver a sacar el -1!!
                        Else
                            If iTabData.LeftTab Then
                                iTabLeft = 1 + iRowPerspectiveSpace * (mRows - mTabData(t).RowPos - 1) + 1
                            Else
                                iTabLeft = iTabLeft + mTabSeparation2
                            End If
                            .Left = iTabLeft
                            If iTabWidthStyle = ssTWSJustified Then
                                .Right = .Left + iTabData.Width + 9
                            Else
                                .Right = .Left + iTabData.PicAndCaptionWidth + 9
                            End If
                            iTabLeft = .Right + 1
                        End If
                        If iTabData.RightTab Then
                            iLng = iScaleWidth - iRowPerspectiveSpace * mTabData(t).RowPos - 1
                            If mAppearanceIsPP Then
                                iLng = iLng - 2
                                If mControlIsThemed Then
                                    If (iTabWidthStyle <> ssTWSJustified) Or iTabData.Selected Then
                                        iLng = iLng - 1
                                    End If
                                End If
                            End If
                            If t = mTabSel Then
                                If mControlIsThemed Then
                                    iLng = iLng + 1
                                End If
                            End If
                            If Abs(.Right - iLng) < 6 Then
                                .Right = iLng - IIf(mControlIsThemed, mThemedTabBodyRightShadowPixels - 2, 0)
                            End If
                        End If
                    End With
                    mTabData(t) = iTabData
                End If
            End If
        Next t
    Next iRow
    
    For t = 0 To mTabs - 1
        If mTabData(t).Visible Then
            If mTabData(t).PosH > 1 Then
                If mTabData(t).TabRect.Left <= mTabData(t - 1).TabRect.Right Then
                    iLng = t - 1
                    Do Until mTabData(iLng).Visible = True
                        iLng = iLng - 1
                        If iLng < 0 Then Exit Do
                    Loop
                    If iLng >= 0 Then
                        mTabData(t).TabRect.Left = mTabData(iLng).TabRect.Right + 1
                    End If
                End If
            End If
        End If
    Next t
    
    iLng = 0
    For iRow = 0 To mRows - 1
        For t = 0 To mTabs - 1
            If mTabData(t).RowPos = iRow Then
                If mTabData(t).TabRect.Left > (iLng - 2) Then
                    mTabData(t).TopTab = True
                End If
                If mTabData(t).RightTab Then
                    iLng = mTabData(t).TabRect.Right
                End If
            End If
        Next t
    Next iRow
    
    If Not mRedraw Then Exit Sub
    
    ' Do the draw
    
    ' How the "light" need to come according to TabOrientation (because the image later will be rotated). Note: in Windows the llight comes from top-left, and shadows are in bottom right.
    ' Top: from top-left
    ' Left: from top-right
    ' Right: from bottom-left
    ' Bottom: from bottom-left

    ' Do the draw
    picDraw.Width = iScaleWidth
    picDraw.Height = iScaleHeight
    
    If picDraw.BackColor <> mTabBackColor Then
        picDraw.BackColor = mTabBackColor ' the pic backcolor determines the focusrect color
    End If
    picDraw.Cls
    
    ' BackColor
    picDraw.Line (0, 0)-(iScaleWidth, iScaleHeight), mBackColor, BF
    
    ' shadow is at the bottom and all need to be shifted
    If (mTabOrientation = ssTabOrientationLeft) And mControlIsThemed Then
        For t = 0 To mTabs - 1
            mTabData(t).TabRect.Left = mTabData(t).TabRect.Left + mThemedTabBodyRightShadowPixels
            mTabData(t).TabRect.Right = mTabData(t).TabRect.Right + mThemedTabBodyRightShadowPixels
        Next t
    End If
    
    ' draw inactive tabs
    For iRow = 0 To mRows - 1
        For t = 0 To mTabs - 1
            If mTabData(t).Visible Then
                If mTabData(t).RowPos = iRow Then
                    If t <> mTabSel Then
                        If mTabData(t).RightTab And Not (mTabData(t).RowPos = mRows - 1) Then
                            iLng = 4
                            If mAppearanceIsPP Then
                                iLng = iLng + 2 + IIf(mControlIsThemed, mThemedTabBodyRightShadowPixels - 2, 0)
                            End If
                            If (iTabWidthStyle <> ssTWSJustified) Or iShowsRowsPerspective Then
                                DrawInactiveTabBodyPart iRowPerspectiveSpace * (mRows - mTabData(t).RowPos - 1) + 3, mTabData(t).TabRect.Bottom + 5, mTabBodyWidth - iLng, CLng(mTabBodyHeight), iLng, 1
                            End If
                        End If
                        If mAppearanceIsPP Then
                            mTabData(t).TabRect.Top = mTabData(t).TabRect.Top + 2
                        End If
                        DrawTab t
                        DrawTabPicureAndCaption t
                    End If
                End If
            End If
        Next t
    Next iRow
    
    ' Draw body
    DrawBody iScaleHeight
    
    ' Draw active tab
    If mAppearanceIsPP Then
        mTabData(mTabSel).TabRect.Left = mTabData(mTabSel).TabRect.Left - 2
        mTabData(mTabSel).TabRect.Right = mTabData(mTabSel).TabRect.Right + 2
    End If
    DrawTab CLng(mTabSel)
    DrawTabPicureAndCaption CLng(mTabSel)
    
    mEndOfTabs = 0
    For t = 0 To mTabs - 1
        If mTabData(t).Visible Then
            If mTabData(t).TabRect.Right > mEndOfTabs Then
                mEndOfTabs = mTabData(t).TabRect.Right
            End If
        End If
    Next t
    mEndOfTabs = mEndOfTabs + 1
    
    Select Case mTabOrientation
        Case ssTabOrientationTop
            mTabBodyRect.Top = mTabBodyStart
            mTabBodyRect.Left = 2
            mTabBodyRect.Bottom = mScaleHeight - 4
            mTabBodyRect.Right = mTabBodyWidth - 4
        Case ssTabOrientationBottom
            mTabBodyRect.Top = 2
            mTabBodyRect.Left = 2
            mTabBodyRect.Bottom = mTabBodyHeight - 4
            mTabBodyRect.Right = mTabBodyWidth - 4
        Case ssTabOrientationLeft
            mTabBodyRect.Top = mScaleHeight - mTabBodyWidth + 2
            mTabBodyRect.Left = mTabBodyStart
            mTabBodyRect.Bottom = mScaleHeight - 4
            mTabBodyRect.Right = mScaleWidth - 4
        Case Else ' ssTabOrientationRight
            mTabBodyRect.Top = 2
            mTabBodyRect.Left = 2
            mTabBodyRect.Bottom = mTabBodyWidth - 4
            mTabBodyRect.Right = mTabBodyHeight - 4
    End Select
    
    iAlreadyNeedToBePainted = GetUpdateRect(mUserControlHwnd, iTmpRect, 0&) <> 0&
    
    Select Case mTabOrientation
        Case ssTabOrientationTop
            'BitBlt UserControl.hDC, 0, 0, iScaleWidth, iScaleHeight, picDraw.hDC, 0, 0, vbSrcCopy
            Set UserControl.Picture = picDraw.Image
        Case ssTabOrientationBottom
            UserControl.PaintPicture picDraw.Image, 0, iScaleHeight - 1, iScaleWidth, -iScaleHeight
            Set UserControl.Picture = UserControl.Image
            UserControl.Cls
        Case ssTabOrientationLeft
            RotatePic picDraw, picRotate, efn90DegreesCounterClockWise
            'BitBlt UserControl.hDc, 0, 0, mScaleWidth, mScaleHeight, picRotate.hDc, 0, 0, vbSrcCopy
            Set UserControl.Picture = picRotate.Image
        Case Else ' ssTabOrientationRight
            RotatePic picDraw, picRotate, efn90DegreesClockWise
            'BitBlt UserControl.hDc, 0, 0, mScaleWidth, mScaleHeight, picRotate.hDc, 0, 0, vbSrcCopy
            Set UserControl.Picture = picRotate.Image
    End Select
    picDraw.Cls
    
    ' to avoid flickering on windowless contained controls, if not changed, validate the tab body area
    If (Not mTabBodyReset) Then
        If Not iAlreadyNeedToBePainted Then
            GetClientRect mUserControlHwnd, iTmpRect
            If mTabOrientation = ssTabOrientationTop Then
                iTmpRect.Top = mTabBodyStart + 3
            ElseIf mTabOrientation = ssTabOrientationBottom Then
                iTmpRect.Bottom = iTmpRect.Bottom - mTabBodyStart - 3
            ElseIf mTabOrientation = ssTabOrientationLeft Then
                iTmpRect.Left = mTabBodyStart + 3
            ElseIf mTabOrientation = ssTabOrientationRight Then
                iTmpRect.Right = iTmpRect.Right - mTabBodyStart - 3
            End If
            ValidateRect mUserControlHwnd, iTmpRect
        End If
    End If
    mTabBodyReset = False
    
    ' rotate caption RECTs according to TabOrientation
    If mTabOrientation = ssTabOrientationBottom Then
        For t = 0 To mTabs - 1
            iTabData = mTabData(t)
            If iTabData.Visible Then
                With iTabData.TabRect
                    iLng = .Top - 2
                    .Top = iScaleHeight - 3 - .Bottom
                    .Bottom = iScaleHeight - 3 - iLng
                End With
            End If
            mTabData(t) = iTabData
        Next t
    ElseIf mTabOrientation = ssTabOrientationLeft Then
        For t = 0 To mTabs - 1
            iTabData = mTabData(t)
            If iTabData.Visible Then
                With iTabData.TabRect
                    iTmpRect.Top = .Top
                    iTmpRect.Bottom = .Bottom
                    iTmpRect.Left = .Left
                    iTmpRect.Right = .Right
                    .Top = iScaleWidth - iTmpRect.Right
                    .Bottom = .Top + iTmpRect.Right - iTmpRect.Left
                    .Left = iTmpRect.Top
                    .Right = .Left + iTmpRect.Bottom - iTmpRect.Top
                End With
            End If
            mTabData(t) = iTabData
        Next t
    ElseIf mTabOrientation = ssTabOrientationRight Then
        For t = 0 To mTabs - 1
            iTabData = mTabData(t)
            If iTabData.Visible Then
                With iTabData.TabRect
                    iTmpRect.Top = .Top
                    iTmpRect.Bottom = .Bottom
                    iTmpRect.Left = .Left
                    iTmpRect.Right = .Right
                    .Top = iTmpRect.Left
                    .Bottom = .Top + iTmpRect.Right - iTmpRect.Left
                    .Left = iScaleHeight - iTmpRect.Bottom
                    .Right = .Left + iTmpRect.Bottom - iTmpRect.Top
                End With
            End If
            mTabData(t) = iTabData
        Next t
    End If
    If mRows <> mRows_Prev Then
        RaiseEvent RowsChange
    End If
    mRows_Prev = mRows
    If ((mTabBodyStart <> mTabBodyStart_Prev) And (mAutoRelocateControls = ssRelocateAlways) Or (mTabOrientation <> mTabOrientation_Prev) And (mAutoRelocateControls > 0)) And (mTabOrientation_Prev <> -1) Then
        RearrangeContainedControlsPositions
    End If
    mTabBodyStart_Prev = mTabBodyStart
    mTabOrientation_Prev = mTabOrientation
    
    If mBackStyle = ssOpaque Then
        Set UserControl.MaskPicture = Nothing
        tmrCheckContainedControlsAdditionDesignTime.Enabled = False
        tmrCheckContainedControlsAdditionDesignTime.Interval = 1
    Else
        tmrCheckContainedControlsAdditionDesignTime.Interval = 50
        tmrCheckContainedControlsAdditionDesignTime.Enabled = True
        picAux.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        Set picAux.Picture = UserControl.Picture
        
        Dim iCtl As Control
        Dim iLeft As Long
        Dim iWidth As Long
        
        On Error Resume Next
        For Each iCtl In UserControl.ContainedControls
            iLeft = -mLeftShiftToHide
            iLeft = iCtl.Left
            If iLeft > -mLeftShiftToHide Then
                iWidth = -1
                iWidth = iCtl.Width
                If iWidth <> -1 Then
                    picAux.Line (ScaleX(iLeft, vbTwips, vbPixels), ScaleY(iCtl.Top, vbTwips, vbPixels))-(ScaleX(iLeft + iWidth, vbTwips, vbPixels), ScaleY(iCtl.Top + iCtl.Height, vbTwips, vbPixels)), mTabSelBackColor2 Xor &H1, BF
                End If
            End If
        Next
        On Error GoTo 0
        If Not Ambient.UserMode Then mLastContainedControlsPositionsStr = GetContainedControlsPositionsStr
        
        Set UserControl.MaskPicture = picAux.Image
        
        Set picAux.Picture = Nothing
        picAux.Cls
    End If
    
    If (mTabBodyRect_Prev.Left <> mTabBodyRect.Left) Or (mTabBodyRect_Prev.Top <> mTabBodyRect.Top) Or (mTabBodyRect_Prev.Right <> mTabBodyRect.Right) Or (mTabBodyRect_Prev.Bottom <> mTabBodyRect.Bottom) Then
        RaiseEvent TabBodyResize
    End If
    
    mTabBodyRect_Prev.Left = mTabBodyRect.Left
    mTabBodyRect_Prev.Top = mTabBodyRect.Top
    mTabBodyRect_Prev.Right = mTabBodyRect.Right
    mTabBodyRect_Prev.Bottom = mTabBodyRect.Bottom
    
    If mSubclassControlsPaintingPending Then SubclassControlsPainting
    
TheExit:
    UserControl.ScaleMode = vbTwips
    If mTheme <> 0 Then
        CloseThemeData mTheme
        mTheme = 0
    End If
    mDrawing = False
End Sub

Private Sub DrawTab(nTab As Long)
    Dim iCurv As Long
    Dim iLeftShift As Long
    Dim iRightShift As Long
    Dim iTopShift As Long
    Dim iBottomShift As Long
    Dim iHighlighted As Boolean
    Dim iTabData As T_TabData
    Dim iExtI As Long
    Dim iActive As Boolean
    Dim iState As Long
    Dim iTRect As RECT
    Dim iTRect2 As RECT
    Dim iPartId As Long
    Dim iLeft As Long
    Dim iTop As Long
    Dim iRoundedTabs As Boolean
    Dim iTabBackColor2 As Long
    
    Dim i3DShadow As Long
    Dim i3DDKShadow As Long
    Dim i3DHighlight As Long
    Dim i3DHighlightH As Long
    Dim i3DHighlightV As Long
    Dim i3DShadowV As Long
    Dim iGlowColor As Long
    
    iTabData = mTabData(nTab)
    iActive = iTabData.Selected
    iRoundedTabs = (mTabAppearance2 = ssTAPropertyPageRounded) Or (mTabAppearance2 = ssTATabbedDialogRounded)
    If iActive Then
        iHighlighted = mTabSelHighlight And iTabData.Enabled
        iTabBackColor2 = mTabSelBackColor2
        i3DDKShadow = m3DDKShadow_Sel
        i3DHighlightH = m3DHighlightH_Sel
        i3DHighlightV = m3DHighlightV_Sel
        i3DShadowV = m3DShadowV_Sel
        i3DShadow = m3DShadow_Sel
        i3DHighlight = m3DHighlight_Sel
        iGlowColor = mGlowColor_Sel
    Else
        iHighlighted = (mTabHoverHighlight <> ssTHHNo) And iTabData.Hovered And (mEnabled Or (Not mAmbientUserMode)) And iTabData.Enabled
        iTabBackColor2 = mTabBackColor2
        i3DDKShadow = m3DDKShadow
        i3DHighlightH = m3DHighlightH
        i3DHighlightV = m3DHighlightV
        i3DShadowV = m3DShadowV
        i3DShadow = m3DShadow
        i3DHighlight = m3DHighlight
        iGlowColor = mGlowColor
    End If
    
    With iTabData.TabRect
        If mControlIsThemed Then
            If Not iTabData.Enabled Then
                iState = TIS_DISABLED
            ElseIf ((iActive And ControlHasFocus) And (Not mShowFocusRect) And mAmbientUserMode) Or iActive And ((mTabOrientation = ssTabOrientationBottom) Or (mTabOrientation = ssTabOrientationRight)) Then
                iState = TIS_SELECTED ' I had to put TIS_SELECTED instead of TIS_FOCUSED before
            ElseIf iActive Then
                iState = TIS_SELECTED
            ElseIf iHighlighted Then
                iState = TIS_HOT
            Else
                iState = TIS_NORMAL
            End If
            
            If mTabSeparation2 = 0 Then
                iPartId = IIf(iTabData.RightTab, TABP_TABITEMRIGHTEDGE, IIf(iTabData.LeftTab, TABP_TABITEMLEFTEDGE, TABP_TABITEM))
            Else
                iPartId = TABP_TABITEMLEFTEDGE
            End If
            If (mBackColor = vbButtonFace) And (Not (iTabData.RightTab Or (iState = TIS_FOCUSED)) Or (mTabSeparation2 > 0)) Then
                iTRect.Top = .Top
                iTRect.Left = .Left
                iTRect.Right = .Right + 1
                iTRect.Bottom = .Bottom + 1
                If Not iActive Then
                    If (mTabSeparation2 > 0) Then
                        If iTabData.RightTab Then
                            iTRect.Bottom = iTRect.Bottom + 1
                        End If
                    End If
                End If
                If mTabData(nTab).RowPos <> mRows - 1 Then
                    iTRect.Bottom = iTRect.Bottom + 4
                End If
                iTRect2 = iTRect
                iTRect2.Bottom = iTRect.Bottom + 1
                DrawThemeBackground mTheme, picDraw.hDC, iPartId, iState, iTRect2, iTRect
            Else
                iTRect.Left = 0
                iTRect.Top = 0
                iTRect.Bottom = .Bottom - .Top
                iTRect.Bottom = iTRect.Bottom + 1
                If (mTabOrientation = ssTabOrientationBottom) Or (mTabOrientation = ssTabOrientationRight) Then
                    iTRect.Bottom = iTRect.Bottom + 1
                End If
                If Not iActive Then
                    If iTabData.RightTab Then
                        iTRect.Bottom = iTRect.Bottom + 1
                    End If
                End If
                If mTabData(nTab).RowPos <> mRows - 1 Then
                    iTRect.Bottom = iTRect.Bottom + 4
                End If
                iTRect.Right = .Right - .Left + 1
                iLeft = .Left
                iTop = .Top
                On Error Resume Next
                picAux.Width = iTRect.Right
                picAux.Height = iTRect.Bottom
                picAux.Cls
                
                iTRect2 = iTRect
                iTRect2.Bottom = iTRect.Bottom + 1
                DrawThemeBackground mTheme, picAux.hDC, iPartId, iState, iTRect2, iTRect
'                SetThemedTabTransparentPixels iTabData.LeftTab, (iState = TIS_FOCUSED) Or (iTabData.RightTab And Not iState = TIS_SELECTED), (iTabData.TopTab Or (mTabSeparation2 > 0)) And Not (iState = TIS_SELECTED)
                SetThemedTabTransparentPixels iTabData.LeftTab, (iState = TIS_FOCUSED) Or iTabData.RightTab, (iTabData.TopTab Or (mTabSeparation2 > 0)) And Not (iState = TIS_SELECTED)
                Call TransparentBlt(picDraw.hDC, iLeft, iTop, iTRect.Right, iTRect.Bottom, picAux.hDC, 0, 0, iTRect.Right, iTRect.Bottom, cAuxTransparentColor)
                picAux.Cls
                On Error GoTo 0
            End If
        Else
            If iActive Then
                ' active tab background
                If mAppearanceIsPP Then
'                    If (mTabOrientation = ssTabOrientationBottom) Or (mTabOrientation = ssTabOrientationRight) Then
                    iExtI = 2
'                    Else
'                        iExtI = 2
'                    End If
                    iLeftShift = 1
                    iRightShift = 1
                    iTopShift = 0
                    iBottomShift = 1
                    If iRoundedTabs Then
                        iCurv = 3
                    Else
                        iCurv = 2
                    End If
                Else
                    iExtI = 1
                    If iRoundedTabs Then
                        iLeftShift = -1
                    Else
                        iLeftShift = 0
                    End If
                    iRightShift = 0
                    iTopShift = 0
                    iBottomShift = 2
                    iCurv = 4
                End If
            Else
                ' inactive tab background
                iExtI = 6
                If mAppearanceIsPP Then
                    iLeftShift = 0
                    iRightShift = 0
                    iTopShift = 0 '2
                    iBottomShift = 5
                    If iRoundedTabs Then
                        iCurv = 3
                    Else
                        iCurv = 2
                    End If
                Else
                    If iRoundedTabs Then
                        iLeftShift = -1
                    Else
                        iLeftShift = 0
                    End If
                    iTopShift = 0
                    iRightShift = -1
                    iBottomShift = 6
'                    If iTabData.RightTab Then
 '                       iBottomShift = iBottomShift + 1
  '                  End If
                    iCurv = 3
                End If
            End If
            
            If iHighlighted And mAmbientUserMode Then
'                Call FillCurvedGradient(.Left, .Top + iTopShift, .Right + iRightShift, .Bottom + iBottomShift, iTabBackColor2, iGlowColor, iCurv, True, True)
                Call FillCurvedGradient(.Left, .Top + iTopShift, .Right + iRightShift, (.Bottom + iBottomShift + .Top + iTopShift) / 2 + 2, iTabBackColor2, iGlowColor, iCurv, True, True)
                Call FillCurvedGradient(.Left, (.Bottom + iBottomShift + .Top + iTopShift) / 2, .Right + iRightShift, .Bottom + iBottomShift, iGlowColor, iTabBackColor2, iCurv, True, True)
                
                'Call FillCurvedGradient(.Left, .Bottom - (.Bottom - .Top - iBottomShift) / 2 + iTopShift, .Right + iRightShift, .Bottom - iBottomShift, IIf(iTabData.Selected, iGlowColor, iGlowColor), iTabBackColor2, 0, False, False)
            Else
                Call FillCurvedGradient(.Left, .Top + iTopShift, .Right + iRightShift, .Bottom + iBottomShift, iTabBackColor2, iTabBackColor2, iCurv, True, True)
            End If
            
            'top line
            If Not mAppearanceIsPP Then
                If iRoundedTabs Then
                    picDraw.Line (.Left + iLeftShift + 4, .Top)-(.Right - 3, .Top), i3DDKShadow
                    picDraw.Line (.Left + iLeftShift + 4, .Top + 1)-(.Right - 4, .Top + 1), i3DHighlightH
                    If iActive Then
                        picDraw.Line (.Left + iLeftShift + 4, .Top + 2)-(.Right - 4, .Top + 2), i3DHighlightH
                    End If
                Else
                    picDraw.Line (.Left + iLeftShift + 3, .Top)-(.Right - 3, .Top), i3DDKShadow
                    picDraw.Line (.Left + iLeftShift + 3, .Top + 1)-(.Right - 3, .Top + 1), i3DHighlightH
                    If iActive Then
                        picDraw.Line (.Left + iLeftShift + 3, .Top + 2)-(.Right - 3, .Top + 2), i3DHighlightH
                    End If
                End If
            Else
                If iRoundedTabs Then
                    picDraw.Line (.Left + iLeftShift + 2, .Top)-(.Right - 2, .Top), i3DHighlightH
                Else
                    picDraw.Line (.Left + iLeftShift + 2, .Top)-(.Right - 1, .Top), i3DHighlightH
                End If
                If (mTabOrientation = ssTabOrientationBottom) Or (mTabOrientation = ssTabOrientationRight) Then
                    If iRoundedTabs Then
                        picDraw.Line (.Left + iLeftShift + 4, .Top - 1)-(.Right - 3, .Top - 1), i3DHighlightH
                    Else
                        picDraw.Line (.Left + iLeftShift + 3, .Top - 1)-(.Right - 2, .Top - 1), i3DHighlightH
                    End If
                End If
            End If
            
            'right line
            If Not mAppearanceIsPP Then
                picDraw.Line (.Right, .Top + 4)-(.Right, .Bottom + iExtI), i3DDKShadow
                picDraw.Line (.Right - 1, .Top + 4)-(.Right - 1, .Bottom + iExtI), i3DShadowV
                If iActive Then
                    picDraw.Line (.Right - 2, .Top + 4)-(.Right - 2, .Bottom + 1 + iExtI), i3DShadowV
                    If iTabData.RightTab Then
                        picDraw.Line (.Right - 1, .Bottom + iExtI)-(.Right - 1, .Bottom + iExtI + 2), i3DShadowV ' points of top line of body
                        picDraw.Line (.Right - 2, .Bottom + iExtI + 1)-(.Right - 2, .Bottom + iExtI + 2), i3DShadowV ' point of top line of body
                    Else
                        picDraw.Line (.Right - 1, .Bottom + iExtI)-(.Right - 1, .Bottom + iExtI + 2), i3DHighlightH  ' points of top line of body
                        picDraw.Line (.Right - 2, .Bottom + iExtI + 1)-(.Right - 2, .Bottom + iExtI + 2), i3DHighlightH  ' point of top line of body
                    End If
                End If
            Else
                If mTabOrientation = ssTabOrientationTop Then
                    picDraw.Line (.Right, .Top + 3)-(.Right, .Bottom + iExtI), i3DDKShadow
                    picDraw.Line (.Right - 1, .Top + 3)-(.Right - 1, .Bottom + iExtI), i3DShadowV
                ElseIf mTabOrientation = ssTabOrientationLeft Then
                    picDraw.Line (.Right, .Top + 3)-(.Right, .Bottom + iExtI), i3DHighlightH
                    picDraw.Line (.Right - 1, .Top + 3)-(.Right - 1, .Bottom + iExtI), iTabBackColor2
                Else
                    picDraw.Line (.Right, .Top + 3)-(.Right, .Bottom + iExtI - 1), i3DDKShadow
                    picDraw.Line (.Right, .Bottom + iExtI - 1)-(.Right + 1, .Bottom + iExtI - 1), i3DShadowV
                    picDraw.Line (.Right - 1, .Top + 3)-(.Right - 1, .Bottom + iExtI), i3DShadowV
                End If
            End If
            
            'left line
            If Not mAppearanceIsPP Then
                If iRoundedTabs Then
                    If iTabData.LeftTab Then
                        picDraw.Line (.Left, .Top + 5)-(.Left, .Bottom + iExtI), i3DDKShadow
                    End If
                    If mTabOrientation = ssTabOrientationLeft Then
                        
                        If iActive Then
                            picDraw.Line (.Left, .Top + 5)-(.Left, .Bottom + iExtI + 1), i3DHighlightV
                            picDraw.Line (.Left + 1, .Top + 5)-(.Left + 1, .Bottom + 2 + iExtI), i3DHighlightV
                        Else
                            picDraw.Line (.Left, .Top + 5)-(.Left, .Bottom + iExtI), i3DHighlightV
                            picDraw.Line (.Left + 1, .Top + 5)-(.Left + 1, .Bottom + iExtI), iTabBackColor2
                        End If
                    Else
                        If iActive Then
                            picDraw.Line (.Left, .Top + 5)-(.Left, .Bottom + iExtI + 1), i3DHighlightV
                            picDraw.Line (.Left + 1, .Top + 5)-(.Left + 1, .Bottom + 2 + iExtI), i3DHighlightV
                        Else
                            picDraw.Line (.Left, .Top + 5)-(.Left, .Bottom + iExtI), i3DHighlightV ' iTabBackColor2
                            'picDraw.Line (.Left + 1, .Top + 5)-(.Left + 1, .Bottom + iExtI), i3DHighlightV
                        End If
                        
                    End If
                Else
                    picDraw.Line (.Left, .Top + 4)-(.Left, .Bottom + iExtI), i3DHighlightV
                    If iActive Then
                        picDraw.Line (.Left + 1, .Top + 4)-(.Left + 1, .Bottom + 1 + iExtI), i3DHighlightV
                        picDraw.Line (.Left, .Bottom + iExtI)-(.Left, .Bottom + iExtI + 2), i3DHighlightV   ' points of top line of body
                        picDraw.Line (.Left + 1, .Bottom + iExtI + 1)-(.Left + 1, .Bottom + iExtI + 2), i3DHighlightV ' point of top line of body
                    End If
                End If
                picDraw.Line (.Left - 1, .Top + 5)-(.Left - 1, .Bottom + iExtI), i3DDKShadow
            Else
                If mTabOrientation <> ssTabOrientationLeft Then
                    If iRoundedTabs Then
                        picDraw.Line (.Left, .Top + 3)-(.Left, .Bottom + iExtI), i3DHighlightV
                    Else
                        picDraw.Line (.Left, .Top + 2)-(.Left, .Bottom + iExtI), i3DHighlightV
                    End If
                Else
                    picDraw.Line (.Left, .Top + 2)-(.Left, .Bottom), i3DDKShadow
                    picDraw.Line (.Left + 1, .Top + 2)-(.Left + 1, .Bottom + iExtI), i3DShadow
                    If iRoundedTabs Then
                        picDraw.Line (.Left, .Top + 3)-(.Left, .Bottom + iExtI), i3DDKShadow
                        picDraw.Line (.Left + 1, .Top + 3)-(.Left + 1, .Bottom + iExtI), i3DShadow
                    Else
                        picDraw.Line (.Left, .Top + 2)-(.Left, .Bottom + iExtI), i3DDKShadow
                        picDraw.Line (.Left + 1, .Top + 2)-(.Left + 1, .Bottom + iExtI), i3DShadow
                    End If
                End If
            End If
            
            'top-right corner
            If Not mAppearanceIsPP Then
                If iRoundedTabs Then
                    picDraw.Line (.Right - 1, .Top + 4)-(.Right - 1, .Top + 1), i3DDKShadow
                    picDraw.Line (.Right - 2, .Top + 1)-(.Right - 4, .Top + 1), i3DDKShadow
                    picDraw.Line (.Right - 2, .Top + 2)-(.Right - 3, .Top + 2), i3DShadowV
                    picDraw.Line (.Right - 4, .Top + 1)-(.Right - 3, .Top + 1), i3DShadowV
                    picDraw.Line (.Right - 2, .Top + 3)-(.Right - 2, .Top + 4), i3DShadowV
                    If iActive Then
                        picDraw.Line (.Right - 3, .Top + 2)-(.Right, .Top + 5), i3DShadowV
                        picDraw.Line (.Right - 4, .Top + 2)-(.Right - 1, .Top + 5), i3DShadowV
                        picDraw.Line (.Right - 3, .Top + 4)-(.Right - 1, .Top + 6), i3DShadowV
                    End If
                Else
                    picDraw.Line (.Right - 4, .Top)-(.Right, .Top + 4), i3DDKShadow
                    If iActive Then
                        picDraw.Line (.Right - 3, .Top + 2)-(.Right, .Top + 5), i3DShadowV
                        picDraw.Line (.Right - 4, .Top + 2)-(.Right - 1, .Top + 5), i3DShadowV
                        picDraw.Line (.Right - 4, .Top + 3)-(.Right - 1, .Top + 6), i3DShadowV
                    Else
                        picDraw.Line (.Right - 4, .Top + 1)-(.Right - 1, .Top + 4), i3DShadowV
                    End If
                End If
            Else
                If mTabOrientation <> ssTabOrientationLeft Then
                    If iRoundedTabs Then
                        picDraw.Line (.Right - 2, .Top + 1)-(.Right - 2, .Top + 2), i3DShadowV
                        picDraw.Line (.Right - 1, .Top + 1)-(.Right - 1, .Top + 2), i3DShadowV
                        picDraw.Line (.Right - 1, .Top + 2)-(.Right - 1, .Top + 3), i3DDKShadow
                    Else
                        picDraw.Line (.Right - 1, .Top + 1)-(.Right - 1, .Top + 2), i3DDKShadow
                        picDraw.Line (.Right - 1, .Top + 2)-(.Right - 1, .Top + 3), i3DShadowV
                        picDraw.Line (.Right, .Top + 2)-(.Right, .Top + 3), i3DDKShadow
                    End If
                Else
                    If iRoundedTabs Then
                        picDraw.Line (.Right - 2, .Top + 1)-(.Right - 2, .Top + 2), i3DHighlight
                        picDraw.Line (.Right - 1, .Top + 1)-(.Right - 1, .Top + 2), i3DHighlight
                        picDraw.Line (.Right - 1, .Top + 2)-(.Right - 1, .Top + 3), i3DHighlight
                    Else
                        picDraw.Line (.Right - 1, .Top + 1)-(.Right - 1, .Top + 2), i3DHighlight
                        picDraw.Line (.Right - 1, .Top + 2)-(.Right - 1, .Top + 3), i3DHighlight
                        picDraw.Line (.Right, .Top + 2)-(.Right, .Top + 3), i3DHighlight
                    End If
                End If
            End If
            
            'top-left corner
            If Not mAppearanceIsPP Then
                If iRoundedTabs Then
                    picDraw.Line (.Left + iLeftShift + 1, .Top + 4)-(.Left + iLeftShift + 1, .Top + 1), i3DDKShadow
                    picDraw.Line (.Left + iLeftShift + 2, .Top + 1)-(.Left + iLeftShift + 4, .Top + 1), i3DDKShadow
                    picDraw.Line (.Left + iLeftShift + 2, .Top + 3)-(.Left + iLeftShift + 2, .Top + 2), i3DHighlightH
                    picDraw.Line (.Left + iLeftShift + 2, .Top + 2)-(.Left + iLeftShift + 4, .Top + 2), i3DHighlightH
                    picDraw.Line (.Left + iLeftShift, .Top + 4)-(.Left + iLeftShift, .Top + 3), i3DDKShadow
                    picDraw.Line (.Left + iLeftShift + 1, .Top + 4)-(.Left + iLeftShift + 1, .Top + 5), i3DHighlightH
                    If iActive Then
                        picDraw.Line (.Left + iLeftShift + 2, .Top + 3)-(.Left + iLeftShift + 4, .Top + 1), i3DHighlightH
                        picDraw.Line (.Left + iLeftShift + 2, .Top + 4)-(.Left + iLeftShift + 5, .Top + 1), i3DHighlightH
                        picDraw.Line (.Left + iLeftShift + 2, .Top + 5)-(.Left + iLeftShift + 5, .Top + 2), i3DHighlightH
                    End If
                Else
                    picDraw.Line (.Left + iLeftShift - 1, .Top + 4)-(.Left + iLeftShift + 3, .Top), i3DDKShadow
                    If iActive Then
                        picDraw.Line (.Left + iLeftShift + 1, .Top + 3)-(.Left + iLeftShift + 3, .Top + 1), i3DHighlightH
                        picDraw.Line (.Left + iLeftShift + 1, .Top + 4)-(.Left + iLeftShift + 4, .Top + 1), i3DHighlightH
                        picDraw.Line (.Left + iLeftShift + 2, .Top + 4)-(.Left + iLeftShift + 4, .Top + 2), i3DHighlightH
                    Else
                        picDraw.Line (.Left + iLeftShift, .Top + 4)-(.Left + iLeftShift + 3, .Top + 1), i3DHighlightH
                    End If
                End If
            Else
                If mTabOrientation <> ssTabOrientationLeft Then
                    If iRoundedTabs Then
                        picDraw.Line (.Left + 1, .Top + 2)-(.Left + 1, .Top + 3), i3DHighlightH
                        picDraw.Line (.Left, .Top + 3)-(.Left, .Top + 4), i3DHighlightH
                        picDraw.Line (.Left + 1, .Top + 1)-(.Left + 3, .Top + 1), i3DHighlightH
                    Else
                        picDraw.Line (.Left, .Top + 2)-(.Left + 3, .Top - 1), i3DHighlightH
                    End If
                Else
                    If iRoundedTabs Then
                        picDraw.Line (.Left + 1, .Top + 2)-(.Left + 1, .Top + 3), i3DHighlightV
                        picDraw.Line (.Left + 1, .Top + 1)-(.Left + 3, .Top + 1), i3DHighlightV
                    Else
                        picDraw.Line (.Left, .Top + 2)-(.Left + 3, .Top - 1), i3DHighlightV
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub DrawInactiveTabBodyPart(nLeft As Long, nTop As Long, ByVal nWidth As Long, nHeight As Long, nXShift As Long, nSectionID_ForTesting As Long)
    Dim iDoRightLine As Boolean
    Dim iDoBottomLine As Boolean
    Dim iTesting As Boolean
    Dim iTabBackColor As Long
    
    If (nWidth < 1) Or (nHeight < 1) Or (nXShift > mTabBodyWidth) Then Exit Sub
    
'    iTesting = True
    
    If iTesting Then
        Select Case nSectionID_ForTesting
            Case 1
                iTabBackColor = vbGreen
            Case 2
                iTabBackColor = vbMagenta
            Case 3
                iTabBackColor = vbBlue
            Case 4
                iTabBackColor = vbCyan
        End Select
    Else
        iTabBackColor = mTabBackColor2
    End If
    
    If mControlIsThemed Then
        If (nWidth > mThemedTabBodyRightShadowPixels) Then
            EnsureInactiveTabBodyThemedReady
            BitBlt picDraw.hDC, nLeft, nTop, nWidth, nHeight, picInactiveTabBodyThemed.hDC, nXShift, 0, vbSrcCopy
        End If
    Else
        iDoRightLine = mTabBodyWidth - (nWidth + nXShift) <= 0
        iDoBottomLine = mTabBodyHeight - nHeight <= 0
        'nWidth = nWidth - 1
        
        picDraw.Line (nLeft, nTop)-(nLeft + nWidth, nTop + nHeight), iTabBackColor, BF
        
        'top line
        If Not mAppearanceIsPP Then
            picDraw.Line (nLeft - 1, nTop)-(nLeft + nWidth, nTop), m3DDKShadow
            picDraw.Line (nLeft - 1, nTop + 1)-(nLeft + 1 + nWidth, nTop + 1), m3DHighlightH
        Else
            If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationLeft) Then
                picDraw.Line (nLeft - 1, nTop)-(nLeft + nWidth, nTop), m3DHighlight
            Else
                picDraw.Line (nLeft - 1, nTop)-(nLeft + nWidth, nTop), m3DDKShadow
                picDraw.Line (nLeft - 1, nTop + 1)-(nLeft + nWidth - 1, nTop + 1), m3DShadow
            End If
        End If
        
        'right line
        If iDoRightLine Then
            If (mTabOrientation <> ssTabOrientationLeft) Or (Not mAppearanceIsPP) Then
                picDraw.Line (nLeft + nWidth, nTop)-(nLeft + nWidth, nTop + nHeight), m3DDKShadow
                picDraw.Line (nLeft + nWidth - 1, nTop + 1)-(nLeft + nWidth - 1, nTop + nHeight), m3DShadowV
            Else
                picDraw.Line (nLeft + nWidth, nTop)-(nLeft + nWidth, nTop + nHeight), m3DHighlightH
            End If
        End If
        
        'bottom line
        If iDoBottomLine Then
            If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationLeft) Then
                picDraw.Line (nLeft - 1, nTop - 1 + nHeight)-(nLeft + nWidth, nTop - 1 + nHeight), m3DShadow
                picDraw.Line (nLeft - 1, nTop + nHeight)-(nLeft + nWidth + 1, nTop + nHeight), m3DDKShadow
            Else
                picDraw.Line (nLeft - 1, nTop - 1 + nHeight)-(nLeft + nWidth, nTop - 1 + nHeight), m3DHighlight
            End If
        End If
    End If
End Sub


Private Sub DrawBody(nScaleHeight As Long)
    Dim iLng As Long
    
    If mControlIsThemed Then
        EnsureTabBodyThemedReady
        BitBlt picDraw.hDC, 0, mTabBodyStart - 2, picTabBodyThemed.ScaleWidth, picTabBodyThemed.ScaleHeight, picTabBodyThemed.hDC, 0, 0, vbSrcCopy
    Else
        ' background
        If mAppearanceIsPP Then
            iLng = -1
        Else
            iLng = 1
        End If
        picDraw.Line (0, mTabBodyStart + iLng)-(mTabBodyWidth - 1, nScaleHeight - 1), mTabSelBackColor2, BF
        
        If Not mAppearanceIsPP Then
            ' top line
            picDraw.Line (0, mTabBodyStart - 2)-(mTabBodyWidth - 1, mTabBodyStart - 2), m3DDKShadow_Sel
            picDraw.Line (2, mTabBodyStart - 1)-(mTabBodyWidth - 1, mTabBodyStart - 1), m3DHighlightH_Sel
            picDraw.Line (3, mTabBodyStart)-(mTabBodyWidth - 2, mTabBodyStart), m3DHighlightH_Sel
            
            ' left line
            picDraw.Line (0, mTabBodyStart - 1)-(0, nScaleHeight - 1), m3DDKShadow_Sel
            picDraw.Line (1, mTabBodyStart - 1)-(1, nScaleHeight - 2), m3DHighlightV_Sel
            picDraw.Line (2, mTabBodyStart + 1)-(2, nScaleHeight - 3), m3DHighlightV_Sel

            ' right line
            picDraw.Line (mTabBodyWidth - 1, mTabBodyStart - 2)-(mTabBodyWidth - 1, nScaleHeight - 1), m3DDKShadow_Sel
            picDraw.Line (mTabBodyWidth - 2, mTabBodyStart - 1)-(mTabBodyWidth - 2, nScaleHeight - 2), m3DShadowV_Sel
            picDraw.Line (mTabBodyWidth - 3, mTabBodyStart)-(mTabBodyWidth - 3, nScaleHeight - 3), m3DShadowV_Sel
            
            ' bottom line
            picDraw.Line (0, nScaleHeight - 1)-(mTabBodyWidth, nScaleHeight - 1), m3DDKShadow_Sel
            If mTabBodyHeight > 3 Then
                picDraw.Line (1, nScaleHeight - 2)-(mTabBodyWidth - 1, nScaleHeight - 2), m3DShadowH_Sel
            End If
            If mTabBodyHeight > 4 Then
                picDraw.Line (2, nScaleHeight - 3)-(mTabBodyWidth - 2, nScaleHeight - 3), m3DShadowH_Sel
            End If
        
        Else
            ' top line
            If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationLeft) Then
                picDraw.Line (1, mTabBodyStart - 2)-(mTabBodyWidth - 1, mTabBodyStart - 2), m3DHighlightH_Sel
            Else
                picDraw.Line (0, mTabBodyStart - 2)-(mTabBodyWidth - 1, mTabBodyStart - 2), m3DDKShadow_Sel
                picDraw.Line (1, mTabBodyStart - 1)-(mTabBodyWidth - 1, mTabBodyStart - 1), m3DShadow_Sel
            End If
            
            If (mTabOrientation = ssTabOrientationTop) Then
                'left line
                picDraw.Line (0, mTabBodyStart - 1)-(0, nScaleHeight - 1), m3DHighlightV_Sel
                
                'right line
                picDraw.Line (mTabBodyWidth - 1, mTabBodyStart - 2)-(mTabBodyWidth - 1, nScaleHeight - 1), m3DDKShadow_Sel
                picDraw.Line (mTabBodyWidth - 2, mTabBodyStart - 1)-(mTabBodyWidth - 2, nScaleHeight - 2), m3DShadowV_Sel
                
                'bottom line
                picDraw.Line (0, nScaleHeight - 1)-(mTabBodyWidth, nScaleHeight - 1), m3DDKShadow_Sel
                If mTabBodyHeight > 3 Then
                    picDraw.Line (1, nScaleHeight - 2)-(mTabBodyWidth - 1, nScaleHeight - 2), m3DShadowH_Sel
                End If
            ElseIf (mTabOrientation = ssTabOrientationLeft) Then
                'left line
                picDraw.Line (0, mTabBodyStart - 1)-(0, nScaleHeight - 1), m3DDKShadow_Sel
                picDraw.Line (1, mTabBodyStart - 1)-(1, nScaleHeight - 1), m3DShadow_Sel
            
                'right line
                picDraw.Line (mTabBodyWidth - 1, mTabBodyStart - 2)-(mTabBodyWidth - 1, nScaleHeight - 1), m3DHighlight_Sel
            
                'bottom line
                picDraw.Line (0, nScaleHeight - 1)-(mTabBodyWidth, nScaleHeight - 1), m3DDKShadow_Sel
                If mTabBodyHeight > 3 Then
                    picDraw.Line (1, nScaleHeight - 2)-(mTabBodyWidth - 1, nScaleHeight - 2), m3DShadowH_Sel
                End If
            Else 'ssTabOrientationBottom OR ssTabOrientationRight
                'left line
                picDraw.Line (0, mTabBodyStart - 1)-(0, nScaleHeight - 1), m3DHighlightV_Sel
                
                'right line
                picDraw.Line (mTabBodyWidth - 1, mTabBodyStart - 2)-(mTabBodyWidth - 1, nScaleHeight), m3DDKShadow_Sel
                picDraw.Line (mTabBodyWidth - 2, mTabBodyStart - 1)-(mTabBodyWidth - 2, nScaleHeight - 1), m3DShadowV_Sel
                
                ' bottom line
                picDraw.Line (0, nScaleHeight - 1)-(mTabBodyWidth - 1, nScaleHeight - 1), m3DShadowH_Sel
            End If
        End If
    End If
End Sub

Private Sub DrawTabPicureAndCaption(t As Long)
    Dim iTabData As T_TabData
    Dim iTabSpaceRect As RECT
    Dim iCaptionRect As RECT
    Dim iMeasureRect As RECT
    Dim iFocusRect As RECT
    Dim iAuxPicture As StdPicture
    Dim iPicWidth As Long
    Dim iPicHeight As Long
    Dim iCaption As String
    Dim iAuxPic As PictureBox
    Dim iFontBoldPrev As Boolean
    Dim iFlags As Long
    Dim iPicLeft As Long
    Dim iPicTop As Long
    Dim iLng As Long
    Dim iPicSourceShiftX As Long
    Dim iPicSourceShiftY As Long
    Dim iTabSpaceWidth As Long
    Dim iTabSpaceHeight As Long
    Dim iMeasureWidth As Long
    Dim iMeasureHeight As Long
    Dim iPicWidthToShow As Long
    Dim iPicHeightToShow As Long
    Dim iTabPictureAlignment As vbExTabPictureAlignmentConstants
    Dim iTabBackColor2 As Long
    Dim iForeColor As Long
    Dim iGrayText As Long
    Dim iForeColor2 As Long
    
    If Not mTabData(t).Visible Then Exit Sub
    If Not mTabData(t).PicToUseSet Then SetPicToUse t
    
    iTabData = mTabData(t)
    
    If t = mTabSel Then
        iTabBackColor2 = mTabSelBackColor2
        iForeColor = mTabSelForeColor
        iGrayText = mGrayText_Sel
    Else
        iTabBackColor2 = mTabBackColor2
        iForeColor = mForeColor
        iGrayText = mGrayText
    End If
    
    If mTabOrientation = ssTabOrientationBottom Then
        If iTabData.Enabled And mEnabled Then
            picAux.ForeColor = iForeColor
        Else
            picAux.ForeColor = iGrayText
        End If
        iForeColor2 = picAux.ForeColor
        
        iFontBoldPrev = picAux.FontBold
        If t = mTabSel Then
            If mAppearanceIsPP And (mTabSelFontBold = ssYNAuto) Then
                picAux.FontBold = mFont.Bold
            ElseIf (mTabSelFontBold = ssYes) Or (mTabSelFontBold = ssYNAuto) Then
                picAux.FontBold = True
            Else
                picAux.FontBold = False
            End If
        Else
            picAux.FontBold = mFont.Bold
        End If
    Else
        If iTabData.Enabled And mEnabled Then
            picDraw.ForeColor = iForeColor
        Else
            picDraw.ForeColor = iGrayText
        End If
        iForeColor2 = picDraw.ForeColor
        
        iFontBoldPrev = picDraw.FontBold
        If t = mTabSel Then
            If mAppearanceIsPP And (mTabSelFontBold = ssYNAuto) Then
                picDraw.FontBold = mFont.Bold
            ElseIf (mTabSelFontBold = ssYes) Or (mTabSelFontBold = ssYNAuto) Then
                picDraw.FontBold = True
            Else
                picDraw.FontBold = False
            End If
        Else
            picDraw.FontBold = mFont.Bold
        End If
    End If
    
    iTabSpaceRect.Left = iTabData.TabRect.Left + 2
    iTabSpaceRect.Top = iTabData.TabRect.Top + 2
    iTabSpaceRect.Bottom = iTabData.TabRect.Bottom '- 2
    iTabSpaceRect.Right = iTabData.TabRect.Right - 2
    
    If mAppearanceIsPP And iTabData.Selected Then
        iTabSpaceRect.Top = iTabSpaceRect.Top - 1
    End If
    
    If Not iTabData.PicToUse Is Nothing Then
        If iTabData.Enabled And mEnabled Then
            Set iAuxPicture = iTabData.PicToUse
        Else
            If iTabData.PicToUse.Type = vbPicTypeBitmap Then
                If Not iTabData.PicDisabledSet Then
                    Set mTabData(t).PicDisabled = PictureToGrayScale(iTabData.PicToUse)
                End If
                Set iAuxPicture = mTabData(t).PicDisabled
            Else
                Set iAuxPicture = iTabData.PicToUse
            End If
        End If
        
        iPicWidth = pScaleX(iAuxPicture.Width, vbHimetric, vbPixels)
        iPicHeight = pScaleY(iAuxPicture.Height, vbHimetric, vbPixels)
        If mTabOrientation = ssTabOrientationLeft Then
            picAux.Width = iPicWidth
            picAux.Height = iPicHeight
            picAux.Cls
            picAux.BackColor = mTabBackColor
            picRotate.Cls
            picAux.PaintPicture iAuxPicture, 0, 0
            RotatePic picAux, picRotate, efn90DegreesClockWise
            Set iAuxPicture = picRotate.Image
            picRotate.Cls
            picAux.Cls
            iPicWidth = pScaleX(iAuxPicture.Width, vbHimetric, vbPixels)
            iPicHeight = pScaleY(iAuxPicture.Height, vbHimetric, vbPixels)
        ElseIf mTabOrientation = ssTabOrientationRight Then
            picAux.Width = iPicWidth
            picAux.Height = iPicHeight
            picAux.Cls
            picAux.BackColor = mTabBackColor
            picRotate.Cls
            picAux.PaintPicture iAuxPicture, 0, 0
            RotatePic picAux, picRotate, efn90DegreesCounterClockWise
            Set iAuxPicture = picRotate.Image
            picRotate.Cls
            picAux.Cls
            iPicWidth = pScaleX(iAuxPicture.Width, vbHimetric, vbPixels)
            iPicHeight = pScaleY(iAuxPicture.Height, vbHimetric, vbPixels)
        End If
        
        iTabPictureAlignment = mTabPictureAlignment
        If mTabOrientation = ssTabOrientationLeft Then
            If iTabPictureAlignment = ssPicAlignAfterCaption Then
                iTabPictureAlignment = ssPicAlignBeforeCaption
            ElseIf iTabPictureAlignment = ssPicAlignBeforeCaption Then
                iTabPictureAlignment = ssPicAlignAfterCaption
            ElseIf iTabPictureAlignment = ssPicAlignCenteredAfterCaption Then
                iTabPictureAlignment = ssPicAlignCenteredBeforeCaption
            ElseIf iTabPictureAlignment = ssPicAlignCenteredBeforeCaption Then
                iTabPictureAlignment = ssPicAlignCenteredAfterCaption
            End If
        End If
    End If
    
    If mTabOrientation = ssTabOrientationBottom Then ' in this case everything must be flipped vertically
        If Not picAux2.Font Is mFont Then
            Set picAux2.Font = mFont
        End If
        iTabSpaceRect.Left = iTabSpaceRect.Left - iTabData.TabRect.Left - 1
        iTabSpaceRect.Top = iTabSpaceRect.Top - iTabData.TabRect.Top
        iTabSpaceRect.Right = iTabSpaceRect.Right - iTabData.TabRect.Left
        iTabSpaceRect.Bottom = iTabSpaceRect.Bottom - iTabData.TabRect.Top + 2
        picAux2.Width = iTabSpaceRect.Right - iTabSpaceRect.Left
        picAux2.Height = iTabSpaceRect.Bottom - iTabSpaceRect.Top - 2
        picAux2.Cls
        picAux2.BackColor = iTabBackColor2
        If mControlIsThemed Or mAmbientUserMode And (mEnabled And iTabData.Enabled And (iTabData.Selected And mTabSelHighlight) Or (iTabData.Hovered And Not iTabData.Selected)) Then
        'If True Then
            BitBlt picAux2.hDC, 0, 0, picAux2.ScaleWidth, picAux2.ScaleHeight, picDraw.hDC, iTabData.TabRect.Left + iTabSpaceRect.Left, iTabData.TabRect.Top + iTabSpaceRect.Top + 2, vbSrcCopy
        End If
    End If
    
    If mTabOrientation = ssTabOrientationBottom Then
        Set iAuxPic = picAux2
    Else
        Set iAuxPic = picDraw
    End If
    
    iTabSpaceWidth = (iTabSpaceRect.Right - iTabSpaceRect.Left)
    iTabSpaceHeight = (iTabSpaceRect.Bottom - iTabSpaceRect.Top)
    
    ' Calculate iMeasureRect for one liner and without elipsis for both cases, WordWrap or not
    iMeasureRect = iTabSpaceRect
    
    iMeasureRect.Bottom = iMeasureRect.Top + 5
    
    iFlags = DT_CALCRECT Or DT_SINGLELINE Or DT_CENTER
    iCaption = iTabData.Caption
    DrawTextW iAuxPic.hDC, StrPtr(iCaption & IIf(iAuxPic.Font.Italic, "  ", "")), -1, iMeasureRect, iFlags Or IIf(mRightToLeft, DT_RTLREADING, 0)
    iMeasureWidth = (iMeasureRect.Right - iMeasureRect.Left)
    
    If Not iAuxPicture Is Nothing Then
        If iPicWidth + iMeasureWidth + cTabPictureDistanceToCaption > iTabSpaceWidth Then
            If iPicWidth < iTabSpaceWidth / 2 Then
                iPicWidthToShow = iPicWidth
            Else
                If mWordWrap Then
                    If iPicWidth > iTabSpaceWidth * 0.67 Then
                        iPicWidthToShow = iTabSpaceWidth * 0.67
                    Else
                        iPicWidthToShow = iPicWidth
                    End If
                Else
                    If iPicWidth > iTabSpaceWidth * 0.5 Then
                        iPicWidthToShow = iTabSpaceWidth * 0.5
                    Else
                        iPicWidthToShow = iPicWidth
                    End If
                End If
            End If
            If iPicWidthToShow + iMeasureWidth + cTabPictureDistanceToCaption < iTabSpaceWidth Then
                iPicWidthToShow = iTabSpaceWidth - iMeasureWidth - cTabPictureDistanceToCaption
            End If
            If iPicWidthToShow > iPicWidth Then
                iPicWidthToShow = iPicWidth
            End If
        Else
            iPicWidthToShow = iPicWidth
        End If
    End If
    
    If iPicHeight > iTabSpaceHeight Then
        iPicHeightToShow = iTabSpaceHeight
    Else
        iPicHeightToShow = iPicHeight
    End If
    
    iMeasureRect.Right = iTabSpaceRect.Right
    If Not iAuxPicture Is Nothing Then
        iMeasureRect.Left = iTabSpaceRect.Left + iPicWidthToShow + cTabPictureDistanceToCaption
    Else
        iMeasureRect.Left = iTabSpaceRect.Left
    End If
    iMeasureRect.Bottom = 5
    iMeasureRect.Top = 0

    iCaptionRect.Left = iMeasureRect.Left
    iCaptionRect.Right = iMeasureRect.Right
    
    ' Calculate iMeasureRect again, without elipsis for WordWrap and with elipsis for single line, and without both text centering
    If mWordWrap Then
        iFlags = DT_CALCRECT Or DT_WORDBREAK
    Else
        iFlags = DT_CALCRECT Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
    End If
    iCaption = iTabData.Caption
    DrawTextW iAuxPic.hDC, StrPtr(iCaption & IIf(iAuxPic.Font.Italic, "  ", "")), -1, iMeasureRect, iFlags Or IIf(mRightToLeft, DT_RTLREADING, 0)
    iMeasureWidth = (iMeasureRect.Right - iMeasureRect.Left)
    iMeasureHeight = (iMeasureRect.Bottom - iMeasureRect.Top)
    
    If Not iAuxPicture Is Nothing Then
        If (iTabPictureAlignment = ssPicAlignAfterCaption) Or (iTabPictureAlignment = ssPicAlignCenteredAfterCaption) Then
            iLng = iTabSpaceRect.Right - iPicWidthToShow - cTabPictureDistanceToCaption
            iCaptionRect.Left = iCaptionRect.Left - iCaptionRect.Right + iLng
            iCaptionRect.Right = iLng
        End If
    End If
    
'    If iMeasureHeight > iTabSpaceHeight Then
 '       iCaptionRect.Top = iTabSpaceRect.Top
  '      iCaptionRect.Bottom = iTabSpaceRect.Bottom
   ' Else
    iCaptionRect.Top = iTabSpaceRect.Top + iTabSpaceHeight / 2 - iMeasureHeight / 2 - 1
    iCaptionRect.Bottom = iCaptionRect.Top + iMeasureHeight + 1
    'End If
    
    If mTabOrientation = ssTabOrientationBottom Then
        iTabSpaceRect.Top = iTabSpaceRect.Top - 2
        iCaptionRect.Top = iCaptionRect.Top - 1
    End If
    
    If Not iAuxPicture Is Nothing Then
        
        ' Position of pic
        iPicTop = iTabSpaceRect.Top + iTabSpaceHeight / 2 - iPicHeightToShow / 2
        If iTabData.Caption <> "" Then
            If iTabPictureAlignment = ssPicAlignBeforeCaption Then
                iPicLeft = (iCaptionRect.Right + iCaptionRect.Left) / 2 - iMeasureWidth / 2 - cTabPictureDistanceToCaption - iPicWidthToShow
            ElseIf iTabPictureAlignment = ssPicAlignAfterCaption Then
                iPicLeft = (iCaptionRect.Right + iCaptionRect.Left) / 2 + iMeasureWidth / 2 + cTabPictureDistanceToCaption
            ElseIf iTabPictureAlignment = ssPicAlignCenteredBeforeCaption Then
                iPicLeft = iTabSpaceRect.Left + (((iCaptionRect.Right + iCaptionRect.Left) / 2 - iMeasureWidth / 2) - iTabSpaceRect.Left) / 2 - iPicWidthToShow / 2
            ElseIf iTabPictureAlignment = ssPicAlignCenteredAfterCaption Then
                iPicLeft = iTabSpaceRect.Right - (iTabSpaceRect.Right - ((iCaptionRect.Right + iCaptionRect.Left) / 2 + iMeasureWidth / 2)) / 2 - iPicWidthToShow / 2
            End If
        Else
            iPicLeft = (iTabSpaceRect.Right + iTabSpaceRect.Left) / 2 - iPicWidthToShow / 2
        End If
        If iPicLeft < iTabSpaceRect.Left Then
            iPicLeft = iTabSpaceRect.Left
        End If
        If (iPicLeft + iPicWidthToShow) > iTabSpaceRect.Right Then
            iPicLeft = iTabSpaceRect.Right - iPicWidthToShow
        End If
        
        If iPicHeightToShow >= iPicHeight Then
            iPicSourceShiftY = 0
        Else
            iPicSourceShiftY = (iPicHeight - iPicHeightToShow) / 2
        End If
        If iPicWidthToShow >= iPicWidth Then
            iPicSourceShiftX = 0
        Else
            iPicSourceShiftX = (iPicWidth - iPicWidthToShow) / 2
        End If
        
        If iPicWidth < 1 Then iPicWidth = 1
        If iPicHeight < 1 Then iPicHeight = 1
        
        ' draw the picture
        If iAuxPicture.Type = vbPicTypeBitmap And mUseMaskColor Then
            Call DrawImage(iAuxPic.hDC, iAuxPicture.Handle, TranslatedColor(mMaskColor), iPicLeft, iPicTop, iPicWidthToShow, iPicHeightToShow, iPicSourceShiftX, iPicSourceShiftY)
        Else
            iAuxPic.PaintPicture iAuxPicture, iPicLeft, iPicTop, iPicWidthToShow, iPicHeightToShow, iPicSourceShiftX, iPicSourceShiftY, iPicWidthToShow, iPicHeightToShow
        End If
    End If
    
    'Now draw the text
    If mWordWrap Then
        iFlags = DT_WORDBREAK Or DT_END_ELLIPSIS Or DT_MODIFYSTRING Or DT_CENTER
    Else
        iFlags = DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_MODIFYSTRING Or DT_CENTER Or DT_VCENTER
    End If

    iCaption = iTabData.Caption
    If iCaptionRect.Bottom > iTabData.TabRect.Bottom Then
        iCaptionRect.Bottom = iTabData.TabRect.Bottom
    End If
    iAuxPic.ForeColor = iForeColor2
    DrawTextW iAuxPic.hDC, StrPtr(iCaption), -1, iCaptionRect, iFlags Or IIf(mRightToLeft, DT_RTLREADING, 0) Or IIf(mRightToLeft, DT_RTLREADING, 0)
    
    ' Draw the focus rect
    If mAmbientUserMode Then    'only at run time
        If (t = mTabSel) And ControlHasFocus And mShowFocusRect Then
            If mAppearanceIsPP Then
                iFocusRect = iTabData.TabRect
                If mTabOrientation = ssTabOrientationBottom Then
                    iFocusRect.Right = iFocusRect.Right - iFocusRect.Left - 4
                    iFocusRect.Bottom = iFocusRect.Bottom - iFocusRect.Top '- 5
                    iFocusRect.Left = 2
                    iFocusRect.Top = 1
                Else
                    iFocusRect.Left = iFocusRect.Left + 3
                    iFocusRect.Top = iFocusRect.Top + 4
                    iFocusRect.Right = iFocusRect.Right - 2
'                    If mControlIsThemed Then
'                        iFocusRect.Top = iFocusRect.Top + 1
'                    End If
                    'iFocusRect.Bottom = iFocusRect.Bottom - 1
                    If mTabOrientation = ssTabOrientationLeft Then
                        iFocusRect.Left = iFocusRect.Left + 1
                        iFocusRect.Right = iFocusRect.Right + 1
                    End If
                End If
            Else
                If mTabOrientation = ssTabOrientationBottom Then
                    iFocusRect.Left = (iCaptionRect.Left + iCaptionRect.Right) / 2 - iMeasureWidth / 2 - 2
                    iFocusRect.Right = iFocusRect.Left + iMeasureWidth + 4
                    iFocusRect.Top = (iTabSpaceRect.Top + iTabSpaceRect.Bottom) / 2 - iMeasureHeight / 2 - 1
                    iFocusRect.Bottom = iFocusRect.Top + iMeasureHeight
                    
                    If iFocusRect.Top < 0 Then
                       iFocusRect.Top = 0
                    End If
                    If iFocusRect.Bottom > (picAux.ScaleHeight - 1) Then
                       iFocusRect.Bottom = picAux.ScaleHeight - 1
                    End If
                Else
                    iFocusRect.Left = (iCaptionRect.Left + iCaptionRect.Right) / 2 - iMeasureWidth / 2 - 2
                    iFocusRect.Right = iFocusRect.Left + iMeasureWidth + 3
                    iFocusRect.Top = (iTabSpaceRect.Top + iTabSpaceRect.Bottom) / 2 - iMeasureHeight / 2 - 1
                    iFocusRect.Bottom = iFocusRect.Top + iMeasureHeight + 2
                End If
            End If
            iAuxPic.ForeColor = iForeColor
            If mAppearanceIsPP Then
                iFocusRect.Top = iFocusRect.Top - 1
                iFocusRect.Bottom = iFocusRect.Bottom - 1
            End If
            
            If iFocusRect.Right > (iTabSpaceRect.Right) Then
                iFocusRect.Right = iTabSpaceRect.Right
            End If
            If iFocusRect.Left < (iTabSpaceRect.Left + 1) Then
                iFocusRect.Left = iTabSpaceRect.Left + 1
            End If
            If iFocusRect.Bottom > (iTabSpaceRect.Bottom) Then
                iFocusRect.Bottom = iTabSpaceRect.Bottom
            End If
            If iFocusRect.Top < (iTabSpaceRect.Top + 1) Then
                iFocusRect.Top = iTabSpaceRect.Top + 1
            End If
            
            Call DrawFocusRect(iAuxPic.hDC, iFocusRect)
        End If
    End If

    If mTabOrientation = ssTabOrientationBottom Then
        picDraw.PaintPicture picAux2.Image, iTabData.TabRect.Left + iTabSpaceRect.Left, iTabData.TabRect.Top + iTabSpaceRect.Bottom - 1, picAux2.Width, -picAux2.Height
    End If
    
    If mTabOrientation = ssTabOrientationBottom Then
        If picAux.FontBold <> iFontBoldPrev Then
            picAux.FontBold = iFontBoldPrev
        End If
    Else
        If picDraw.FontBold <> iFontBoldPrev Then
            picDraw.FontBold = iFontBoldPrev
        End If
    End If
End Sub

' The following procedure was taken from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56462&lngWId=1
' Kinda over-riden function for pFillCurvedGradientR, performs same job,
' but takes integers instead of Rect as parameter
Private Sub FillCurvedGradient(ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long, ByVal lStartColor As Long, ByVal lEndColor As Long, Optional ByVal iCurveValue As Integer = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)
    Dim utRect As RECT
    
    utRect.Left = lLeft
    utRect.Top = lTop
    utRect.Right = lRight
    utRect.Bottom = lBottom
    
    Call FillCurvedGradientR(utRect, lStartColor, lEndColor, iCurveValue, bCurveLeft, bCurveRight)
End Sub

' The following procedure was taken from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56462&lngWId=1
' function used to Fill a rectangular area by Gradient
' This function can draw using the curve value to generate a rounded rect kinda effect
Private Sub FillCurvedGradientR(utRect As RECT, ByVal lStartColor As Long, ByVal lEndColor As Long, Optional ByVal iCurveValue As Integer = -1, Optional bCurveLeft As Boolean = False, Optional bCurveRight As Boolean = False)

    Dim sngRedInc As Single, sngGreenInc As Single, sngBlueInc As Single
    Dim sngRed As Single, sngGreen As Single, sngBlue As Single
    
    lStartColor = TranslatedColor(lStartColor)
    lEndColor = TranslatedColor(lEndColor)

    Dim intCnt As Integer
    
    sngRed = GetRedValue(lStartColor)
    sngGreen = GetGreenValue(lStartColor)
    sngBlue = GetBlueValue(lStartColor)
    
    sngRedInc = (GetRedValue(lEndColor) - sngRed) / (utRect.Bottom - utRect.Top)
    sngGreenInc = (GetGreenValue(lEndColor) - sngGreen) / (utRect.Bottom - utRect.Top)
    sngBlueInc = (GetBlueValue(lEndColor) - sngBlue) / (utRect.Bottom - utRect.Top)

    If sngRed > 255 Then sngRed = 255
    If sngGreen > 255 Then sngGreen = 255
    If sngBlue > 255 Then sngBlue = 255
    If sngRed < 0 Then sngRed = 0
    If sngGreen < 0 Then sngGreen = 0
    If sngBlue < 0 Then sngBlue = 0

    If iCurveValue < 1 Then
        For intCnt = utRect.Top To utRect.Bottom
            picDraw.Line (utRect.Left, intCnt)-(utRect.Right, intCnt), RGB(sngRed, sngGreen, sngBlue)
            sngRed = sngRed + sngRedInc
            sngGreen = sngGreen + sngGreenInc
            sngBlue = sngBlue + sngBlueInc
            
            If sngRed > 255 Then sngRed = 255
            If sngGreen > 255 Then sngGreen = 255
            If sngBlue > 255 Then sngBlue = 255
            If sngRed < 0 Then sngRed = 0
            If sngGreen < 0 Then sngGreen = 0
            If sngBlue < 0 Then sngBlue = 0
        Next
    Else
        If bCurveLeft And bCurveRight Then
            For intCnt = utRect.Top To utRect.Bottom
                picDraw.Line (utRect.Left + iCurveValue + 1, intCnt)-(utRect.Right - iCurveValue, intCnt), RGB(sngRed, sngGreen, sngBlue)
                sngRed = sngRed + sngRedInc
                sngGreen = sngGreen + sngGreenInc
                sngBlue = sngBlue + sngBlueInc

                If sngRed > 255 Then sngRed = 255
                If sngGreen > 255 Then sngGreen = 255
                If sngBlue > 255 Then sngBlue = 255
                If sngRed < 0 Then sngRed = 0
                If sngGreen < 0 Then sngGreen = 0
                If sngBlue < 0 Then sngBlue = 0

                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        ElseIf bCurveLeft Then
            For intCnt = utRect.Top To utRect.Bottom
                picDraw.Line (utRect.Left + iCurveValue + 1, intCnt)-(utRect.Right, intCnt), RGB(sngRed, sngGreen, sngBlue)

                sngRed = sngRed + sngRedInc
                sngGreen = sngGreen + sngGreenInc
                sngBlue = sngBlue + sngBlueInc

                If sngRed > 255 Then sngRed = 255
                If sngGreen > 255 Then sngGreen = 255
                If sngBlue > 255 Then sngBlue = 255
                If sngRed < 0 Then sngRed = 0
                If sngGreen < 0 Then sngGreen = 0
                If sngBlue < 0 Then sngBlue = 0

                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        Else    'curve right
            For intCnt = utRect.Top To utRect.Bottom
                picDraw.Line (utRect.Left, intCnt)-(utRect.Right - iCurveValue, intCnt), RGB(sngRed, sngGreen, sngBlue)

                sngRed = sngRed + sngRedInc
                sngGreen = sngGreen + sngGreenInc
                sngBlue = sngBlue + sngBlueInc
                
                If sngRed > 255 Then sngRed = 255
                If sngGreen > 255 Then sngGreen = 255
                If sngBlue > 255 Then sngBlue = 255
                If sngRed < 0 Then sngRed = 0
                If sngGreen < 0 Then sngGreen = 0
                If sngBlue < 0 Then sngBlue = 0
                
                If iCurveValue > 0 Then
                    iCurveValue = iCurveValue - 1
                End If
            Next
        End If
    End If
End Sub

Private Sub DrawImage(ByVal lDestHDC As Long, ByVal lhBmp As Long, ByVal lTransColor As Long, ByVal iLeft As Integer, ByVal iTop As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, Optional nXOrigin As Long, Optional nYOrigin As Long)
    Dim lHDC As Long
    Dim lhBmpOld As Long
    Dim utBmp As BITMAP
    
    lHDC = CreateCompatibleDC(lDestHDC)
    lhBmpOld = SelectObject(lHDC, lhBmp)
    Call GetObjectA(lhBmp, Len(utBmp), utBmp)
    Call TransparentBlt(lDestHDC, iLeft, iTop, iWidth, iHeight, lHDC, nXOrigin, nYOrigin, iWidth, iHeight, lTransColor)
    Call SelectObject(lHDC, lhBmpOld)
    DeleteDC (lHDC)
End Sub

Private Function TranslatedColor(lOleColor As Long) As Long
    Dim lRGBColor As Long
    
    Call TranslateColor(lOleColor, 0, lRGBColor)
    TranslatedColor = lRGBColor
End Function

'extract Red component from a color
Private Function GetRedValue(RGBValue As Long) As Integer
    GetRedValue = RGBValue And &HFF
End Function

'extract Green component from a color
Private Function GetGreenValue(RGBValue As Long) As Integer
    GetGreenValue = ((RGBValue And &HFF00) / &H100) And &HFF
End Function

'extract Blue component from a color
Private Function GetBlueValue(RGBValue As Long) As Integer
    GetBlueValue = ((RGBValue And &HFF0000) / &H10000) And &HFF
End Function

Private Sub SetColors()
    Dim iBCol As Long
    Dim iCol_L As Integer
    Dim iCol_S As Integer
    Dim iCol_H As Integer
    Dim iTabBackColor_H As Integer
    Dim iTabBackColor_L As Integer
    Dim iTabBackColor_S As Integer
    Dim iAmbientBackColor_H As Integer
    Dim iAmbientBackColor_L As Integer
    Dim iAmbientBackColor_S As Integer
    Dim c As Long
    
    ResetAllPicsDisabled
    mTabBodyReset = True
    
    If mHighContrastThemeOn Or (mTabBackColor = vbButtonFace) And (Not mSoftEdges) Then
        m3DDKShadow = vb3DDKShadow
        m3DHighlight = vb3DHighlight
        m3DShadow = vb3DShadow
        mGrayText = vbGrayText
        
        iBCol = TranslatedColor(mTabBackColor)
        ColorRGBToHLS iBCol, iTabBackColor_H, iTabBackColor_L, iTabBackColor_S
        mTabBackColorDisabled = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.98, iTabBackColor_S * 0.6)
    Else
        iBCol = TranslatedColor(mTabBackColor)
        ColorRGBToHLS iBCol, iTabBackColor_H, iTabBackColor_L, iTabBackColor_S
        
        iBCol = TranslatedColor(Ambient.BackColor)
        ColorRGBToHLS iBCol, iAmbientBackColor_H, iAmbientBackColor_L, iAmbientBackColor_S
        If mSoftEdges Then
            If (iTabBackColor_L > 150) Or (iAmbientBackColor_L > 150) Then
                m3DDKShadow = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.65, iTabBackColor_S * 0.5)
                m3DShadow = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.818, iTabBackColor_S * 0.6)
            Else
                iCol_L = iTabBackColor_L * 3
                If iCol_L > 240 Then iCol_L = 240
                m3DDKShadow = ColorHLSToRGB(iTabBackColor_H, iCol_L * 0.65, iTabBackColor_S * 0.5)
                iCol_L = iTabBackColor_L * 2
                If iCol_L > 240 Then iCol_L = 240
                m3DShadow = ColorHLSToRGB(iTabBackColor_H, iCol_L * 0.818, iTabBackColor_S * 0.6)
            End If
        Else
            If (iTabBackColor_L > 150) Or (iAmbientBackColor_L > 150) Then
                m3DDKShadow = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.47, iTabBackColor_S * 0.18)
                m3DShadow = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.718, iTabBackColor_S * 0.3)
            Else
                iCol_L = iTabBackColor_L * 3
                If iCol_L > 240 Then iCol_L = 240
                m3DDKShadow = ColorHLSToRGB(iTabBackColor_H, iCol_L, iTabBackColor_S * 0.18)
                iCol_L = iTabBackColor_L * 2
                If iCol_L > 240 Then iCol_L = 240
                m3DShadow = ColorHLSToRGB(iTabBackColor_H, iCol_L * 0.718, iTabBackColor_S * 0.3)
            End If
        End If
        mTabBackColorDisabled = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.98, iTabBackColor_S * 0.6)
        mGrayText = vbGrayText
        
        If iTabBackColor_L > 150 Then
            If (iTabBackColor_L > 200) And (iTabBackColor_S < 150) Then
                iCol_L = iTabBackColor_L * 1.2
                If iCol_L > 240 Then iCol_L = 240
                m3DHighlight = ColorHLSToRGB(iTabBackColor_H, iCol_L, iTabBackColor_S * 0.3)
            Else
                iCol_L = iTabBackColor_L * 1.1
                If iCol_L > 240 Then iCol_L = 240
                m3DHighlight = ColorHLSToRGB(iTabBackColor_H, iCol_L, iTabBackColor_S * 0.2)
            End If
        Else
            iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.7
            If iCol_L > 240 Then iCol_L = 240
            m3DHighlight = ColorHLSToRGB(iTabBackColor_H, iCol_L, iTabBackColor_S * 0.9)
        End If
    End If
    mBlendDisablePicWithTabBackColor_NotThemed = (iTabBackColor_L < 200)
    If mBlendDisablePicWithTabBackColor_NotThemed Then
        mTabBackColor_R = iBCol And 255
        mTabBackColor_G = (iBCol \ 256) And 255
        mTabBackColor_B = (iBCol \ 65536) And 255
    End If
    
    If iTabBackColor_L > 150 Then
        If (iTabBackColor_L > 233) Then
            If iTabBackColor_S = 0 Then
                iCol_L = iTabBackColor_L * 0.95
                iCol_S = 0
                iCol_H = iTabBackColor_H
            Else
                iCol_L = iTabBackColor_L * 0.9
                iCol_S = 80
                iCol_H = iTabBackColor_H
            End If
            mGlowColor = ColorHLSToRGB(iCol_H, iCol_L, iCol_S)
        ElseIf (iTabBackColor_L > 200) And (iTabBackColor_S < 150) Then
            iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.1 + iTabBackColor_L * 0.05 + 5
            If iCol_L > 240 Then iCol_L = 240
            mGlowColor = ColorHLSToRGB(iTabBackColor_H, iCol_L, iTabBackColor_S)
        Else
            iCol_S = iTabBackColor_S
            If iTabBackColor_L > 160 Then
                iCol_L = iTabBackColor_L * 1.15
            Else
                iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.2 + iTabBackColor_L * 0.06 + 5
            End If
            If iCol_L > 240 Then iCol_L = 240
            If iCol_L > 200 Then
                If iTabBackColor_L > 210 Then
                    iCol_S = 1
                Else
                    If iCol_S > 100 Then
                        If ((iTabBackColor_H > 35) And (iTabBackColor_H < 45)) Then
                            iCol_S = iCol_S - 100
                            If iCol_S < 1 Then iCol_S = 1
                            iCol_L = iCol_L + 20
                            If iCol_L > 240 Then iCol_L = 240
                        End If
                    End If
                End If
            End If
            mGlowColor = ColorHLSToRGB(iTabBackColor_H, iCol_L, iCol_S)
        End If
    Else
        If iTabBackColor_S > 60 Then
            Select Case iTabBackColor_H
                Case 0 To 30, 220 To 240 ' reds
                    If iTabBackColor_L < 100 Then
                        iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.07
                    Else
                        iCol_L = iTabBackColor_L
                    End If
                Case 200 To 219 ' violet
                    iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.3
                Case 31 To 120 ' yellows, greenes, cyanes
                    iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.2
                Case Else ' blues
                    If iTabBackColor_L < 100 Then
                        iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.15
                    Else
                        iCol_L = iTabBackColor_L '+ (240 - iTabBackColor_L) * 0.07
                    End If
            End Select
        Else ' gray
            iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.2
        End If
        iCol_L = iCol_L + 15
        If iCol_L > 240 Then iCol_L = 240
        iCol_S = iTabBackColor_S
        If iCol_S > 200 Then
            iCol_S = iCol_S * 0.65
            If iCol_S < 1 Then iCol_S = 1
            iCol_L = iCol_L * 1.4
            If iCol_L > 240 Then iCol_L = 240
        Else
            iCol_S = iCol_S * 1.1
        End If
        If iCol_S > 240 Then iCol_S = 240
        
        mGlowColor = ColorHLSToRGB(iTabBackColor_H, iCol_L, iCol_S)
    End If

    For c = 1 To 5
        mHoverEffectColors(c) = MixColors(mGlowColor, mTabBackColor, 20 * c)
    Next c
    mGlowColor_Bk = mGlowColor
    
    
    ' the same for TabSel (active tab)
    If mHighContrastThemeOn Or (mTabBackColor = vbButtonFace) And (Not mSoftEdges) Then
        m3DDKShadow_Sel = vb3DDKShadow
        m3DHighlight_Sel = vb3DHighlight
        m3DShadow_Sel = vb3DShadow
        mGrayText_Sel = vbGrayText
        
        iBCol = TranslatedColor(mTabBackColor)
        ColorRGBToHLS iBCol, iTabBackColor_H, iTabBackColor_L, iTabBackColor_S
        mTabSelBackColorDisabled = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.98, iTabBackColor_S * 0.6)
    Else
        iBCol = TranslatedColor(mTabSelBackColor)
        ColorRGBToHLS iBCol, iTabBackColor_H, iTabBackColor_L, iTabBackColor_S
        
        iBCol = TranslatedColor(Ambient.BackColor)
        ColorRGBToHLS iBCol, iAmbientBackColor_H, iAmbientBackColor_L, iAmbientBackColor_S
        If mSoftEdges Then
            If (iTabBackColor_L > 150) Or (iAmbientBackColor_L > 150) Then
                m3DDKShadow_Sel = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.65, iTabBackColor_S * 0.5)
                m3DShadow_Sel = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.818, iTabBackColor_S * 0.6)
            Else
                iCol_L = iTabBackColor_L * 3
                If iCol_L > 240 Then iCol_L = 240
                m3DDKShadow_Sel = ColorHLSToRGB(iTabBackColor_H, iCol_L * 0.65, iTabBackColor_S * 0.5)
                iCol_L = iTabBackColor_L * 2
                If iCol_L > 240 Then iCol_L = 240
                m3DShadow_Sel = ColorHLSToRGB(iTabBackColor_H, iCol_L * 0.818, iTabBackColor_S * 0.6)
            End If
        Else
            If (iTabBackColor_L > 150) Or (iAmbientBackColor_L > 150) Then
                m3DDKShadow_Sel = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.47, iTabBackColor_S * 0.18)
                m3DShadow_Sel = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.718, iTabBackColor_S * 0.3)
            Else
                iCol_L = iTabBackColor_L * 3
                If iCol_L > 240 Then iCol_L = 240
                m3DDKShadow_Sel = ColorHLSToRGB(iTabBackColor_H, iCol_L, iTabBackColor_S * 0.18)
                iCol_L = iTabBackColor_L * 2
                If iCol_L > 240 Then iCol_L = 240
                m3DShadow_Sel = ColorHLSToRGB(iTabBackColor_H, iCol_L * 0.718, iTabBackColor_S * 0.3)
            End If
        End If
        mTabBackColorDisabled = ColorHLSToRGB(iTabBackColor_H, iTabBackColor_L * 0.98, iTabBackColor_S * 0.6)
        mGrayText_Sel = vbGrayText
        
        If iTabBackColor_L > 150 Then
            If (iTabBackColor_L > 200) And (iTabBackColor_S < 150) Then
                iCol_L = iTabBackColor_L * 1.2
                If iCol_L > 240 Then iCol_L = 240
                m3DHighlight_Sel = ColorHLSToRGB(iTabBackColor_H, iCol_L, iTabBackColor_S * 0.3)
            Else
                iCol_L = iTabBackColor_L * 1.1
                If iCol_L > 240 Then iCol_L = 240
                m3DHighlight_Sel = ColorHLSToRGB(iTabBackColor_H, iCol_L, iTabBackColor_S * 0.2)
            End If
        Else
            iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.7
            If iCol_L > 240 Then iCol_L = 240
            m3DHighlight_Sel = ColorHLSToRGB(iTabBackColor_H, iCol_L, iTabBackColor_S * 0.9)
        End If
    End If
    mBlendDisablePicWithTabBackColor_NotThemed = (iTabBackColor_L < 200)
    If mBlendDisablePicWithTabBackColor_NotThemed Then
        mTabSelBackColor_R = iBCol And 255
        mTabSelBackColor_G = (iBCol \ 256) And 255
        mTabSelBackColor_B = (iBCol \ 65536) And 255
    End If
    
    If iTabBackColor_L > 150 Then
        If (iTabBackColor_L > 233) Then
            If iTabBackColor_S = 0 Then
                iCol_L = iTabBackColor_L * 0.95
                iCol_S = 0
                iCol_H = iTabBackColor_H
            Else
                iCol_L = iTabBackColor_L * 0.9
                iCol_S = 80
                iCol_H = iTabBackColor_H
            End If
            mGlowColor_Sel = ColorHLSToRGB(iCol_H, iCol_L, iCol_S)
        ElseIf (iTabBackColor_L > 200) And (iTabBackColor_S < 150) Then
            iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.1 + iTabBackColor_L * 0.05 + 5
            If iCol_L > 240 Then iCol_L = 240
            mGlowColor_Sel = ColorHLSToRGB(iTabBackColor_H, iCol_L, iTabBackColor_S)
        Else
            iCol_S = iTabBackColor_S
            If iTabBackColor_L > 160 Then
                iCol_L = iTabBackColor_L * 1.15
            Else
                iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.2 + iTabBackColor_L * 0.06 + 5
            End If
            If iCol_L > 240 Then iCol_L = 240
            If iCol_L > 200 Then
                If iTabBackColor_L > 210 Then
                    iCol_S = 1
                Else
                    If iCol_S > 100 Then
                        If ((iTabBackColor_H > 35) And (iTabBackColor_H < 45)) Then
                            iCol_S = iCol_S - 100
                            If iCol_S < 1 Then iCol_S = 1
                            iCol_L = iCol_L + 20
                            If iCol_L > 240 Then iCol_L = 240
                        End If
                    End If
                End If
            End If
            mGlowColor_Sel = ColorHLSToRGB(iTabBackColor_H, iCol_L, iCol_S)
        End If
    Else
        If iTabBackColor_S > 60 Then
            Select Case iTabBackColor_H
                Case 0 To 30, 220 To 240 ' reds
                    If iTabBackColor_L < 100 Then
                        iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.07
                    Else
                        iCol_L = iTabBackColor_L
                    End If
                Case 200 To 219 ' violet
                    iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.3
                Case 31 To 120 ' yellows, greenes, cyanes
                    iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.2
                Case Else ' blues
                    If iTabBackColor_L < 100 Then
                        iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.15
                    Else
                        iCol_L = iTabBackColor_L '+ (240 - iTabBackColor_L) * 0.07
                    End If
            End Select
        Else ' gray
            iCol_L = iTabBackColor_L + (240 - iTabBackColor_L) * 0.2
        End If
        iCol_L = iCol_L + 15
        If iCol_L > 240 Then iCol_L = 240
        iCol_S = iTabBackColor_S
        If iCol_S > 200 Then
            iCol_S = iCol_S * 0.65
            If iCol_S < 1 Then iCol_S = 1
            iCol_L = iCol_L * 1.4
            If iCol_L > 240 Then iCol_L = 240
        Else
            iCol_S = iCol_S * 1.1
        End If
        If iCol_S > 240 Then iCol_S = 240
        
        mGlowColor_Sel = ColorHLSToRGB(iTabBackColor_H, iCol_L, iCol_S)
    End If

    mGlowColor_Sel_Bk = mGlowColor_Sel
    
End Sub

Private Function MixColors(nColor1 As Long, nColor2 As Long, ByVal nPercentageColor1 As Long) As Long
    Dim iColor1 As Long
    Dim iColor2 As Long
    Dim iColor1_R  As Byte
    Dim iColor1_G   As Byte
    Dim iColor1_B   As Byte
    Dim iColor2_R  As Byte
    Dim iColor2_G   As Byte
    Dim iColor2_B   As Byte
    
    iColor1 = TranslatedColor(nColor1)
    iColor2 = TranslatedColor(nColor2)
    
    iColor1_R = iColor1 And 255
    iColor1_G = (iColor1 \ 256) And 255
    iColor1_B = (iColor1 \ 65536) And 255
    iColor2_R = iColor2 And 255
    iColor2_G = (iColor2 \ 256) And 255
    iColor2_B = (iColor2 \ 65536) And 255
    
    If nPercentageColor1 > 100 Then nPercentageColor1 = 100
    If nPercentageColor1 < 0 Then nPercentageColor1 = 0
    
    MixColors = RGB((iColor1_R * nPercentageColor1 + iColor2_R * (100 - nPercentageColor1)) / 100, (iColor1_G * nPercentageColor1 + iColor2_G * (100 - nPercentageColor1)) / 100, (iColor1_B * nPercentageColor1 + iColor2_B * (100 - nPercentageColor1)) / 100)
    
End Function



Public Property Get TabBodyLeft() As Single
Attribute TabBodyLeft.VB_Description = "Returns the left of the tab body."
    EnsureDrawn
    TabBodyLeft = FixRoundingError(UserControl.ScaleX(mTabBodyRect.Left, vbPixels, vbTwips))
End Property

Public Property Get TabBodyTop() As Single
Attribute TabBodyTop.VB_Description = "Returns the space occupied by tabs, or where the tab body begins."
    EnsureDrawn
    TabBodyTop = FixRoundingError(UserControl.ScaleY(mTabBodyRect.Top, vbPixels, vbTwips))
End Property

Public Property Get TabBodyWidth() As Single
Attribute TabBodyWidth.VB_Description = "Returns the width of the tab body."
    EnsureDrawn
    TabBodyWidth = FixRoundingError(UserControl.ScaleX(mTabBodyRect.Right - mTabBodyRect.Left, vbPixels, vbTwips))
End Property

Public Property Get TabBodyHeight() As Single
Attribute TabBodyHeight.VB_Description = "Returns the height of the tab body."
    EnsureDrawn
    TabBodyHeight = FixRoundingError(UserControl.ScaleY(mTabBodyRect.Bottom - mTabBodyRect.Top, vbPixels, vbTwips))
End Property

Private Sub EnsureDrawn()
    If (Not mFirstDraw) Or tmrDraw.Enabled Or mDrawMessagePosted Then
        mEnsureDrawn = True
        Draw
        mEnsureDrawn = False
    End If
End Sub

Private Sub RotatePic(picSrc As PictureBox, picDest As PictureBox, nDirection As efnRotatePicDirection)
    Dim pt(0 To 2) As POINTAPI
    Dim iHeight As Long
    Dim iWidth As Long
    
    iHeight = picSrc.Height
    iWidth = picSrc.Width
    
    'Set the coordinates of the parallelogram
    If nDirection = efn90DegreesClockWise Then ' 90 degrees
        pt(0).X = iHeight
        pt(0).Y = 0
        pt(1).X = iHeight
        pt(1).Y = iWidth
        pt(2).X = 0
        pt(2).Y = 0
    ElseIf nDirection = efn90DegreesCounterClockWise Then ' 270 degrees
        pt(0).X = 0
        pt(0).Y = iWidth
        pt(1).X = 0
        pt(1).Y = 0
        pt(2).X = iHeight
        pt(2).Y = iWidth
    ElseIf nDirection = efnFlipVertical Then ' vertical
        pt(0).X = 0
        pt(0).Y = iHeight
        pt(1).X = iWidth
        pt(1).Y = iHeight
        pt(2).X = 0
        pt(2).Y = 0
    ElseIf nDirection = efnFlipHorizontal Then ' horizontal
        pt(0).X = iWidth
        pt(0).Y = 0
        pt(1).X = 0
        pt(1).Y = 0
        pt(2).X = iWidth
        pt(2).Y = iHeight
    End If
    
    picDest.Width = picSrc.Height
    picDest.Height = picSrc.Width
    picDest.Cls
    
    picDest.Cls
    PlgBlt picDest.hDC, pt(0), picSrc.hDC, 0, 0, iWidth, iHeight, ByVal 0&, ByVal 0&, ByVal 0&
End Sub

Private Function ContainerScaleMode() As ScaleModeConstants
    ContainerScaleMode = vbTwips
    On Error Resume Next
    ContainerScaleMode = UserControl.Extender.Container.ScaleMode
    Err.Clear
End Function

Private Function FromContainerSizeY(nValue, Optional nToScale As ScaleModeConstants = vbTwips) As Single
    FromContainerSizeY = pScaleY(nValue, ContainerScaleMode, nToScale)
End Function

Private Function ToContainerSizeY(nValue, Optional nFromScale As ScaleModeConstants = vbTwips) As Single
    ToContainerSizeY = pScaleY(nValue, nFromScale, ContainerScaleMode)
End Function


Private Function FromContainerSizeX(nValue, Optional nToScale As ScaleModeConstants = vbTwips) As Single
    FromContainerSizeX = pScaleX(nValue, ContainerScaleMode, nToScale)
End Function

Private Function ToContainerSizeX(nValue, Optional nFromScale As ScaleModeConstants = vbTwips) As Single
    ToContainerSizeX = pScaleX(nValue, nFromScale, ContainerScaleMode)
End Function

Private Function FixRoundingError(nNumber As Single, Optional nDecimals As Long = 3) As Single
    Dim iNum As Single
    
    iNum = Round(nNumber * 10 ^ nDecimals) / 10 ^ nDecimals
    If iNum = Int(iNum) Then
        FixRoundingError = iNum
    Else
        If (ContainerScaleMode = vbTwips) Or (ContainerScaleMode = vbPixels) Then
            FixRoundingError = Round(nNumber)
        Else
            FixRoundingError = nNumber
        End If
    End If
End Function
    
Private Sub SetControlsBackColor(nColor As Long, Optional nPrevColor As Long = -1)
    Dim iCtl As Control
    Dim iLng As Long
    Dim iCancel As Boolean
    Dim iControls As Object
    Dim iContainer As Object
    Dim iContainer_Prev As Object
    Dim iStr As String
    
    On Error Resume Next
    Set iControls = UserControl.Parent.Controls
    
    If iControls Is Nothing Then ' at least let's change the backcolor of the contained controls in the usercontrol
        For Each iCtl In UserControl.ContainedControls
            iLng = -1
            iLng = iCtl.BackColor
            If iLng <> -1 Then
                If (iLng = vbButtonFace) And (nColor <> vbButtonFace) Or (iLng = nPrevColor) Then
                    iCancel = False
                    iStr = iCtl.Name
                    RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                    If Not iCancel Then
                        iCtl.BackColor = nColor
                    End If
                End If
            End If
        Next
    Else 'let's change the backcolor of all the controls inside the tabs
        For Each iCtl In iControls
            Set iContainer = Nothing
            Set iContainer = iCtl.Container
            Do Until iContainer Is Nothing
                If iContainer Is UserControl.Extender Then
                    iLng = -1
                    iLng = iCtl.BackColor
                    If iLng <> -1 Then
                        If (iLng = vbButtonFace) And (nColor <> vbButtonFace) Or (iLng = nPrevColor) Then
                            iCancel = False
                            If Not iContainer_Prev Is Nothing Then
                                If iContainer_Prev.Container Is UserControl.Extender Then
                                    iStr = iContainer_Prev.Name
                                    RaiseEvent ChangeControlBackColor(iStr, TypeName(iContainer_Prev), iCancel)
                                End If
                            End If
                            If Not iCancel Then
                                iCancel = False
                                iStr = iCtl.Name
                                RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                If Not iCancel Then
                                    iCtl.BackColor = nColor
                                End If
                            End If
                        End If
                    End If
                End If
                Set iContainer_Prev = iContainer
                Set iContainer = Nothing
                Set iContainer = iContainer_Prev.Container
            Loop
        Next
    End If
    Err.Clear
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Redraws the control."
    Dim iWv As Boolean
    
    iWv = IsWindowVisible(mUserControlHwnd) <> 0
    If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, False, 0&
    mTabBodyReset = True
    If mChangeControlsBackColor Then
        SetControlsBackColor mTabSelBackColor
    End If
    StoreControlsTabStop
    mRedraw = True
    mSubclassControlsPaintingPending = True
    mRepaintSubclassedControls = True
    Draw
    If iWv Then SendMessage mUserControlHwnd, WM_SETREDRAW, True, 0&
    If iWv Then RedrawWindow mUserControlHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
End Sub

Private Sub RaiseEvent_TabMouseEnter(nTab As Integer)
    mTabData(nTab).Hovered = True
    RaiseEvent TabMouseEnter(nTab)
    If (mTabHoverHighlight = ssTHHEffect) And Not mControlIsThemed Then
        tmrTabHoverEffect.Enabled = False
        tmrTabHoverEffect.Enabled = True
        mTabHoverEffect_Step = 1
        mGlowColor = mHoverEffectColors(mTabHoverEffect_Step)
    End If
    If (mTabHoverHighlight <> ssTHHNo) Then PostDrawMessage
    
    If mThereAreTabsToolTipTexts Then
        ShowTabTTT nTab
    End If
End Sub

Private Sub RaiseEvent_TabMouseLeave(nTab As Integer)
    If tmrTabHoverEffect.Enabled Then
        tmrTabHoverEffect.Enabled = False
        mGlowColor = mGlowColor_Bk
    End If
    mTabData(nTab).Hovered = False
    RaiseEvent TabMouseLeave(nTab)
    If nTab <> mTabSel Then
        If (mTabHoverHighlight <> ssTHHNo) Then PostDrawMessage
    End If
    
    If mThereAreTabsToolTipTexts Then RestoreExtenderTTT
End Sub

Private Sub ShowTabTTT(nTab As Integer)
    Dim iTCtl As String
    Dim iTTab As String
    
    iTTab = mTabData(nTab).ToolTipText
    On Error Resume Next
    iTCtl = UserControl.Extender.ToolTipText
    On Error GoTo 0
    If (iTCtl <> mLastTabToolTipTextSet) And (mLastTabToolTipTextSet <> "") Then
        mExtenderToolTipText = iTCtl
        mLastTabToolTipTextSet = ""
    End If
    If (iTTab = "") And (iTCtl = mLastTabToolTipTextSet) Then
        On Error Resume Next
        UserControl.Extender.ToolTipText = mExtenderToolTipText
        On Error GoTo 0
        mLastTabToolTipTextSet = ""
        mExtenderToolTipText = ""
    End If
    If (iTTab <> "") Then
        mExtenderToolTipText = iTCtl
        On Error Resume Next
        UserControl.Extender.ToolTipText = iTTab
        mLastTabToolTipTextSet = iTTab
        On Error GoTo 0
    End If
End Sub

Private Sub RestoreExtenderTTT()
    Dim iTCtl As String
    
    On Error Resume Next
    iTCtl = UserControl.Extender.ToolTipText
    On Error GoTo 0
    If (iTCtl = mLastTabToolTipTextSet) Then
        On Error Resume Next
        UserControl.Extender.ToolTipText = mExtenderToolTipText
        On Error GoTo 0
    End If
    mExtenderToolTipText = ""
    mLastTabToolTipTextSet = ""
End Sub

Private Sub CheckIfThereAreTabsToolTipTexts()
    Dim c As Long
    
    If Not mAmbientUserMode Then Exit Sub
    mThereAreTabsToolTipTexts = False
    For c = 0 To mTabs - 1
        If mTabData(c).ToolTipText <> "" Then
            mThereAreTabsToolTipTexts = True
            Exit Sub
        End If
    Next c
End Sub

Private Sub SetButtonFaceColor()
    Dim iCol As Long
    
    iCol = TranslatedColor(vbButtonFace)
    ColorRGBToHLS iCol, mButtonFace_H, mButtonFace_L, mButtonFace_S
    
End Sub

Private Sub SetThemedTabTransparentPixels(nIsLeftTab As Boolean, nIsRightTab As Boolean, nIsTopTab As Boolean)
    Dim X As Long
    Dim X2 As Long
    Dim iYLenght As Long
    
    If nIsLeftTab Or nIsTopTab Then
        For X = 0 To 5
            iYLenght = mTABITEM_TopLeftCornerTransparencyMask(X)
            If iYLenght < 0 Then
                iYLenght = picAux.ScaleHeight - iYLenght
            End If
            If iYLenght > 0 Then
                picAux.Line (X, 0)-(X, iYLenght), cAuxTransparentColor
            End If
        Next X
    End If
    If nIsRightTab Then
        For X = 0 To 5
            X2 = picAux.ScaleWidth - 1 - X
            iYLenght = mTABITEMRIGHTEDGE_RightSideTransparencyMask(X)
            If iYLenght < 0 Then
                iYLenght = picAux.ScaleHeight - iYLenght
            End If
            If iYLenght > 0 Then
                picAux.Line (X2, 0)-(X2, iYLenght), cAuxTransparentColor
            End If
        Next X
    ElseIf nIsTopTab Then
        For X = 0 To 5
            X2 = picAux.ScaleWidth - 1 - X
            iYLenght = mTABITEM_TopRightCornerTransparencyMask(X)
            If iYLenght < 0 Then
                iYLenght = picAux.ScaleHeight - iYLenght
            End If
            If iYLenght > 0 Then
                picAux.Line (X2, 0)-(X2, iYLenght), cAuxTransparentColor
            End If
        Next X
    End If
    
End Sub

Private Sub EnsureTabBodyThemedReady()
    If Not mTabBodyThemedReady Then
        Dim iRect As RECT
        
        iRect.Left = 0
        iRect.Top = 0
        iRect.Right = mTabBodyWidth  '+ 1 '- 1
        iRect.Bottom = mTabBodyHeight '- 1 '+ 1 '- 1
        picTabBodyThemed.Width = iRect.Right
        picTabBodyThemed.Height = iRect.Bottom
        picTabBodyThemed.BackColor = mBackColor
        picTabBodyThemed.Cls
        If (mTabOrientation = ssTabOrientationTop) Then
            DrawThemeBackground mTheme, picTabBodyThemed.hDC, TABP_PANE, 0&, iRect, iRect
        ElseIf (mTabOrientation = ssTabOrientationLeft) Then
            ' shadow must be at the bottom, and since the image will be rotated it must be at the left here.
            picAux.Cls
            picAux.Width = picTabBodyThemed.Width
            picAux.Height = picTabBodyThemed.Height
            DrawThemeBackground mTheme, picAux.hDC, TABP_PANE, 0&, iRect, iRect
            picTabBodyThemed.PaintPicture picAux.Image, picAux.ScaleWidth - 1, 0, -picAux.ScaleWidth, picAux.ScaleHeight
        Else ' (mTabOrientation = ssTabOrientationBottom) Or (mTabOrientation = ssTabOrientationRight)
            picAux.Cls
            picAux.Width = picTabBodyThemed.Width
            picAux.Height = picTabBodyThemed.Height
            iRect.Bottom = iRect.Bottom + mThemedTabBodyBottomShadowPixels
            DrawThemeBackground mTheme, picAux.hDC, TABP_PANE, 0&, iRect, iRect
            picTabBodyThemed.PaintPicture picAux.Image, 0, picAux.ScaleHeight - 1, picAux.ScaleWidth, -picAux.ScaleHeight
        End If
        mThemedTabBodyReferenceTopBackColor = GetPixel(picTabBodyThemed.hDC, picTabBodyThemed.ScaleWidth / 2, picTabBodyThemed.ScaleHeight * 0.1)
        mTabBodyThemedReady = True
    End If
End Sub

Private Sub EnsureInactiveTabBodyThemedReady()
    If Not mInactiveTabBodyThemedReady Then
        Dim iCA As COLORADJUSTMENT
        
        EnsureTabBodyThemedReady
        picInactiveTabBodyThemed.Width = picTabBodyThemed.Width
        picInactiveTabBodyThemed.Height = picTabBodyThemed.Height
        iCA = GetInactiveTabBodyColorAdjustment
        picAux2.Cls
        picAux2.Width = picTabBodyThemed.Width
        picAux2.Height = picTabBodyThemed.Height
        
        SetStretchBltMode picAux2.hDC, HALFTONE
        SetColorAdjustment picAux2.hDC, iCA
        
        StretchBlt picAux2.hDC, 0, 0, picTabBodyThemed.Width, picTabBodyThemed.Height, picTabBodyThemed.hDC, 0, 0, picTabBodyThemed.Width, picTabBodyThemed.Height, vbSrcCopy
        picInactiveTabBodyThemed.Cls
        BitBlt picInactiveTabBodyThemed.hDC, 0, 0, picAux2.ScaleWidth, picAux2.ScaleHeight, picAux2.hDC, 0, 0, vbSrcCopy
        mInactiveTabBodyThemedReady = True
    End If
End Sub

Private Sub SetThemeExtraData()
    Dim iRect As RECT
    Dim X As Long
    Dim X2 As Long
    Dim Y As Long
    Dim iCol As Long
    Dim iCol_H As Integer
    Dim iCol_L As Integer
    Dim iCol_S As Integer
    Dim iToChange As Boolean
    Const cHTolerance As Integer = 3
    Const cLTolerance As Integer = 5
    Const cSTolerance As Integer = 14
    Dim iColB As Long
    Dim iColB_H As Integer
    Dim iColB_L As Integer
    Dim iColB_S As Integer
    Dim iThreshold As Integer
    
    If mThemeExtraDataAlreadySet Then Exit Sub
    mThemeExtraDataAlreadySet = True
    
    iRect.Left = 0
    iRect.Top = 0
    iRect.Right = 30
    iRect.Bottom = 30
    picAux.Width = 30
    picAux.Height = 30
    
    DrawThemeBackground mTheme, picAux.hDC, TABP_TABITEM, TIS_NORMAL, iRect, iRect
    mThemedInactiveReferenceTabBackColor = GetPixel(picAux.hDC, 15, 27)
    ColorRGBToHLS mThemedInactiveReferenceTabBackColor, mThemedInactiveReferenceTabBackColor_H, mThemedInactiveReferenceTabBackColor_L, mThemedInactiveReferenceTabBackColor_S
    
    ' transparency mask for top left corner of TABITEM and TABITEMRIGHTEDGE
    For X = 0 To 5
        mTABITEM_TopLeftCornerTransparencyMask(X) = 0
    Next X
    For X = 0 To 5
        For Y = 0 To picAux.ScaleHeight - 1
            iToChange = False
            iCol = GetPixel(picAux.hDC, X, Y)
            ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
            If Abs(iCol_H - mButtonFace_H) <= cHTolerance Then
                If Abs(iCol_L - mButtonFace_L) <= cLTolerance Then
                    If Abs(iCol_S - mButtonFace_S) <= cSTolerance Then
                        iToChange = True
                    End If
                End If
            End If
            If Not iToChange Then
                If Y < (6) Then
                    mTABITEM_TopLeftCornerTransparencyMask(X) = Y
                Else
                    mTABITEM_TopLeftCornerTransparencyMask(X) = Y - picAux.ScaleHeight - 1 ' negative values point to pixels left to reach the bottom
                End If
                Exit For
            End If
        Next Y
        If Y = picAux.ScaleHeight Then
            mTABITEM_TopLeftCornerTransparencyMask(X) = -1
        End If
        If mTABITEM_TopLeftCornerTransparencyMask(X) = 0 Then Exit For
    Next X
    
    ' transparency mask for top right corner of TABITEM
    For X = 0 To 5
        mTABITEM_TopRightCornerTransparencyMask(X) = 0
    Next X
    For X = 0 To 5
        X2 = picAux.ScaleWidth - 1 - X
        For Y = 0 To picAux.ScaleHeight - 1
            iToChange = False
            iCol = GetPixel(picAux.hDC, X2, Y)
            ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
            If Abs(iCol_H - mButtonFace_H) <= cHTolerance Then
                If Abs(iCol_L - mButtonFace_L) <= cLTolerance Then
                    If Abs(iCol_S - mButtonFace_S) <= cSTolerance Then
                        iToChange = True
                    End If
                End If
            End If
            If Not iToChange Then
                If Y < (6) Then
                    mTABITEM_TopRightCornerTransparencyMask(X) = Y
                Else
                    mTABITEM_TopRightCornerTransparencyMask(X) = Y - picAux.ScaleHeight - 1 ' negative values point to pixels left to reach the bottom
                End If
                Exit For
            End If
        Next Y
        If Y = picAux.ScaleHeight Then
            mTABITEM_TopRightCornerTransparencyMask(X) = -1
        End If
        If mTABITEM_TopRightCornerTransparencyMask(X) = 0 Then Exit For
    Next X
    
    ' transparency mask for right side of TABITEMRIGHTEDGE
    picAux.Cls
    DrawThemeBackground mTheme, picAux.hDC, TABP_TABITEMRIGHTEDGE, TIS_NORMAL, iRect, iRect
    For X = 0 To 5
        mTABITEMRIGHTEDGE_RightSideTransparencyMask(X) = 0
    Next X
    For X = 0 To 5
        X2 = picAux.ScaleWidth - 1 - X
        For Y = 0 To picAux.ScaleHeight - 1
            iToChange = False
            iCol = GetPixel(picAux.hDC, X2, Y)
            ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
            If Abs(iCol_H - mButtonFace_H) <= cHTolerance Then
                If Abs(iCol_L - mButtonFace_L) <= cLTolerance Then
                    If Abs(iCol_S - mButtonFace_S) <= cSTolerance Then
                        iToChange = True
                    End If
                End If
            End If
            If Not iToChange Then
                If Y < (6) Then
                    mTABITEMRIGHTEDGE_RightSideTransparencyMask(X) = Y
                Else
                    mTABITEMRIGHTEDGE_RightSideTransparencyMask(X) = Y - picAux.ScaleHeight - 1 ' negative values point to pixels left to reach the bottom
                End If
                Exit For
            End If
        Next Y
        If Y = picAux.ScaleHeight Then
            mTABITEMRIGHTEDGE_RightSideTransparencyMask(X) = -1 ' all the column of pixels
        End If
        If mTABITEMRIGHTEDGE_RightSideTransparencyMask(X) = 0 Then Exit For
    Next X
    
    DrawThemeBackground mTheme, picAux.hDC, TABP_PANE, 0&, iRect, iRect
    iColB = GetPixel(picAux.hDC, 15, 10)
    ColorRGBToHLS iColB, iColB_H, iColB_L, iColB_S
    
    mBlendDisablePicWithTabBackColor_Themed = (iColB_L <= 200)
    If mBlendDisablePicWithTabBackColor_Themed Then
        mThemedTabBodyBackColor_R = iColB And 255
        mThemedTabBodyBackColor_G = (iColB \ 256) And 255
        mThemedTabBodyBackColor_B = (iColB \ 65536) And 255
    End If
    
    iThreshold = 120
    mThemedTabBodyBottomShadowPixels = 0
    Do
        For Y = picAux.ScaleHeight - 9 To picAux.ScaleHeight - 1
            iCol = GetPixel(picAux.hDC, 15, Y)
            ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
            If Abs(iCol_L - iColB_L) > iThreshold Then
                mThemedTabBodyBottomShadowPixels = picAux.ScaleHeight - Y - 1
                Exit For
            End If
        Next Y
        If mThemedTabBodyBottomShadowPixels = 0 Then
            iThreshold = iThreshold - 10
            If iThreshold < 1 Then
                iThreshold = 20
                Exit Do
            End If
        End If
    Loop While mThemedTabBodyBottomShadowPixels = 0
    
    mThemedTabBodyRightShadowPixels = 0
    For X = picAux.ScaleWidth - 9 To picAux.ScaleWidth - 1
        iCol = GetPixel(picAux.hDC, X, 15)
        ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
        If Abs(iCol_L - iColB_L) > iThreshold Then
            mThemedTabBodyRightShadowPixels = picAux.ScaleWidth - X - 1
            Exit For
        End If
    Next X
    
    picAux.Cls
End Sub

Private Function GetInactiveTabBodyColorAdjustment() As COLORADJUSTMENT
    Dim iCA As COLORADJUSTMENT
    Dim iCol As Long
    Dim iCol_H As Integer
    Dim iCol_L As Integer
    Dim iCol_S As Integer
    Dim c As Long
    Dim iLng As Long
    
    picAux2.Width = 1
    picAux2.Height = 1
    picAux2.Cls
    SetStretchBltMode picAux2.hDC, HALFTONE
    GetColorAdjustment picAux2.hDC, iCA
    
    picAux.Width = 1
    picAux.Height = 1
    SetPixelV picAux.hDC, 0, 0, mThemedTabBodyReferenceTopBackColor
    
    ' luminance
    c = 0
    Do
        c = c + 1
        StretchBlt picAux2.hDC, 0, 0, 1, 1, picAux.hDC, 0, 0, 1, 1, vbSrcCopy
        iCol = GetPixel(picAux2.hDC, 0, 0)
        ColorRGBToHLS iCol, iCol_H, iCol_L, iCol_S
        If Abs(mThemedInactiveReferenceTabBackColor_L - iCol_L) < 3 Then
            Exit Do
        ElseIf c > 5 Then
            Exit Do
        End If
        iLng = mThemedInactiveReferenceTabBackColor_L - iCol_L
        If iLng > 50 Then iLng = 50
        If iLng < -50 Then iLng = -50
        iCA.caBrightness = iLng
        SetColorAdjustment picAux2.hDC, iCA
    Loop
    
    GetInactiveTabBodyColorAdjustment = iCA
End Function

Private Sub ResetCachedThemeImages()
    mTabBodyThemedReady = False
    mInactiveTabBodyThemedReady = False
    mSubclassControlsPaintingPending = True
    mRepaintSubclassedControls = True
    mTabBodyReset = True
End Sub

Private Function MeasureTabPictureAndCaption(t As Long) As Long
    Dim iPicWidth As Long
    Dim iCaptionWidth As Long
    Dim iCaptionRect As RECT
    Dim iTabMaxWidth As Long
    Dim iFlags As Long
    Dim iFontBoldPrev As Boolean
    Dim iCaption As String
    
    ' pic
    iPicWidth = 0
    If Not mTabData(t).PicToUseSet Then SetPicToUse t
    If Not mTabData(t).PicToUse Is Nothing Then
        If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
            iPicWidth = pScaleX(mTabData(t).PicToUse.Width, vbHimetric, vbPixels)
        Else
            iPicWidth = pScaleX(mTabData(t).PicToUse.Height, vbHimetric, vbPixels)
        End If
    End If
    
    ' caption
    iFontBoldPrev = picAux.FontBold
    If t = mTabSel Then
        If mAppearanceIsPP And (mTabSelFontBold = ssYNAuto) Then
            picAux.FontBold = mFont.Bold
        ElseIf (mTabSelFontBold = ssYes) Or (mTabSelFontBold = ssYNAuto) Then
            picAux.FontBold = True
        Else
            picAux.FontBold = False
        End If
    Else
        picAux.FontBold = mFont.Bold
    End If
    
    With mTabData(t).TabRect
        iCaptionRect.Left = 0
        iCaptionRect.Top = 0
        iCaptionRect.Bottom = .Bottom - .Top - 4
        iCaptionRect.Right = mScaleWidth
    End With
    
    If mTabMaxWidth > 0 Then
        iTabMaxWidth = pScaleX(mTabMaxWidth, vbHimetric, vbPixels)
        iCaptionRect.Right = iTabMaxWidth
    Else
        iFlags = DT_CALCRECT Or DT_SINGLELINE Or DT_CENTER Or DT_VCENTER
        iCaption = mTabData(t).Caption & IIf(picAux.Font.Italic, " ", "") & IIf((mTabWidthStyle = ssTWSAuto) And ((mStyle = ssStyleTabStrip) Or (mVisualStyles And (IsAppThemeEnabled Or mForceVisualStyles))), "  ", "")
        DrawTextW picAux.hDC, StrPtr(iCaption), -1, iCaptionRect, iFlags
    End If
    iCaptionWidth = iCaptionRect.Right '- iCaptionRect.Left
    
    If picAux.FontBold <> iFontBoldPrev Then
        picAux.FontBold = iFontBoldPrev
    End If
    
    MeasureTabPictureAndCaption = iPicWidth + cTabPictureDistanceToCaption + iCaptionWidth
End Function

Public Function IsVisualStyleApplied() As Boolean
Attribute IsVisualStyleApplied.VB_Description = "Returns a boolean value indicating whether the visual styles are actually applied to the control."
    Dim iTheme As Long
    
    IsVisualStyleApplied = mVisualStyles And (IsAppThemeEnabled Or mForceVisualStyles) And (mBackStyle <> ssTransparent)
    If IsVisualStyleApplied Then
        iTheme = OpenThemeData(mUserControlHwnd, StrPtr("Tab"))
        If iTheme = 0 Then
            IsVisualStyleApplied = False
        Else
            CloseThemeData iTheme
        End If
    End If
End Function

' Hidden property mainly intended for testing purposes
Public Property Get ForceVisualStyles() As Boolean
Attribute ForceVisualStyles.VB_Description = "Hidden property intended for testing purposes. Allows the control to show visual styles on an un-themed IDE."
Attribute ForceVisualStyles.VB_MemberFlags = "40"
    ForceVisualStyles = mForceVisualStyles
End Property

Public Property Let ForceVisualStyles(ByVal nValue As Boolean)
    If nValue <> mForceVisualStyles Then
        mForceVisualStyles = nValue
        If (Not (mVisualStyles And IsAppThemeEnabled)) Or (Not mForceVisualStyles) Then
            mSubclassControlsPaintingPending = True
            mRepaintSubclassedControls = True
            PostDrawMessage
        End If
    End If
End Property

Private Function IsAppThemeEnabled() As Boolean
    If GetComCtlVersion() >= 6 Then
        If IsThemeActive() <> 0 Then
            If IsAppThemed() <> 0 Then
                IsAppThemeEnabled = True
            ElseIf (GetThemeAppProperties() And STAP_ALLOW_CONTROLS) <> 0 Then
                IsAppThemeEnabled = True
            End If
        End If
    End If
End Function


Private Function GetComCtlVersion() As Long
    Static sValue As Long
    
    If sValue = 0 Then
        Dim iVersion As DLLVERSIONINFO
        On Error Resume Next
        iVersion.cbSize = LenB(iVersion)
        If DllGetVersion(iVersion) = S_OK Then
            sValue = iVersion.dwMajor
        End If
        Err.Clear
    End If
    GetComCtlVersion = sValue
End Function

Private Sub SetVisibleControls(iPreviousTab As Integer)
    Dim iCtl As Control
    Dim iCtlName
    Dim iContainedControlsString As String
    Dim iHwnd As Long
    Dim c As Long
    Dim iLeft As Long
    Dim iIsLine As Boolean
    
    If mUserControlTerminated Then Exit Sub
    If Not Ambient.UserMode Then CheckIfContainedControlChangedToArray
    
    If (Not mAmbientUserMode) And mChangeControlsBackColor And (mTabBackColor <> vbButtonFace) Then
        iContainedControlsString = GetContainedControlsString
        If iContainedControlsString <> mLastContainedControlsString Then
            SetControlsBackColor mTabSelBackColor
        End If
    End If
    
    If mSubclassedControlsForMoveHwnds.Count > 0 Then
        For c = 1 To mSubclassedControlsForMoveHwnds.Count
            iHwnd = mSubclassedControlsForMoveHwnds(c)
            DetachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
        Next c
        Set mSubclassedControlsForMoveHwnds = New Collection
    End If
    
    If mPendingLeftShift <> 0 Then
        DoPendingLeftShift
    End If
    
    ' hide controls in previous tab
    If mAmbientUserMode Then StoreControlsTabStop
    If (iPreviousTab >= 0) And (iPreviousTab <= UBound(mTabData)) Then
        Set mTabData(iPreviousTab).Controls = New Collection
    End If
    For Each iCtl In UserControl.ContainedControls
        iIsLine = TypeName(iCtl) = "Line"
        iLeft = -15001
        On Error Resume Next
        If iIsLine Then
            iLeft = iCtl.X1
        Else
            iLeft = iCtl.Left
        End If
        On Error GoTo 0
        If iLeft > -mLeftThresholdHided Then
            iCtlName = ControlName(iCtl)
            If (iPreviousTab >= 0) And (iPreviousTab <= UBound(mTabData)) Then
                If Not IsControlInOtherTab(iCtlName, iPreviousTab) Then
                    mTabData(iPreviousTab).Controls.Add iCtlName, iCtlName
                End If
            End If
            If iIsLine Then
                iCtl.X1 = iCtl.X1 - mLeftShiftToHide
                iCtl.X2 = iCtl.X2 - mLeftShiftToHide
            Else
                iCtl.Left = iCtl.Left - mLeftShiftToHide
            End If
        End If
    Next
    
    ' show controls in selected tab
    If (mTabSel > -1) And (mTabSel < mTabs) Then
        For Each iCtlName In mTabData(mTabSel).Controls
            Set iCtl = GetContainedControlByName(iCtlName)
            If Not iCtl Is Nothing Then
                On Error Resume Next
                iIsLine = TypeName(iCtl) = "Line"
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 + mLeftShiftToHide
                    iCtl.X2 = iCtl.X2 + mLeftShiftToHide
                Else
                    iCtl.Left = iCtl.Left + mLeftShiftToHide
                End If
                On Error GoTo 0
                If mAmbientUserMode Then
                    On Error Resume Next
                    iCtl.TabStop = mParentControlsTabStop(iCtlName)
                    iCtl.UseMnemonic = mParentControlsUseMnemonic(iCtlName)
                    If TypeName(iCtl) = "ComboBox" Then
                        ' ComboBox fix
                        If iCtl.Style = vbComboDropdown Then
                            iCtl.SelLength = 0
                        End If
                    End If
                    On Error GoTo 0
                    If ControlIsContainer(iCtlName) Then
                        SetTabStopsToParentControlsContainedInControl iCtl
                    End If
                End If
            End If
        Next
    End If
    
    If (Not mAmbientUserMode) And mChangeControlsBackColor Then
        mLastContainedControlsString = iContainedControlsString
    End If
    
    If mSubclassed Then
        On Error Resume Next
        For Each iCtl In UserControl.ContainedControls
            If iCtl.Left < -mLeftThresholdHided Then
                iHwnd = 0
                iHwnd = iCtl.hWnd
                If iHwnd <> 0 Then
                    mSubclassedControlsForMoveHwnds.Add iHwnd
                    AttachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
                End If
            End If
        Next
        Err.Clear
    End If

End Sub

Private Function IsControlInOtherTab(nCtlName, nTab As Integer) As Boolean
    Dim t As Long
    Dim iStr As String
    
    On Error Resume Next
    For t = 0 To mTabs - 1
        If t <> nTab Then
            iStr = ""
            iStr = mTabData(t).Controls(nCtlName)
            If iStr <> "" Then
                IsControlInOtherTab = True
                Exit Function
            End If
        End If
    Next t
End Function

Private Function GetContainedControlsString() As String
    Dim iCtl As Control
    
    For Each iCtl In UserControl.ContainedControls
        GetContainedControlsString = GetContainedControlsString & iCtl.Name
    Next
End Function

Private Sub StoreControlsTabStop(Optional nInitialize As Boolean)
    Dim iControls As Object
    Dim iCtl As Control
    Dim iContainer As Object
    Dim iContainer_Prev As Object
    Dim iStr As String
    Dim iParent As Object
    Dim iVisible As Boolean
    
    On Error Resume Next
    Set iParent = UserControl.Parent
    Set iControls = iParent.Controls
    If iControls Is Nothing Then ' this parent doesn't have a controls collection
        Set iControls = UserControl.ContainedControls ' let's do it just with the contained controls then
    End If
    For Each iCtl In iControls
        Set iContainer_Prev = Nothing
        Set iContainer = Nothing
        Set iContainer = iCtl.Container
        Do Until iContainer Is Nothing
            If iContainer Is UserControl.Extender Then
                iVisible = False
                If Not (iContainer_Prev Is Nothing Or iContainer_Prev Is iParent) Then ' the control is contained in another control that is contained in the usercontrol
                    iVisible = iContainer_Prev.Left > -mLeftThresholdHided
                    If iVisible Or nInitialize Then
                        iStr = ControlName(iCtl)
                        mParentControlsTabStop.Add iCtl.TabStop, iStr
                        mParentControlsUseMnemonic.Add iCtl.UseMnemonic, iStr
                        iStr = ControlName(iContainer_Prev)
                        mContainedControlsThatAreContainers.Add iStr, iStr
                        If nInitialize Then
                            If Not iVisible Then
                                iCtl.TabStop = False
                                iCtl.UseMnemonic = False
                            End If
                        Else
                            iCtl.TabStop = False
                            iCtl.UseMnemonic = False
                        End If
                    End If
                Else ' the control is directly contained in the usercontrol
                    iVisible = iCtl.Left > -mLeftThresholdHided
                    If iVisible Or nInitialize Then
                        iStr = ControlName(iCtl)
                        mParentControlsTabStop.Add iCtl.TabStop, iStr
                        mParentControlsUseMnemonic.Add iCtl.UseMnemonic, iStr
                        If nInitialize Then
                            If Not iVisible Then
                                iCtl.TabStop = False
                                iCtl.UseMnemonic = False
                            End If
                        Else
                            iCtl.TabStop = False
                            iCtl.UseMnemonic = False
                        End If
                    End If
                End If
                Exit Do
            End If
            Set iContainer_Prev = iContainer
            Set iContainer = Nothing
            Set iContainer = iContainer_Prev.Container
        Loop
    Next
    mTabStopsInitialized = True
    Err.Clear
End Sub

Private Sub SubclassControlsPainting()
    Dim iSubclassTheControls As Boolean
    Dim iHwnd As Long
    Dim c As Long
    Dim iBKColor As Long
    Dim iControls As Object
    Dim iCtl As Control
    Dim iContainer As Object
    Dim iContainer_Prev As Object
    Dim iParent As Object
    Dim iVisible As Boolean
    Dim iTabBackColor As Long
    Dim iCancel As Boolean
    Dim iClassNotHandled As Boolean
    Dim iCtlTypeName As String
    Dim iStr As String
    
  '  If Not mAmbientUserMode Then Exit Sub
    If Not mSubclassed Then Exit Sub
    If Not mUserControlShown Then
        tmrSubclassControls.Enabled = True
        Exit Sub
    End If
    
    mSubclassControlsPaintingPending = False
    
    iSubclassTheControls = mVisualStyles And (IsAppThemeEnabled Or mForceVisualStyles) And mChangeControlsBackColor
    If mSubclassedControlsForPaintingHwnds.Count > 0 Then
        For c = 1 To mSubclassedControlsForPaintingHwnds.Count
            iHwnd = mSubclassedControlsForPaintingHwnds(c)
            DetachMessage Me, iHwnd, WM_PAINT
            DetachMessage Me, iHwnd, WM_MOVE
            If Not iSubclassTheControls And mRepaintSubclassedControls Then
                ' redraw the control
                RedrawWindow iHwnd, ByVal 0&, 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_INTERNALPAINT Or RDW_ALLCHILDREN
            End If
        Next c
        Set mSubclassedControlsForPaintingHwnds = New Collection
    End If
    
    If mSubclassedFramesHwnds.Count > 0 Then
        For c = 1 To mSubclassedFramesHwnds.Count
            iHwnd = mSubclassedFramesHwnds(c)
            DetachMessage Me, iHwnd, WM_PRINTCLIENT
            DetachMessage Me, iHwnd, WM_MOUSELEAVE
        Next c
        Set mSubclassedFramesHwnds = New Collection
    End If
    
    If Not iSubclassTheControls Then
        mRepaintSubclassedControls = False
'        Exit Sub
    End If
    
    If mChangeControlsBackColor Then
        If Not mChangedControlsBackColor Then
            SetControlsBackColor mTabSelBackColor
            mChangedControlsBackColor = True
        End If
    End If
    
    If mShowDisabledState And (Not mEnabled) Then
        iTabBackColor = mTabSelBackColorDisabled
    Else
        iTabBackColor = mTabSelBackColor
    End If
    
    On Error Resume Next
    Set iParent = UserControl.Parent
    Set iControls = iParent.Controls
    If iControls Is Nothing Then ' this parent doesn't have a controls collection
        Set iControls = UserControl.ContainedControls ' let's do it just with the contained controls then
    End If
    For Each iCtl In iControls
        iCtlTypeName = TypeName(iCtl)
        iClassNotHandled = (iCtlTypeName = "ButtonEx") Or (iCtlTypeName = "ButtonExNoFocus")
        If Not iClassNotHandled Then
            Set iContainer_Prev = Nothing
            Set iContainer = Nothing
            Set iContainer = iCtl.Container
            If iContainer Is Nothing Then
                iHwnd = 0
                iHwnd = iCtl.hWnd
                If iHwnd <> 0 Then
                    Set iContainer = GetContainerByHwnd(iHwnd)
                End If
            End If
            Do Until iContainer Is Nothing
                If iContainer Is UserControl.Extender Then
                    iVisible = False
                    If Not (iContainer_Prev Is Nothing Or iContainer_Prev Is iParent) Then ' the control is contained in another control that is contained in the usercontrol
                        iVisible = iContainer_Prev.Left > -mLeftThresholdHided
                        If iVisible Then
                            iHwnd = 0
                            iHwnd = iCtl.hWnd
                            If iHwnd <> 0 Then
                                If iSubclassTheControls Then
                                    iBKColor = -1
                                    iBKColor = iCtl.BackColor
                                    If (iBKColor = iTabBackColor) Then
                                        iCancel = False
                                        If iContainer_Prev.Container Is UserControl.Extender Then
                                            iStr = iContainer_Prev.Name
                                            RaiseEvent ChangeControlBackColor(iStr, TypeName(iContainer_Prev), iCancel)
                                        End If
                                        If Not iCancel Then
                                            iCancel = False
                                            iStr = iCtl.Name
                                            RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                            If Not iCancel Then
                                                mSubclassedControlsForPaintingHwnds.Add iHwnd, CStr(iHwnd)
                                            End If
                                        End If
                                    End If
                                End If
                                If iCtlTypeName = "Frame" Then
                                    mSubclassedFramesHwnds.Add iHwnd, CStr(iHwnd)
                                End If
                            ElseIf iCtlTypeName = "Label" Then
                                If iCtl.BackStyle = 1 Then ' solid
                                    If iSubclassTheControls Then
                                        iBKColor = -1
                                        iBKColor = iCtl.BackColor
                                        If (iBKColor = iTabBackColor) Then
                                            iCancel = False
                                            If iContainer_Prev.Container Is UserControl.Extender Then
                                                iStr = iContainer_Prev.Name
                                                RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                            End If
                                            If Not iCancel Then
                                                iCancel = False
                                                iStr = iCtl.Name
                                                RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                                If Not iCancel Then
                                                    iCtl.BackStyle = 0 ' transparent
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else ' the control is directly contained in the usercontrol
                        iVisible = iCtl.Left > -mLeftThresholdHided
                        If iVisible Then
                            iHwnd = 0
                            iHwnd = iCtl.hWnd
                            If iHwnd <> 0 Then
                                If iSubclassTheControls Then
                                    iBKColor = -1
                                    iBKColor = iCtl.BackColor
                                    If (iBKColor = iTabBackColor) Then
                                        iCancel = False
                                        iStr = iCtl.Name
                                        RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                        If Not iCancel Then
                                            mSubclassedControlsForPaintingHwnds.Add iHwnd, CStr(iHwnd)
                                        End If
                                    End If
                                End If
                                If iCtlTypeName = "Frame" Then
                                    mSubclassedFramesHwnds.Add iHwnd, CStr(iHwnd)
                                End If
                            ElseIf iCtlTypeName = "Label" Then
                                If iCtl.BackStyle = 1 Then ' solid
                                    If iSubclassTheControls Then
                                        iBKColor = -1
                                        iBKColor = iCtl.BackColor
                                        If (iBKColor = iTabBackColor) Then
                                            iCancel = False
                                            If iContainer_Prev.Container Is UserControl.Extender Then
                                                iStr = iContainer_Prev.Name
                                                RaiseEvent ChangeControlBackColor(iStr, TypeName(iContainer_Prev), iCancel)
                                            End If
                                            If Not iCancel Then
                                                iCancel = False
                                                iStr = iCtl.Name
                                                RaiseEvent ChangeControlBackColor(iStr, TypeName(iCtl), iCancel)
                                                If Not iCancel Then
                                                    iCtl.BackStyle = 0 ' transparent
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    Exit Do
                End If
                Set iContainer_Prev = iContainer
                Set iContainer = Nothing
                Set iContainer = iContainer_Prev.Container
            Loop
        End If
    Next
    On Error GoTo 0
    
    
    For c = 1 To mSubclassedFramesHwnds.Count
        iHwnd = mSubclassedFramesHwnds(c)
        AttachMessage Me, iHwnd, WM_PRINTCLIENT
        AttachMessage Me, iHwnd, WM_MOUSELEAVE
    Next
    For c = 1 To mSubclassedControlsForPaintingHwnds.Count
        iHwnd = mSubclassedControlsForPaintingHwnds(c)
        AttachMessage Me, iHwnd, WM_PAINT
        AttachMessage Me, iHwnd, WM_MOVE
        If mRepaintSubclassedControls Then
            ' redraw the control
            RedrawWindow iHwnd, ByVal 0&, 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_INTERNALPAINT Or RDW_ALLCHILDREN
        End If
    Next c
    mRepaintSubclassedControls = False
    
End Sub

Private Function GetContainerByHwnd(nHwnd As Long) As Object
    Dim iParent As Object
    Dim iControls As Object
    Dim iCtl As Control
    Dim iHwndParent As Long
    Dim iHwnd As Long
    
    On Error Resume Next
    Set iParent = UserControl.Extender.Parent
    If iParent Is Nothing Then GoTo Exit_Function
    Set iControls = iParent.Controls
    If iControls Is Nothing Then GoTo Exit_Function
    
    iHwndParent = GetParent(nHwnd)
    
    For Each iCtl In iControls
        iHwnd = 0
        iHwnd = iCtl.hWnd
        If iHwnd = iHwndParent Then
            Set GetContainerByHwnd = iCtl
            GoTo Exit_Function
        End If
    Next
    
Exit_Function:
    Err.Clear
End Function

Private Function ControlIsContainer(nControlName As Variant) As Boolean
    Dim iStr As String
    
    On Error Resume Next
    iStr = mContainedControlsThatAreContainers(nControlName)
    ControlIsContainer = Err.Number = 0
    Err.Clear
End Function

Private Sub SetTabStopsToParentControlsContainedInControl(nContainer As Object)
    Dim iControls As Object
    Dim iCtl As Control
    Dim iContainer As Object
    Dim iStr As String
    Dim iObj As Object
    
    If nContainer Is Nothing Then Exit Sub
    On Error Resume Next
    Set iControls = GetContainedControlsInControlContainer(nContainer)
    If Not iControls Is Nothing Then
        For Each iCtl In iControls
            Set iContainer = Nothing
            Set iContainer = iCtl.Container
            Do Until iContainer Is Nothing
                If iContainer Is nContainer Then
                    iStr = ControlName(iCtl)
                    iCtl.TabStop = mParentControlsTabStop(iStr)
                    iCtl.UseMnemonic = mParentControlsUseMnemonic(iStr)
                    If TypeName(iCtl) = "ComboBox" Then
                        ' ComboBox fix
                        If iCtl.Style = vbComboDropdown Then
                            iCtl.SelLength = 0
                        End If
                    End If
                End If
                Set iObj = iContainer
                Set iContainer = Nothing
                Set iContainer = iObj.Container
            Loop
        Next
    End If
    Err.Clear
End Sub

Private Function GetContainedControlsInControlContainer(nContainer As Object) As Object
    Dim iControls As Object
    Dim iCtl As Control
    Dim iContainer As Object
    Dim iContainer_Prev As Object
    
    Set GetContainedControlsInControlContainer = New Collection
    
    If nContainer Is Nothing Then Exit Function
    On Error Resume Next
    Set iControls = UserControl.Parent.Controls
    If iControls Is Nothing Then GoTo Exit_Function
    
    For Each iCtl In iControls
        Set iContainer_Prev = Nothing
        Set iContainer = Nothing
        Set iContainer = iCtl.Container
        Do Until iContainer Is Nothing
            If iContainer Is nContainer Then
                GetContainedControlsInControlContainer.Add iCtl
            End If
            Set iContainer_Prev = iContainer
            Set iContainer = Nothing
            Set iContainer = iContainer_Prev.Container
        Loop
    Next
    
Exit_Function:
    Err.Clear
End Function

Private Function ControlName(nCtl As Object) As String
    Dim iIndex As Integer
    
    On Error GoTo NoIndex:
    ControlName = nCtl.Name
    iIndex = -1
    iIndex = nCtl.Index
    If iIndex >= 0 Then
        ControlName = ControlName & "(" & iIndex & ")"
    End If

NoIndex:
End Function

Private Function GetContainedControlByName(ByVal nControlName As String) As Object
    Dim iCtl As Object

    For Each iCtl In UserControl.ContainedControls
        If StrComp(nControlName, ControlName(iCtl), vbTextCompare) = 0 Then
            Set GetContainedControlByName = iCtl
            Exit For
        End If
    Next
End Function

Private Sub SetAccessKeys()
    Dim c As Long
    Dim iPos As Long
    Dim iChr As String
    Dim iAsc As Long
    Dim iAK As String
    
    mAccessKeys = ""
    iAK = ""
    
    For c = 0 To mTabs - 1
        iChr = ""
        If mTabData(c).Enabled And mTabData(c).Visible Then
            iPos = InStr(mTabData(c).Caption, "&")
            If iPos > 0 Then
                iChr = LCase(Mid$(mTabData(c).Caption, iPos + 1, 1))
                If (iChr <> "") Then
                    iAsc = Asc(iChr)
                    If Not (((iAsc > 47) And (iAsc < 58)) Or ((iAsc > 96) And (iAsc < 123))) Then
                        iChr = ""
                    End If
                End If
            End If
        End If
        iAK = iAK & iChr
        If iChr = "" Then iChr = Chr(0)
        mAccessKeys = mAccessKeys & iChr
    Next c
    UserControl.AccessKeys = iAK
    mAccessKeysSet = True
End Sub

Private Sub SetPicToUse(nTab As Long)
    Dim iTx As Single
    
    If mTabData(nTab).PicToUseSet Then Exit Sub
    
    iTx = Screen_TwipsPerPixelX
    If Not mTabData(nTab).Pic16 Is Nothing Then
        If iTx >= 15 Then ' 96 DPI
            Set mTabData(nTab).PicToUse = mTabData(nTab).Pic16
        ElseIf iTx >= 12 Then ' 120 DPI
            If Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = mTabData(nTab).Pic20
            Else
                Set mTabData(nTab).PicToUse = mTabData(nTab).Pic16
            End If
        ElseIf iTx >= 10 Then ' 144 DPI
            If Not mTabData(nTab).Pic24 Is Nothing Then
                Set mTabData(nTab).PicToUse = mTabData(nTab).Pic24
            ElseIf Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = mTabData(nTab).Pic20
            Else
                Set mTabData(nTab).PicToUse = mTabData(nTab).Pic16
            End If
        ElseIf iTx >= 7 Then ' 192 DPI
            Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 2)
        ElseIf iTx >= 6 Then
            If Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic20, 2)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 2)
            End If
        ElseIf iTx >= 5 Then
            If Not mTabData(nTab).Pic24 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic24, 2)
            ElseIf Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic20, 2)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 3)
            End If
        ElseIf iTx >= 4 Then  ' 289 to 360 DPI
            If Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic20, 3)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 4)
            End If
        ElseIf iTx >= 3 Then   ' 361 to 480 DPI
            If Not mTabData(nTab).Pic24 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic24, 3)
            ElseIf Not mTabData(nTab).Pic20 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic20, 4)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 6)
            End If
        ElseIf iTx >= 2 Then   ' 481 to 720 DPI
            If Not mTabData(nTab).Pic24 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic24, 5)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 8)
            End If
        Else ' greater than 720 DPI
            If Not mTabData(nTab).Pic24 Is Nothing Then
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic24, 10)
            Else
                Set mTabData(nTab).PicToUse = StretchPicNN(mTabData(nTab).Pic16, 16)
            End If
        End If
    Else
        If Not mTabData(nTab).Picture Is Nothing Then
            Set mTabData(nTab).PicToUse = mTabData(nTab).Picture
        Else
            Set mTabData(nTab).PicToUse = Nothing
        End If
    End If
    mTabData(nTab).PicToUseSet = True
End Sub

Private Function StretchPicNN(nPic As StdPicture, nFactor As Long) As StdPicture
    Dim iWidth As Long
    Dim iHeight As Long
    
    iWidth = pScaleX(nPic.Width, vbHimetric, vbPixels)
    iHeight = pScaleX(nPic.Height, vbHimetric, vbPixels)
    picAux.Width = iWidth * nFactor
    picAux.Height = iHeight * nFactor
    picAux.Cls
    
    picAux.PaintPicture nPic, 0, 0, picAux.Width, picAux.Height, 0, 0, iWidth, iHeight
    Set StretchPicNN = picAux.Image
    picAux.Cls
End Function

Private Function PictureToGrayScale(nPic As StdPicture) As StdPicture
    Dim iWidth As Long
    Dim iHeight As Long
    Dim X As Long
    Dim Y As Long
    Dim iColor As Long

    If nPic Is Nothing Then Exit Function
    
    iWidth = pScaleX(nPic.Width, vbHimetric, vbPixels)
    iHeight = pScaleX(nPic.Height, vbHimetric, vbPixels)
    picAux.Width = iWidth
    picAux.Height = iHeight
    picAux.Cls
    picAux2.Width = picAux.Width
    picAux2.Height = picAux.Height
    picAux2.Cls
    
    Set picAux.Picture = nPic

    For X = 0 To picAux.ScaleWidth - 1
        For Y = 0 To picAux.ScaleHeight - 1
            iColor = GetPixel(picAux.hDC, X, Y)
            If iColor <> mMaskColor Then
                iColor = ToGray(iColor)
            End If
            SetPixelV picAux2.hDC, X, Y, iColor
        Next Y
    Next X

    Set PictureToGrayScale = picAux2.Image
    picAux.Cls
    picAux2.Cls
End Function

Private Function ToGray(nColor As Long) As Long
    Dim iR As Long
    Dim iG As Long
    Dim iB As Long
    Dim iC As Long
    Dim iBlendDisablePicWithTabBackColor As Boolean
    
    iR = nColor And 255
    iG = (nColor \ 256) And 255
    iB = (nColor \ 65536) And 255
    iC = (0.2125 * iR + 0.7154 * iG + 0.0721 * iB)
    
    If mControlIsThemed Then
        iBlendDisablePicWithTabBackColor = mBlendDisablePicWithTabBackColor_Themed
    Else
        iBlendDisablePicWithTabBackColor = mBlendDisablePicWithTabBackColor_NotThemed
    End If
        
    If iBlendDisablePicWithTabBackColor Then
        If mControlIsThemed Then
            ToGray = RGB(iC / 255 * mThemedTabBodyBackColor_R * 0.7 + 88, iC / 255 * mThemedTabBodyBackColor_G * 0.7 + 88, iC / 255 * mThemedTabBodyBackColor_B * 0.7 + 88)
        Else
            ToGray = RGB(iC / 255 * mTabBackColor_R * 0.7 + 88, iC / 255 * mTabBackColor_G * 0.7 + 88, iC / 255 * mTabBackColor_B * 0.7 + 88)
        End If
    Else
        ToGray = RGB(iC * 0.6 + 90, iC * 0.6 + 90, iC * 0.6 + 90)
    End If

End Function

Private Sub ResetAllPicsDisabled()
    Dim t As Long
    
    For t = 0 To mTabs - 1
        mTabData(t).PicDisabledSet = False
    Next t
End Sub

Private Function MouseIsOverAContainedControl() As Boolean
    Dim iPt As POINTAPI
    Dim iSM As Long
    Dim iCtl As Control
    Dim iWidth As Long
    
    iSM = UserControl.ScaleMode
    UserControl.ScaleMode = vbTwips
    GetCursorPos iPt
    ScreenToClient mUserControlHwnd, iPt
    iPt.X = iPt.X * Screen_TwipsPerPixelX
    iPt.Y = iPt.Y * Screen_TwipsPerPixely
    
    On Error Resume Next
    For Each iCtl In UserControl.ContainedControls
        iWidth = -1
        iWidth = iCtl.Width
        If iWidth <> -1 Then
            If iCtl.Left <= iPt.X Then
                If iCtl.Left + iCtl.Width >= iPt.X Then
                    If iCtl.Top <= iPt.Y Then
                        If iCtl.Top + iCtl.Height >= iPt.Y Then
                            MouseIsOverAContainedControl = True
                            Err.Clear
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    Err.Clear
    UserControl.ScaleMode = iSM
End Function


Private Sub DrawDelayed()
    If mAmbientUserMode Then
        PostDrawMessage
    Else
        Draw
    End If
End Sub

Private Sub PostDrawMessage()
    If mCanPostDrawMessage Then
        If Not mDrawMessagePosted Then
            PostMessage mUserControlHwnd, WM_DRAW, 0&, 0&
            mDrawMessagePosted = True
        End If
    Else
        tmrDraw.Enabled = True
    End If
End Sub

Friend Property Get TabControlsNames(Index) As Object
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    Set TabControlsNames = mTabData(Index).Controls
End Property

Friend Property Set TabControlsNames(Index, nValue As Object)
    If (Index < 0) Or (Index >= mTabs) Then
        RaiseError 380, TypeName(Me) ' invalid property value
        Exit Property
    End If
    Set mTabData(Index).Controls = nValue
End Property

Friend Sub HideAllContainedControls()
    Dim iCtl As Control
    Dim c As Long
    Dim iHwnd As Long
    Dim iIsLine As Boolean
    
    If mUserControlTerminated Then Exit Sub
    
    If mSubclassedControlsForMoveHwnds.Count > 0 Then
        For c = 1 To mSubclassedControlsForMoveHwnds.Count
            iHwnd = mSubclassedControlsForMoveHwnds(c)
            DetachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
        Next c
        Set mSubclassedControlsForMoveHwnds = New Collection
    End If
    
    On Error Resume Next
    For Each iCtl In UserControl.ContainedControls
        iIsLine = TypeName(iCtl) = "Line"
        If iIsLine Then
            If iCtl.X1 > -mLeftThresholdHided Then
                iCtl.X1 = iCtl.X1 - mLeftShiftToHide
                iCtl.X2 = iCtl.X2 - mLeftShiftToHide
            End If
        Else
            If iCtl.Left > -mLeftThresholdHided Then
                iCtl.Left = iCtl.Left - mLeftShiftToHide
            End If
        End If
    Next
    Err.Clear
End Sub

Friend Sub MakeContainedControlsInSelTabVisible()
    Dim iCtl As Control
    Dim iCtlName As Variant
    Dim iHwnd As Long
    Dim c As Long
    Dim iIsLine As Boolean
    
    If mUserControlTerminated Then Exit Sub
    
    If mSubclassedControlsForMoveHwnds.Count > 0 Then
        For c = 1 To mSubclassedControlsForMoveHwnds.Count
            iHwnd = mSubclassedControlsForMoveHwnds(c)
            DetachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
        Next c
        Set mSubclassedControlsForMoveHwnds = New Collection
    End If
    
    On Error Resume Next
    For Each iCtlName In mTabData(mTabSel).Controls
        Set iCtl = GetContainedControlByName(iCtlName)
        If Not iCtl Is Nothing Then
            iIsLine = TypeName(iCtl) = "Line"
            If iIsLine Then
                If iCtl.X1 < -mLeftThresholdHided Then
                    iCtl.X1 = iCtl.X1 + mLeftShiftToHide
                    iCtl.X2 = iCtl.X2 + mLeftShiftToHide
                End If
            Else
                If iCtl.Left < -mLeftThresholdHided Then
                    iCtl.Left = iCtl.Left + mLeftShiftToHide
                End If
            End If
            If mAmbientUserMode And mSubclassed Then
                iHwnd = 0
                iHwnd = iCtl.hWnd
                If iHwnd <> 0 Then
                    mSubclassedControlsForMoveHwnds.Add iHwnd
                    AttachMessage Me, iHwnd, WM_WINDOWPOSCHANGING
                End If
            End If
        End If
    Next
    Err.Clear
End Sub

Private Sub CheckContainedControlsConsistency(Optional nCheckControlsThatChangedToArray As Boolean)
    Dim t As Long
    Dim iCCList As Collection
    Dim iAllCtInTabs As Collection
    Dim c As Long
    Dim iStr As String
    Dim iCtl As Control
    Dim iCtlName As Variant
    Dim iCtlName2 As Variant
    Dim iCtlsInTabsToRemove As Collection
    Dim iShowedNewControls As Boolean
    Dim iThereAreMissingControls As Boolean
    Dim iAuxFound As Boolean
    Dim iCtlsTypesAndRects As Collection
    Dim iAuxTypeAndRect_CtrlInTab As String
    Dim iAuxTypeAndRect_CC As String
    Dim iFound As Boolean
    Dim t2 As Long
    Dim c2 As Long
    Dim iListCtlsNowArrayToUpdateInfo As Collection
    Dim iIsLine As Boolean
    
    Set iCCList = New Collection
    For Each iCtl In UserControl.ContainedControls
        iStr = ControlName(iCtl)
        iCCList.Add iStr, iStr
    Next
    
    On Error Resume Next
    Set iAllCtInTabs = New Collection
    For t = 0 To mTabs - 1
        For c = 1 To mTabData(t).Controls.Count
            iStr = mTabData(t).Controls(c)
            iAllCtInTabs.Add iStr, iStr
        Next c
    Next t
    On Error GoTo 0
    
    iThereAreMissingControls = False
    For Each iCtlName In iAllCtInTabs
        iAuxFound = False
        For Each iCtlName2 In iCCList
            If iCtlName2 = iCtlName Then
                iAuxFound = True
                Exit For
            End If
        Next
        If Not iAuxFound Then
            iThereAreMissingControls = True
            If nCheckControlsThatChangedToArray Then
                iAuxFound = False
                For Each iCtlName2 In iCCList
                    If iCtlName2 = iCtlName & "(0)" Then
                        iAuxFound = True
                        Exit For
                    End If
                Next
                If iAuxFound Then
                    If iListCtlsNowArrayToUpdateInfo Is Nothing Then Set iListCtlsNowArrayToUpdateInfo = New Collection
                    iListCtlsNowArrayToUpdateInfo.Add iCtlName, iCtlName
                End If
            Else
                Exit For
            End If
        End If
    Next
    
    If iThereAreMissingControls Then
        If nCheckControlsThatChangedToArray Then
            If Not iListCtlsNowArrayToUpdateInfo Is Nothing Then
                For t = 0 To mTabs - 1
                    For c = 1 To mTabData(t).Controls.Count
                        iStr = mTabData(t).Controls(c)
                        iFound = False
                        For Each iCtlName In iListCtlsNowArrayToUpdateInfo
                            If iCtlName = iStr Then
                                iFound = True
                            End If
                        Next
                        If iFound Then
                            iStr = iStr & "(0)"
                            mTabData(t).Controls.Add iStr, iStr, c
                            mTabData(t).Controls.Remove (c + 1)
                        End If
                    Next c
                Next t
            End If
        Else
            ' This fixes SStab paste bug, read http://www.vbforums.com/showthread.php?871285&p=5359379&viewfull=1#post5359379
            Set iCtlsTypesAndRects = New Collection
            For Each iCtl In UserControl.ContainedControls
                iStr = ControlName(iCtl)
                iCtlsTypesAndRects.Add GetControlTypeAndRect(iStr), iStr
            Next
            
            For t = 0 To mTabs - 1
                For c = 1 To mTabData(t).Controls.Count
                    iStr = mTabData(t).Controls(c)
                    iAuxTypeAndRect_CtrlInTab = GetControlTypeAndRect(iStr)
                    If iAuxTypeAndRect_CtrlInTab = "-" Then ' if the control is not found it may have been en converted to an array
                        iAuxTypeAndRect_CtrlInTab = GetControlTypeAndRect(iStr & "(0)")
                    End If
                    For Each iCtlName In iCCList
                        iAuxTypeAndRect_CC = GetControlTypeAndRect(CStr(iCtlName))
                        If iAuxTypeAndRect_CC = iAuxTypeAndRect_CtrlInTab Then
                            iFound = False
                            For t2 = 0 To mTabs - 1
                                For c2 = 1 To mTabData(t).Controls.Count
                                    If mTabData(t).Controls(c) = iCtlName Then
                                        iFound = True
                                    End If
                                Next c2
                            Next t2
                            If Not iFound Then
                                mTabData(t).Controls.Add iCtlName, iCtlName, c
                                mTabData(t).Controls.Remove (c + 1)
                            End If
                        End If
                    Next
                Next c
            Next t
            
            On Error Resume Next
            Set iAllCtInTabs = New Collection
            For t = 0 To mTabs - 1
                For c = 1 To mTabData(t).Controls.Count
                    iStr = mTabData(t).Controls(c)
                    iAllCtInTabs.Add iStr, iStr
                Next c
            Next t
            On Error GoTo 0
        End If
    End If
    
    If nCheckControlsThatChangedToArray Then Exit Sub
    
    ' check if contained control is on any tab
    iShowedNewControls = False
    On Error Resume Next
    For Each iCtlName In iCCList
        iStr = ""
        iStr = iAllCtInTabs(iCtlName)
        If iStr = "" Then ' the control is not placed on any tab
            ' place it in the visible tab
            mTabData(mTabSel).Controls.Add iCtlName, iCtlName
            Set iCtl = GetContainedControlByName(iCtlName)
            iIsLine = TypeName(iCtl) = "Line"
            If iIsLine Then
                If iCtl.X1 <= -mLeftThresholdHided Then
                    iCtl.X1 = iCtl.X1 + mLeftShiftToHide
                    iCtl.X2 = iCtl.X2 + mLeftShiftToHide
                    iShowedNewControls = True
                End If
            Else
                If iCtl.Left <= -mLeftThresholdHided Then
                    iCtl.Left = iCtl.Left + mLeftShiftToHide
                    iShowedNewControls = True
                End If
            End If
        End If
    Next
    
    If iShowedNewControls Then
        mSubclassControlsPaintingPending = True
        mRepaintSubclassedControls = True
        SubclassControlsPainting
    End If
    
    ' now check the inverse: if there are controls in tabs but they don't exists
    Set iCtlsInTabsToRemove = New Collection
    For Each iCtlName In iAllCtInTabs
        iStr = ""
        iStr = iCCList(iCtlName)
        If iStr = "" Then ' the control doesn't exist
            iCtlsInTabsToRemove.Add iStr, iStr
        End If
    Next
    
    ' remove the controls that don't exists, if any
    If iCtlsInTabsToRemove.Count > 0 Then
        For t = 0 To mTabs - 1
            For Each iCtlName In mTabData(t).Controls
                iStr = ""
                iStr = iCtlsInTabsToRemove(iCtlName)
                If iStr <> "" Then ' the control name is in the list of controls to remove
                    mTabData(t).Controls.Remove iCtlName
                End If
            Next
        Next t
    End If
    Err.Clear
End Sub

Private Sub CheckIfContainedControlChangedToArray()
    CheckContainedControlsConsistency True
End Sub

Private Function GetControlTypeAndRect(iCtlName As String) As String
    Dim iCtl As Object
    Dim iSng As Long
    
    Set iCtl = GetParentControlByName(iCtlName)
    If Not iCtl Is Nothing Then
        On Error Resume Next
        GetControlTypeAndRect = TypeName(iCtl) & "."
        iSng = 0
        iSng = iCtl.Left
        GetControlTypeAndRect = GetControlTypeAndRect & CStr(iSng) & "."
        iSng = 0
        iSng = iCtl.Top
        GetControlTypeAndRect = GetControlTypeAndRect & CStr(iSng) & "."
        iSng = 0
        iSng = iCtl.Width
        GetControlTypeAndRect = GetControlTypeAndRect & CStr(iSng) & "."
        iSng = 0
        iSng = iCtl.Height
        GetControlTypeAndRect = GetControlTypeAndRect & CStr(iSng)
    Else
        GetControlTypeAndRect = "-"
    End If
End Function

Private Function GetParentControlByName(ByVal nControlName As String) As Object
    Dim iCtl As Object

    For Each iCtl In UserControl.Parent.Controls
        If StrComp(nControlName, ControlName(iCtl), vbTextCompare) = 0 Then
            Set GetParentControlByName = iCtl
            Exit For
        End If
    Next
End Function


Public Property Get ContainedControls() As VBRUN.ContainedControls
Attribute ContainedControls.VB_Description = "Returns a collection of the controls that were added to the control."
    Set ContainedControls = UserControl.ContainedControls
End Property

Private Sub RaiseError(ByVal Number As Long, Optional ByVal Source, Optional ByVal Description, Optional ByVal HelpFile, Optional ByVal HelpContext)
    If InIDE Then
        On Error Resume Next
        Err.Raise Number, Source, Description, HelpFile, HelpContext
        MsgBox "Error " & Err.Number & ". " & Err.Description, vbCritical
    Else
        Err.Raise Number, Source, Description, HelpFile, HelpContext
    End If
End Sub

Private Function InIDE() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        Err.Clear
        On Error Resume Next
        Debug.Print 1 / 0
        If Err.Number Then
            sValue = 1
        Else
            sValue = 2
        End If
        Err.Clear
    End If
    InIDE = (sValue = 1)
End Function

Private Function ControlHasFocus() As Boolean
    ControlHasFocus = mHasFocus And mFormIsActive
End Function

Private Sub RearrangeContainedControlsPositions()
    Dim iCtl As Control
    Dim iTabBodyStart As Single
    Dim iTabBodyStart_Prev As Single
    Dim iIsLine As Boolean
    
    If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
        iTabBodyStart = pScaleY(mTabBodyStart - 5, vbPixels, vbTwips)
    Else
        iTabBodyStart = pScaleX(mTabBodyStart - 5, vbPixels, vbTwips)
    End If
    If (mTabOrientation_Prev = ssTabOrientationTop) Or (mTabOrientation_Prev = ssTabOrientationBottom) Then
        iTabBodyStart_Prev = pScaleY(mTabBodyStart_Prev - 5, vbPixels, vbTwips)
    Else
        iTabBodyStart_Prev = pScaleX(mTabBodyStart_Prev - 5, vbPixels, vbTwips)
    End If
    
    On Error Resume Next
    If mTabOrientation = mTabOrientation_Prev Then
        For Each iCtl In UserControl.ContainedControls
            iIsLine = TypeName(iCtl) = "Line"
            If mTabOrientation = ssTabOrientationTop Then
                If iIsLine Then
                    iCtl.Y1 = iCtl.Y1 - iTabBodyStart_Prev + iTabBodyStart
                    iCtl.Y2 = iCtl.Y2 - iTabBodyStart_Prev + iTabBodyStart
                Else
                    iCtl.Top = iCtl.Top - iTabBodyStart_Prev + iTabBodyStart
                End If
            ElseIf mTabOrientation = ssTabOrientationBottom Then
                If iIsLine Then
                    iCtl.Y1 = iCtl.Y1 + iTabBodyStart_Prev - iTabBodyStart
                    iCtl.Y2 = iCtl.Y2 + iTabBodyStart_Prev - iTabBodyStart
                Else
                    iCtl.Top = iCtl.Top + iTabBodyStart_Prev - iTabBodyStart
                End If
            ElseIf mTabOrientation = ssTabOrientationLeft Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 - iTabBodyStart_Prev + iTabBodyStart
                    iCtl.X2 = iCtl.X2 - iTabBodyStart_Prev + iTabBodyStart
                Else
                    iCtl.Left = iCtl.Left - iTabBodyStart_Prev + iTabBodyStart
                End If
            ElseIf mTabOrientation = ssTabOrientationRight Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 + iTabBodyStart_Prev - iTabBodyStart
                    iCtl.X2 = iCtl.X2 + iTabBodyStart_Prev - iTabBodyStart
                Else
                    iCtl.Left = iCtl.Left + iTabBodyStart_Prev - iTabBodyStart
                End If
            End If
        Next
    Else
        For Each iCtl In UserControl.ContainedControls
            iIsLine = TypeName(iCtl) = "Line"
            If mTabOrientation_Prev = ssTabOrientationTop Then
                If iIsLine Then
                    iCtl.Y1 = iCtl.Y1 - iTabBodyStart_Prev
                    iCtl.Y2 = iCtl.Y2 - iTabBodyStart_Prev
                Else
                    iCtl.Top = iCtl.Top - iTabBodyStart_Prev
                End If
            ElseIf mTabOrientation_Prev = ssTabOrientationBottom Then
                '
            ElseIf mTabOrientation_Prev = ssTabOrientationLeft Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 - iTabBodyStart_Prev
                    iCtl.X2 = iCtl.X2 - iTabBodyStart_Prev
                Else
                    iCtl.Left = iCtl.Left - iTabBodyStart_Prev
                End If
            ElseIf mTabOrientation_Prev = ssTabOrientationRight Then
                '
            End If
        
            If mTabOrientation = ssTabOrientationTop Then
                If iIsLine Then
                    iCtl.Y1 = iCtl.Y1 + iTabBodyStart
                    iCtl.Y2 = iCtl.Y2 + iTabBodyStart
                Else
                    iCtl.Top = iCtl.Top + iTabBodyStart
                End If
            ElseIf mTabOrientation = ssTabOrientationBottom Then
                '
            ElseIf mTabOrientation = ssTabOrientationLeft Then
                If iIsLine Then
                    iCtl.X1 = iCtl.X1 + iTabBodyStart
                    iCtl.X2 = iCtl.X2 + iTabBodyStart
                Else
                    iCtl.Left = iCtl.Left + iTabBodyStart
                End If
            ElseIf mTabOrientation = ssTabOrientationRight Then
                '
            End If
        Next
    End If
    Err.Clear
End Sub

Public Property Get TabControls(nTab As Integer, Optional GetChilds As Boolean = True) As Collection
Attribute TabControls.VB_Description = "Returns a collection of the controls that are inside a tab."
    Dim iCtlName As Variant
    Dim iCtl As Control
    Dim iCtl2 As Control
    Dim iObj As Object
    
    If (nTab < 0) Or (nTab > (mTabs - 1)) Then
        RaiseError 5, TypeName(Me) ' Invalid procedure call or argument
        Exit Sub
    End If
    
    Set TabControls = New Collection
    
    If GetChilds Then
        If Not mTabStopsInitialized Then
            StoreControlsTabStop True
            mTabStopsInitialized = True
        End If
    End If
    

    For Each iCtlName In mTabData(nTab).Controls
        Set iCtl = GetContainedControlByName(iCtlName)
        If Not iCtl Is Nothing Then
            Set iObj = iCtl
            TabControls.Add iObj, iCtlName
            If GetChilds Then
                If ControlIsContainer(iCtlName) Then
                    For Each iCtl2 In GetContainedControlsInControlContainer(iCtl)
                        Set iObj = iCtl2
                        TabControls.Add iObj, iCtl2.Name
                    Next
                End If
            End If
        End If
    Next
    
End Property

Public Property Get EndOfTabs() As Single
Attribute EndOfTabs.VB_Description = "Returns and value that indicates where the last tab ends."
    EnsureDrawn
    If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
        EndOfTabs = FixRoundingError(ToContainerSizeX(mEndOfTabs, vbPixels))
    Else
        EndOfTabs = FixRoundingError(ToContainerSizeY(mEndOfTabs, vbPixels))
    End If
End Property

Public Property Get MinWidthNeeded() As Single
Attribute MinWidthNeeded.VB_Description = "Returns the minimun Width of the control needed to show all the tab captions in one line without cut."
    If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
        EnsureDrawn
        MinWidthNeeded = FixRoundingError(ToContainerSizeX(mMinSpaceNeeded, vbPixels))
    End If
End Property

Public Property Get MinHeightNeeded() As Single
Attribute MinHeightNeeded.VB_Description = "Returns the minimun Height of the control needed to show all the tab captions in one line without cut."
    If (mTabOrientation = ssTabOrientationLeft) Or (mTabOrientation = ssTabOrientationRight) Then
        EnsureDrawn
        MinHeightNeeded = FixRoundingError(ToContainerSizeY(mMinSpaceNeeded, vbPixels))
    End If
End Property


Public Property Get HandleHighContrastTheme() As Boolean
Attribute HandleHighContrastTheme.VB_Description = "When True (default setting), handles the system changes to High contrast theme automatically."
Attribute HandleHighContrastTheme.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
    HandleHighContrastTheme = mHandleHighContrastTheme
End Property

Public Property Let HandleHighContrastTheme(ByVal nValue As Boolean)
    If nValue <> mHandleHighContrastTheme Then
        mHandleHighContrastTheme = nValue
        PropertyChanged "HandleHighContrastTheme"
        If mHandleHighContrastTheme Then
            CheckHighContrastTheme
        End If
    End If
End Property


Private Function pScaleX(Width, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
    Select Case True
        Case ToScale = vbPixels
            Select Case FromScale
                Case vbCentimeters
                    pScaleX = Width * mDPIX / 2.54
                Case vbCharacters
                    pScaleX = Width / 1440 * mDPIX * 120
                Case vbHimetric
                    pScaleX = Width * mDPIX / 2540
                Case vbInches
                    pScaleX = Width * mDPIX
                Case vbMillimeters
                    pScaleX = Width * mDPIX / 25.4
                Case vbPixels
                    pScaleX = Width
                Case vbPoints
                    pScaleX = Width / 1440 * mDPIX * 20
                Case vbTwips
                    pScaleX = Width / 1440 * mDPIX
                Case Else
                    pScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
            End Select
        Case FromScale = vbPixels
            Select Case ToScale
                Case vbCentimeters
                    pScaleX = Width / mDPIX * 2.54
                Case vbCharacters
                    pScaleX = Width * 1440 / mDPIX / 120
                Case vbHimetric
                    pScaleX = Width / mDPIX * 2540
                Case vbInches
                    pScaleX = Width / mDPIX
                Case vbMillimeters
                    pScaleX = Width / mDPIX * 25.4
                Case vbPixels
                    pScaleX = Width
                Case vbPoints
                    pScaleX = Width * 1440 / mDPIX / 20
                Case vbTwips
                    pScaleX = Width * 1440 / mDPIX
                Case vbUser
                    pScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
                Case Else
                    pScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
            End Select
        Case Else
            pScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
    End Select
End Function

Private Function pScaleY(Height, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
    Select Case True
        Case ToScale = vbPixels
            Select Case FromScale
                Case vbCentimeters
                    pScaleY = Height * mDPIY / 2.54
                Case vbCharacters
                    pScaleY = Height / 1440 * mDPIY * 120
                Case vbHimetric
                    pScaleY = Height * mDPIY / 2540
                Case vbInches
                    pScaleY = Height * mDPIY
                Case vbMillimeters
                    pScaleY = Height * mDPIY / 25.4
                Case vbPixels
                    pScaleY = Height
                Case vbPoints
                    pScaleY = Height / 1440 * mDPIY * 20
                Case vbTwips
                    pScaleY = Height / 1440 * mDPIY
                Case Else
                    pScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
            End Select
        Case FromScale = vbPixels
            Select Case ToScale
                Case vbCentimeters
                    pScaleY = Height / mDPIY * 2.54
                Case vbCharacters
                    pScaleY = Height * 1440 / mDPIY / 120
                Case vbHimetric
                    pScaleY = Height / mDPIY * 2540
                Case vbInches
                    pScaleY = Height / mDPIY
                Case vbMillimeters
                    pScaleY = Height / mDPIY * 25.4
                Case vbPixels
                    pScaleY = Height
                Case vbPoints
                    pScaleY = Height * 1440 / mDPIY / 20
                Case vbTwips
                    pScaleY = Height * 1440 / mDPIY
                Case vbUser
                    pScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
                Case Else
                    pScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
            End Select
        Case Else
            pScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
    End Select
End Function

Private Sub SetDPI()
    Dim iDC As Long
    Dim iTx As Single
    Dim iTY As Single
    
    iDC = GetDC(0)
    mDPIX = GetDeviceCaps(iDC, LOGPIXELSX)
    mDPIY = GetDeviceCaps(iDC, LOGPIXELSY)
    ReleaseDC 0, iDC
    
    iTx = 1440 / mDPIX
    iTY = 1440 / mDPIY
    
    mXCorrection = iTx / Screen.TwipsPerPixelX
    mYCorrection = iTY / Screen.TwipsPerPixelY
    
    SetLeftShiftToHide Screen.TwipsPerPixelX
End Sub

Private Sub SetLeftShiftToHide(nTwipsPerPixel As Long)
    If nTwipsPerPixel >= 6 Then
        mLeftShiftToHide = 75000 ' compatible with original SSTab up to 250% DPI
        mLeftThresholdHided = 15000
    Else
        mLeftShiftToHide = nTwipsPerPixel * 16384 * 0.6 ' Windows has a limit on controls positions out of screen (in pixels), need to handle that for very hight DPI setting (> 250%) https://www.vbforums.com/showthread.php?888201
        If mLeftShiftToHide > 30000 Then
            mLeftThresholdHided = 15000
        Else
            mLeftThresholdHided = mLeftShiftToHide / 2
        End If
    End If
End Sub

Private Function Screen_TwipsPerPixelX() As Single
    Screen_TwipsPerPixelX = Screen.TwipsPerPixelX * mXCorrection
End Function

Private Function Screen_TwipsPerPixely() As Single
    Screen_TwipsPerPixely = Screen.TwipsPerPixelY * mYCorrection
End Function


Public Property Get Object() As Object
Attribute Object.VB_Description = "Returns the control instance without the extender."
    Set Object = Me
End Property

Private Function IsMsgBoxShown() As Boolean
    Dim iHwnd As Long
     
    Do Until IsWindowLocal(iHwnd)
        iHwnd = FindWindowEx(0&, iHwnd, "#32770", vbNullString)
        If iHwnd = 0 Then Exit Function
    Loop
    IsMsgBoxShown = True
End Function

Private Function IsWindowLocal(ByVal nHwnd As Long) As Boolean
    Dim iIdProcess As Long
    
    Call GetWindowThreadProcessId(nHwnd, iIdProcess)
    IsWindowLocal = (iIdProcess = GetCurrentProcessId())
End Function

Private Function IsHighContrastTheme() As Boolean
    Dim iHC As tagHIGHCONTRAST
    
    iHC.cbSize = Len(iHC)
    SystemParametersInfo SPI_GETHIGHCONTRAST, Len(iHC), iHC, 0
    IsHighContrastTheme = (iHC.dwFlags And HCF_HIGHCONTRASTON) = HCF_HIGHCONTRASTON
End Function

Private Sub CheckHighContrastTheme()
    Dim iAuxBool As Boolean
    
    If Not mAmbientUserMode Then Exit Sub
    If mHighContrastThemeOn <> IsHighContrastTheme Then
        iAuxBool = Not mHighContrastThemeOn
        If iAuxBool Then
            mHandleHighContrastTheme_OrigForeColor = ForeColor
            mHandleHighContrastTheme_OrigTabBackColor = TabBackColor
            mHandleHighContrastTheme_OrigTabSelForeColor = TabSelForeColor
            mHandleHighContrastTheme_OrigTabSelBackColor = TabSelBackColor
            ForeColor = vbButtonText
            TabBackColor = vbButtonFace
            TabSelForeColor = vbButtonText
            TabSelBackColor = vbButtonFace
            mHighContrastThemeOn = True
        Else
            mHighContrastThemeOn = False
            ForeColor = mHandleHighContrastTheme_OrigForeColor
            TabBackColor = mHandleHighContrastTheme_OrigTabBackColor
            TabSelForeColor = mHandleHighContrastTheme_OrigTabSelForeColor
            TabSelBackColor = mHandleHighContrastTheme_OrigTabSelBackColor
        End If
    End If
End Sub

Public Property Get LeftShiftToHide() As Long
Attribute LeftShiftToHide.VB_Description = "Returns the shift to the left in twips that it is using to hide the controls in not active tabs."
    LeftShiftToHide = mLeftShiftToHide
End Property


Public Property Get ContainedControlLeft(ByVal ControlName As String) As Single
Attribute ContainedControlLeft.VB_Description = "Returns/sets the left of the contained control whose name was provided by the ControlName parameter."
    Dim iCtl As Control
    Dim iFound As Boolean
    Dim iWithIndex As Boolean
    Dim iName As String
    Dim iIndex As Long
    
    ControlName = LCase$(ControlName)
    iWithIndex = InStr(ControlName, "(") > 0
    For Each iCtl In UserControl.ContainedControls
        iName = LCase$(iCtl.Name)
        If iWithIndex Then
            iIndex = -1
            On Error Resume Next
            iIndex = iCtl.Index
            On Error GoTo 0
            If iIndex <> -1 Then
                iName = iName & "(" & iIndex & ")"
            End If
        End If
        If iName = ControlName Then
            iFound = True
            Exit For
        End If
    Next
    If Not iFound Then
        RaiseError 1501, , "Control not found."
    Else
        If iCtl.Left < -mLeftThresholdHided Then
           ContainedControlLeft = iCtl.Left + mLeftShiftToHide + mPendingLeftShift
        Else
            ContainedControlLeft = iCtl.Left
        End If
    End If
End Property

Public Property Let ContainedControlLeft(ByVal ControlName As String, ByVal Left As Single)
    Dim iCtl As Control
    Dim iFound As Boolean
    
    Left = Left - mPendingLeftShift
    
    ControlName = LCase$(ControlName)
    For Each iCtl In UserControl.ContainedControls
        If LCase$(iCtl.Name) = ControlName Then
            iFound = True
            Exit For
        End If
    Next
    If Not iFound Then
        RaiseError 1501, , "Control not found."
    Else
        If iCtl.Left < -mLeftThresholdHided Then
            iCtl.Left = Left - mLeftShiftToHide
        Else
            iCtl.Left = Left
        End If
    End If
End Property


Private Sub SetAutoTabHeight()
    Dim iHeight As Single
    Dim t As Long
    Dim iPicHeight As Long
    
    If Not mAutoTabHeight Then Exit Sub
    
    If Not picAux2.Font Is mFont Then
        Set picAux2.Font = mFont
    End If
    
    iHeight = picAux2.ScaleY(picAux2.TextHeight("Atjq_"), picAux2.ScaleMode, vbHimetric)
'    mTabHeight = iHeight * 1.1 + pScaleY(4, vbPixels, vbHimetric)
    mTabHeight = iHeight * 1.02 + pScaleY(8, vbPixels, vbHimetric)
    
    For t = 0 To mTabs - 1

        If Not mTabData(t).PicToUseSet Then SetPicToUse t

        iPicHeight = 0
        If Not mTabData(t).PicToUse Is Nothing Then
            If (mTabOrientation = ssTabOrientationTop) Or (mTabOrientation = ssTabOrientationBottom) Then
                iPicHeight = mTabData(t).PicToUse.Height
            Else
                iPicHeight = mTabData(t).PicToUse.Width
            End If
        End If
        iPicHeight = iPicHeight + Screen.TwipsPerPixelY * 12
        If iPicHeight > mTabHeight Then
            mTabHeight = iPicHeight
        End If
    Next
    
    PropertyChanged "TabHeight"
End Sub


'Public Property Get Tab() As Integer
'    Tab = TabSel
'End Property
'
'Public Property Let Tab(ByVal nValue As Integer)
'    TabSel = nValue
'End Property

