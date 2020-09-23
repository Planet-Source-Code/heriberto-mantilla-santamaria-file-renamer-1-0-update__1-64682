VERSION 5.00
Begin VB.UserControl SOfficeButton 
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1110
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   PropertyPages   =   "SOfficeButton.ctx":0000
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   74
   ToolboxBitmap   =   "SOfficeButton.ctx":0035
End
Attribute VB_Name = "SOfficeButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************'
'*        All Rights Reserved © HACKPRO TM 2005        *'
'*******************************************************'
'*                   Version 1.0.3                     *'
'*******************************************************'
'* Control:       SOfficeButton                        *'
'*******************************************************'
'* Author:        Heriberto Mantilla Santamaría        *'
'*******************************************************'
'* Description:   This usercontrol simulates a Office  *'
'*                Button.                              *'
'*                                                     *'
'*                This button is based on the origi-   *'
'*                nal code of fred.cpp, please see     *'
'*                the [CodeId = 56053].                *'
'*                                                     *'
'*                Also many thanks to Paul Caton for   *'
'*                it's spectacular self-subclassing    *'
'*                usercontrol template, please see     *'
'*                the [CodeId = 54117].                *'
'*******************************************************'
'* Started on:    Sunday, 09-jan-2005.                 *'
'*******************************************************'
'* Release date:  Monday, 18-jul-2005.                 *'
'*******************************************************'
'*                                                     *'
'* Note:     Comments, suggestions, doubts or bug      *'
'*           reports are wellcome to these e-mail      *'
'*           addresses:                                *'
'*                                                     *'
'*                  heri_05-hms@mixmail.com or         *'
'*                  hcammus@hotmail.com                *'
'*                                                     *'
'*        Please rate my work on this control.         *'
'*    That lives the Soccer and the América of Cali    *'
'*             Of Colombia for the world.              *'
'*******************************************************'
'*        All Rights Reserved © HACKPRO TM 2005        *'
'*******************************************************'
Option Explicit

'* Private Types.
 Private Type RECT
  xLeft    As Long
  xTop     As Long
  xRight   As Long
  xBottom  As Long
 End Type
 
'*******************************************************'
'*                Subclasser Declarations              *'
'*                                                     *'
'* Author: Paul Caton.                                 *'
'* Mail:   Paul_Caton@hotmail.com                      *'
'* Web:    None                                        *'
'*******************************************************'
 
 Private Const ALL_MESSAGES          As Long = -1
 Private Const GMEM_FIXED            As Long = 0
 Private Const GWL_WNDPROC           As Long = -4
 Private Const PATCH_04              As Long = 88
 Private Const PATCH_05              As Long = 93
 Private Const PATCH_08              As Long = 132
 Private Const PATCH_09              As Long = 137
 Private Const WM_MOUSEMOVE          As Long = &H200
 Private Const WM_MOUSELEAVE         As Long = &H2A3
 Private Const WM_SYSCOLORCHANGE     As Long = &H15
 Private Const WM_THEMECHANGED       As Long = &H31A
 
 Private Type tSubData
  hWnd                               As Long
  nAddrSub                           As Long
  nAddrOrig                          As Long
  nMsgCntA                           As Long
  nMsgCntB                           As Long
  aMsgTblA()                         As Long
  aMsgTblB()                         As Long
 End Type

 Private Enum eMsgWhen
  MSG_AFTER = 1
  MSG_BEFORE = 2
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE
 End Enum
 
 Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
 End Enum

 Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                             As Long
  dwFlags                            As TRACKMOUSEEVENT_FLAGS
  hWndTrack                          As Long
  dwHoverTime                        As Long
 End Type

 Private bTrack                      As Boolean
 Private bTrackUser32                As Boolean
 Private isInCtrl                    As Boolean
 Private sc_aSubData()               As tSubData

 Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
 Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
 Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
 Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
 Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
 Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
 Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
 Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
 Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
 
 Public Event MouseEnter()
 Public Event MouseLeave()
'*******************************************************'

'*******************************************************'
'*                     Tool Tip Class                  *'
'*                                                     *'
'* Author: Mark Mokoski                                *'
'* Mail: markm@cmtelephone.com                         *'
'* Web:  www.rjillc.com                                *'
'*******************************************************'

 '******************************************************
 '* API Functions.                                     *
 '******************************************************
 Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
 Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
 Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
 
 '******************************************************
 '* Constants.                                         *
 '******************************************************
   
 '* Windows API Constants.
 Private Const CW_USEDEFAULT = &H80000000
 Private Const HWND_TOPMOST = -1
 Private Const SWP_NOACTIVATE = &H10
 Private Const SWP_NOMOVE = &H2
 Private Const SWP_NOSIZE = &H1
 Private Const WM_USER = &H400

 '* Tooltip Window Constants.
 Private Const TTF_CENTERTIP = &H2
 Private Const TTF_SUBCLASS = &H10
 Private Const TTM_ACTIVATE = (WM_USER + 1)
 Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
 Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
 Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
 Private Const TTM_SETTITLE = (WM_USER + 32)
 Private Const TTM_ADDTOOLA = (WM_USER + 4)
 Private Const TTS_ALWAYSTIP = &H1
 Private Const TTS_BALLOON = &H40
 Private Const TTS_NOPREFIX = &H2

 '* Tool Tip Icons.
 Private Const TTI_ERROR                   As Long = 3
 Private Const TTI_INFO                    As Long = 1
 Private Const TTI_NONE                    As Long = 0
 Private Const TTI_WARNING                 As Long = 2
 
 '* Tool Tip API Class.
 Private Const TOOLTIPS_CLASSA = "tooltips_class32"

 '******************************************************
 '* Types.                                             *
 '******************************************************

 '* Tooltip Window Types.
 Private Type TOOLINFO
  lSize                             As Long
  lFlags                            As Long
  lhWnd                             As Long
  lId                               As Long
  lpRect                            As RECT
  hInstance                         As Long
  lpStr                             As String
  lParam                            As Long
 End Type

 '******************************************************
 '* Local Class variables and Data .                   *
 '******************************************************

 '* Local variables to hold property values.
 Private ToolActive                        As Boolean
 Private ToolBackColor                     As Long
 Private ToolCentered                      As Boolean
 Private ToolForeColor                     As Long
 Private ToolIcon                          As ToolIconType
 Private TOOLSTYLE                         As ToolStyleEnum
 Private ToolText                          As String
 Private ToolTitle                         As String

 '* Private Data for Class.
 Private m_ltthWnd                         As Long
 Private TI                                As TOOLINFO
 
 Public Enum ToolIconType
  TipNoIcon = TTI_NONE            '= 0
  TipIconInfo = TTI_INFO          '= 1
  TipIconWarning = TTI_WARNING    '= 2
  TipIconError = TTI_ERROR        '= 3
 End Enum

 Public Enum ToolStyleEnum
  StyleStandard = 0
  StyleBalloon = 1
 End Enum
'*******************************************************'

 '* Private Types.
 Private Type POINTAPI
  X      As Long
  Y      As Long
 End Type
  
 '* Private Enum's.
 Public Enum OfficeAlign
  ACenter = &H0
  ALeft = &H1
  ARight = &H2
  ATop = &H3
  ABottom = &H4
 End Enum
 
 Public Enum OfficeState
  OfficeNormal = &H0
  OfficeHighLight = &H1
  OfficeHot = &H2
  OfficeDisabled = &H3
 End Enum
 
 Public Enum ShapeBorder
  Rectangle = &H0
  [Round Rectangle] = &H1
 End Enum
  
 '* Private variables.
 Private g_Font           As StdFont
 Private isAutoSizePic    As Boolean
 Private isBackColor      As OLE_COLOR
 Private isBorderColor    As OLE_COLOR
 Private isButtonShape    As ShapeBorder
 Private isCaption        As String
 Private isDisabledColor  As OLE_COLOR
 Private isEnabled        As Boolean
 Private isFocus          As Boolean
 Private isFontAlign      As OfficeAlign
 Private isForeColor      As OLE_COLOR
 Private isHeight         As Long
 Private isHighLightColor As OLE_COLOR
 Private isHotColor       As OLE_COLOR
 Private isHotTitle       As Boolean
 Private isMultiLine      As Boolean
 Private isPicture        As StdPicture
 Private isPictureAlign   As OfficeAlign
 Private isPictureSize    As Integer
 Private isSetBorder      As Boolean
 Private isSetBorderH     As Boolean
 Private isSetGradient    As Boolean
 Private isSetHighLight   As Boolean
 Private isShadowText     As Boolean
 Private isShowFocus      As Boolean
 Private isState          As OfficeState
 Private isSystemColor    As Boolean
 Private isTxtRect        As RECT
 Private isWidth          As Long
 Private isXPos           As Integer
 Private isYPos           As Integer
 Private m_bGrayIcon      As Boolean
 Private RectButton       As RECT
 
 '* Private Constants.
 Private Const defBackColor      As Long = vbButtonFace
 Private Const defBorderColor    As Long = vbHighlight
 Private Const defDisabledColor  As Long = vbGrayText
 Private Const defForeColor      As Long = vbButtonText
 Private Const defHighLightColor As Long = vbHighlight
 Private Const defHotColor       As Long = vbHighlight
 Private Const defShape          As Integer = &H0
 Private Const DSS_DISABLED      As Long = &H20
 Private Const DSS_MONO          As Long = &H80
 Private Const DSS_NORMAL        As Long = &H0
 Private Const DST_BITMAP        As Long = &H4
 Private Const DST_ICON          As Long = &H3
 Private Const DT_BOTTOM         As Long = &H8
 Private Const DT_CENTER         As Long = &H1
 Private Const DT_LEFT           As Long = &H0
 Private Const DT_RIGHT          As Long = &H2
 Private Const DT_SINGLELINE     As Long = &H20
 Private Const DT_TOP            As Long = &H0
 Private Const DT_VCENTER        As Long = &H4
 Private Const DT_WORDBREAK      As Long = &H10
 Private Const DT_WORD_ELLIPSIS  As Long = &H40000
 Private Const PS_SOLID          As Long = 0
 Private Const SW_SHOWNORMAL     As Long = 1
 Private Const Version           As String = "SOfficeButon 1.0.3 By HACKPRO TM"
 
 '* API's Windows Call.
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
 Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
 Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
 Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
 Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
 Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
 Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
 Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
 
 '* Public Events.
 Public Event Click()
Attribute Click.VB_MemberFlags = "200"
 Public Event ChangedTheme()
 
 '* For Create GrayIcon --> MArio Florez.
 Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
 Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
 Private Declare Function CreateIconIndirect Lib "user32.dll" (ByRef piconinfo As ICONINFO) As Long
 Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
 Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
 Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
 Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
 Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
 Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
 Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
 
 ' Type - GetObjectAPI.lpObject
 Private Type BITMAP
  bmType       As Long    'LONG   // Specifies the bitmap type. This member must be zero.
  bmWidth      As Long    'LONG   // Specifies the width, in pixels, of the bitmap. The width must be greater than zero.
  bmHeight     As Long    'LONG   // Specifies the height, in pixels, of the bitmap. The height must be greater than zero.
  bmWidthBytes As Long    'LONG   // Specifies the number of bytes in each scan line. This value must be divisible by 2, because Windows assumes that the bit values of a bitmap form an array that is word aligned.
  bmPlanes     As Integer 'WORD   // Specifies the count of color planes.
  bmBitsPixel  As Integer 'WORD   // Specifies the number of bits required to indicate the color of a pixel.
  bmBits       As Long    'LPVOID // Points to the location of the bit values for the bitmap. The bmBits member must be a long pointer to an array of character (1-byte) values.
 End Type

 ' Type - CreateIconIndirect / GetIconInfo
 Private Type ICONINFO
  fIcon    As Long 'BOOL    // Specifies whether this structure defines an icon or a cursor. A value of TRUE specifies an icon; FALSE specifies a cursor.
  xHotspot As Long 'DWORD   // Specifies the x-coordinate of a cursor’s hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
  yHotspot As Long 'DWORD   // Specifies the y-coordinate of the cursor’s hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
  hbmMask  As Long 'HBITMAP // Specifies the icon bitmask bitmap. If this structure defines a black and white icon, this bitmask is formatted so that the upper half is the icon AND bitmask and the lower half is the icon XOR bitmask. Under this condition, the height should be an even multiple of two. If this structure defines a color icon, this mask only defines the AND bitmask of the icon.
  hbmColor As Long 'HBITMAP // Identifies the icon color bitmap. This member can be optional if this structure defines a black and white icon. The AND bitmask of hbmMask is applied with the SRCAND flag to the destination; subsequently, the color bitmap is applied (using XOR) to the destination by using the SRCINVERT flag.
 End Type

'* ========================================================================================================
'*  Subclass handler - MUST be the first Public routine in this file. That includes public properties also
'* ========================================================================================================
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
 '* Parameters:
 '*  bBefore  - Indicates whether the the message is _
                being processed before or after the _
                default handler - only really needed _
                if a message is set to callback both _
                before & after.
 '*  bHandled - Set this variable to True in a before _
                callback to prevent the message being _
                subsequently processed by the default _
                handler... and if set, an after _
                callback.
 '*  lReturn  - Set this variable as per your intentions _
                and requirements, see the MSDN _
                documentation for each individual _
                message value.
 '*  hWnd     - The window handle.
 '*  uMsg     - The message number.
 '*  wParam   - Message related data.
 '*  lParam   - Message related data.
 '* Notes: _
     If you really know what youre doing, it's possible _
     to change the values of the hWnd, uMsg, wParam and _
     lParam parameters in a before callback so that _
     different values get passed to the default _
     handler... and optionaly, the after callback.
 Select Case uMsg
  Case WM_MOUSEMOVE
   If (isSetHighLight = False) Then Exit Sub
   If Not (isInCtrl = True) Then
    isInCtrl = True
    Call TrackMouseLeave(lng_hWnd)
    Call Refresh(OfficeHighLight)
    Call UpDate
    RaiseEvent MouseEnter
   End If
  Case WM_MOUSELEAVE
   If (isSetHighLight = False) Then Exit Sub
   isInCtrl = False
   Call Refresh(OfficeNormal)
   RaiseEvent MouseLeave
  Case WM_THEMECHANGED, WM_SYSCOLORCHANGE
   Call UserControl_Resize
   RaiseEvent ChangedTheme
 End Select
End Sub

'*******************************************************'
'* Public Properties.                                  *'
'*******************************************************'
Public Property Get AutoSizePicture() As Boolean
 AutoSizePicture = isAutoSizePic
End Property

'* English: Adjusts the control to the picture size.
Public Property Let AutoSizePicture(ByVal TheAutoSize As Boolean)
 '* Ajusta el control al tamaño de la imagen.
 isAutoSizePic = TheAutoSize
 Call PropertyChanged("AutoSizePicture")
 Call Refresh(isState)
End Property

Public Property Get BackColor() As OLE_COLOR
 BackColor = isBackColor
End Property

'* English: Returns/Sets the background color used to display text and graphics in an object.
Public Property Let BackColor(ByVal theColor As OLE_COLOR)
 '* Devuelve ó establece el color del Usercontrol.
 isBackColor = ConvertSystemColor(theColor)
 Call PropertyChanged("BackColor")
 Call Refresh(isState)
End Property

Public Property Get BorderColor() As OLE_COLOR
 BorderColor = isBorderColor
End Property

'* English: Returns/Sets the color of border of the Object.
Public Property Let BorderColor(ByVal theColor As OLE_COLOR)
 '* Devuelve ó establece el color del borde del objeto.
 isBorderColor = ConvertSystemColor(theColor)
 Call PropertyChanged("BorderColor")
 If (isSetBorder = True) Then Call Refresh(isState)
End Property

Public Property Get ButtonShape() As ShapeBorder
 ButtonShape = isButtonShape
End Property

'* English: Returns/Sets the type of border of the control.
Public Property Let ButtonShape(ByVal theButtonShape As ShapeBorder)
 '* Devuelve ó establece el tipo de borde del botón.
 isButtonShape = theButtonShape
 Call PropertyChanged("ButtonShape")
 If (isSetBorder = True) Then Call Refresh(isState)
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
 Caption = isCaption
End Property

'* English: Returns/Sets "Caption" property.
Public Property Let Caption(ByVal TheCaption As String)
 '* Devuelve ó establece el texto del Objeto.
 isCaption = TheCaption
 Call SetAccessKey(isCaption)
 Call PropertyChanged("Caption")
 Call Refresh(isState)
End Property

Public Property Get CaptionAlign() As OfficeAlign
 CaptionAlign = isFontAlign
End Property

'* English: Returns/Sets alignment of the text.
Public Property Let CaptionAlign(ByVal theAlign As OfficeAlign)
 '* Devuelve ó establece la alineación del texto.
 isFontAlign = theAlign
 Call PropertyChanged("CaptionAlign")
 Call Refresh(isState)
End Property

Public Property Get DisabledColor() As OLE_COLOR
 DisabledColor = isDisabledColor
End Property

'* English: Returns/Sets the color of the disabled text.
Public Property Let DisabledColor(ByVal theColor As OLE_COLOR)
 '* Devuelve ó establece el color del texto deshabilitado.
 isDisabledColor = ConvertSystemColor(theColor)
 Call PropertyChanged("DisabledColor")
 Call Refresh(isState)
End Property

Public Property Get Enabled() As Boolean
 Enabled = isEnabled
End Property

'* English: Returns/Sets the Enabled property of the control.
Public Property Let Enabled(ByVal TheEnabled As Boolean)
 '* Devuelve ó establece si el Usercontrol esta habilitado ó deshabilitado.
 isEnabled = TheEnabled
 UserControl.Enabled = isEnabled
 Call PropertyChanged("Enabled")
 If (isEnabled = True) Then
  isState = OfficeNormal
 Else
  isState = OfficeDisabled
 End If
 Call Refresh(isState)
End Property

Public Property Get Font() As StdFont
 Set Font = g_Font
End Property

'* English: Returns/Sets the Font of the control.
Public Property Set Font(ByVal New_Font As StdFont)
 '* Devuelve ó establece el tipo de fuente del texto.
On Error Resume Next
 With g_Font
  .Name = New_Font.Name
  .Size = New_Font.Size
  .Bold = New_Font.Bold
  .Italic = New_Font.Italic
  .Underline = New_Font.Underline
  .Strikethrough = New_Font.Strikethrough
 End With
 Call PropertyChanged("Font")
 Call Refresh(isState)
End Property

Public Property Get ForeColor() As OLE_COLOR
 ForeColor = isForeColor
End Property

'* English: Use this color for drawing Normal Font.
Public Property Let ForeColor(ByVal theColor As OLE_COLOR)
 '* Devuelve ó establece el color de la fuente.
 isForeColor = ConvertSystemColor(theColor)
 Call PropertyChanged("ForeColor")
 Call Refresh(isState)
End Property

'* English: Control Version.
Public Property Get GetControlVersion() As String
 '* Español: Version del Control.
 GetControlVersion = Version & " © " & Year(Now)
End Property

Public Property Let GrayIcon(ByVal bGrayIcon As Boolean)
 m_bGrayIcon = bGrayIcon
 Call PropertyChanged("GrayIcon")
 Call Refresh
End Property

Public Property Get GrayIcon() As Boolean
 GrayIcon = m_bGrayIcon
End Property

Public Property Get HighLightColor() As OLE_COLOR
 HighLightColor = isHighLightColor
End Property

'* English: Use this color for drawing.
Public Property Let HighLightColor(ByVal theColor As OLE_COLOR)
 '* Color de fondo cuando el mouse pasa sobre el Objeto.
 isHighLightColor = ConvertSystemColor(theColor)
 Call PropertyChanged("HighLightColor")
 Call Refresh(isState)
End Property

Public Property Get HotColor() As OLE_COLOR
 HotColor = isHotColor
End Property

'* English: Use this color for drawing.
Public Property Let HotColor(ByVal theColor As OLE_COLOR)
 '* Color de fondo cuando se tiene presionado el Objeto.
 isHotColor = ConvertSystemColor(theColor)
 Call PropertyChanged("HotColor")
 Call Refresh(isState)
End Property

Public Property Get HotTitle() As Boolean
 HotTitle = isHotTitle
End Property

'* English: Use this color for drawing.
Public Property Let HotTitle(ByVal theTitle As Boolean)
 '* Color de fondo cuando se tiene presionado el Objeto.
 isHotTitle = theTitle
 Call PropertyChanged("HotTitle")
End Property

'* English: Returns a handle to the control.
Public Property Get hWnd() As Long
 '* Devuelve el controlador del control.
 hWnd = UserControl.hWnd
End Property

Public Property Get MouseIcon() As StdPicture
 Set MouseIcon = UserControl.MouseIcon
End Property

'* English: Sets a custom mouse icon.
Public Property Set MouseIcon(ByVal MouseIcon As StdPicture)
 '* Devuelve ó establece un icono de mouse personalizado.
 Set UserControl.MouseIcon = MouseIcon
 Call PropertyChanged("MouseIcon")
End Property

Public Property Get MousePointer() As MousePointerConstants
 MousePointer = UserControl.MousePointer
End Property

'* English: Returns/Sets the type of mouse pointer displayed when over part of an object.
Public Property Let MousePointer(ByVal MousePointer As MousePointerConstants)
 '* Devuelve ó establece el tipo de puntero a mostrar cuando el mouse pase sobre el objeto.
 UserControl.MousePointer = MousePointer
 Call PropertyChanged("MousePointer")
End Property

Public Property Get MultiLine() As Boolean
 MultiLine = isMultiLine
End Property

'* English: Returns/Sets if the text is shown in multiple lines.
Public Property Let MultiLine(ByVal theMultiLine As Boolean)
 '* Devuelve ó establece si el texto se muestra en múltiples líneas.
 isMultiLine = theMultiLine
 Call PropertyChanged("MultiLine")
 Call Refresh(isState)
End Property

Public Property Get Picture() As StdPicture
 Set Picture = isPicture
End Property

'* English: Returns/Sets the image of the control.
Public Property Set Picture(ByVal thePicture As StdPicture)
 '* Devuelve ó establece la imagen del control.
 Set isPicture = thePicture
 Call PropertyChanged("Picture")
 Call Refresh(isState)
End Property

Public Property Get PictureAlign() As OfficeAlign
 PictureAlign = isPictureAlign
End Property

'* English: Returns/Sets the alignment of the image.
Public Property Let PictureAlign(ByVal theAlign As OfficeAlign)
 '* Devuelve ó establece la alineación de la imagen.
 isPictureAlign = theAlign
 Call PropertyChanged("PictureAlign")
 Call Refresh(isState)
End Property

Public Property Get PictureSize() As Integer
 PictureSize = isPictureSize
End Property

'* English: Returns/Sets the picture size.
Public Property Let PictureSize(ByVal theSize As Integer)
 '* Devuelve ó establece el tamaño de la imagen.
 isPictureSize = theSize
 Call PropertyChanged("PictureSize")
 Call Refresh(isState)
End Property

Public Property Get SetBorder() As Boolean
 SetBorder = isSetBorder
End Property

'* English: Returns/Sets if it's always shown the border.
Public Property Let SetBorder(ByVal theSetBorder As Boolean)
 '* Devuelve ó establece si se muestra siempre un borde.
 isSetBorder = theSetBorder
 Call PropertyChanged("SetBorder")
 Call Refresh(isState)
End Property

Public Property Get SetBorderH() As Boolean
 SetBorderH = isSetBorderH
End Property

'* English: Returns/Sets if it's always shown the Hot border.
Public Property Let SetBorderH(ByVal theSetBorderH As Boolean)
 '* Devuelve ó establece si se muestra siempre un borde.
 isSetBorderH = theSetBorderH
 Call PropertyChanged("SetBorderH")
End Property

Public Property Get SetGradient() As Boolean
 SetGradient = isSetGradient
End Property

'* English: Returns/Sets if the background is gradient.
Public Property Let SetGradient(ByVal theSetGradient As Boolean)
 '* Devuelve ó establece si el fondo es en degradado.
 isSetGradient = theSetGradient
 Call PropertyChanged("SetGradient")
 Call Refresh(isState)
End Property

Public Property Get SetHighLight() As Boolean
 SetHighLight = isSetHighLight
End Property

'* English: Returns/Sets if the background change is shown.
Public Property Let SetHighLight(ByVal theSetHighLight As Boolean)
 '* Devuelve ó establece si se muestra el cambio de fondo.
 isSetHighLight = theSetHighLight
 Call PropertyChanged("SetHighLight")
End Property

Public Property Get ShadowText() As Boolean
 ShadowText = isShadowText
End Property

'* English: Returns/Sets if a shadow is shown in the text of the button.
Public Property Let ShadowText(ByVal theShadowText As Boolean)
 '* Devuelve ó establece si se muestra una sombra en el texto del botón.
 isShadowText = theShadowText
 Call PropertyChanged("ShadowText")
End Property

Public Property Get ShowFocus() As Boolean
 ShowFocus = isShowFocus
End Property

'* English: Do you want to show the focus?
Public Property Let ShowFocus(ByVal theFocus As Boolean)
 '* Permite ver el enfoque del control.
 isShowFocus = theFocus
 Call PropertyChanged("ShowFocus")
End Property

Public Property Get SystemColor() As Boolean
 SystemColor = isSystemColor
End Property

'* English: Take the system color.
Public Property Let SystemColor(ByVal theSystemColor As Boolean)
 '* Toma los colores del Sistema.
 isSystemColor = theSystemColor
 Call PropertyChanged("SystemColor")
 Call Refresh(isState)
End Property

Public Property Get XPosPicture() As Integer
 XPosPicture = isXPos
End Property

'* English: Returns/Sets the Position X of the image.
Public Property Let XPosPicture(ByVal theXPos As Integer)
 '* Devuelve ó establece la Posición X de la imagen.
 isXPos = theXPos
 Call PropertyChanged("XPosPicture")
 Call Refresh(isState)
End Property

Public Property Get YPosPicture() As Integer
 YPosPicture = isYPos
End Property

'* English: Returns/Sets the Position Y of the image.
Public Property Let YPosPicture(ByVal theYPos As Integer)
 '* Devuelve ó establece la Posición Y de la imagen.
 isYPos = theYPos
 Call PropertyChanged("YPosPicture")
 Call Refresh(isState)
End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
 If (isEnabled = True) Then RaiseEvent Click
End Sub

Private Sub UserControl_Click()
 If (isHotTitle = False) Then
  Call Refresh(OfficeHighLight)
  RaiseEvent Click
 End If
End Sub

Private Sub UserControl_GotFocus()
 If (isHotTitle = False) Then
  isFocus = True
  Call Refresh(isState)
 End If
End Sub

Private Sub UserControl_InitProperties()
 isAutoSizePic = False
 isBackColor = ConvertSystemColor(defBackColor)
 isBorderColor = ConvertSystemColor(defBorderColor)
 isButtonShape = defShape
 isCaption = Ambient.DisplayName
 isDisabledColor = ConvertSystemColor(defDisabledColor)
 isEnabled = True
 isFontAlign = ACenter
 isForeColor = ConvertSystemColor(defForeColor)
 isHighLightColor = ConvertSystemColor(defHighLightColor)
 isHotColor = ConvertSystemColor(defHotColor)
 isHotTitle = False
 isMultiLine = False
 isPictureAlign = ACenter
 isPictureSize = 16
 isSetBorder = False
 isSetGradient = False
 isSetHighLight = True
 isShadowText = False
 isShowFocus = False
 isSystemColor = True
 isXPos = 4
 isYPos = 4
 m_bGrayIcon = False
 Set g_Font = Ambient.Font
 Set isPicture = Nothing
 ToolActive = False
 ToolBackColor = vbInfoBackground
 ToolCentered = True
 ToolForeColor = vbInfoText
 ToolIcon = 1
 TOOLSTYLE = 1
 ToolTitle = "HACKPRO TM"
 ToolText = Extender.ToolTipText
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13, 32 '* Enter.
   RaiseEvent Click
  Case 37, 38 '* Left Arrow and Up.
   Call SendKeys("+{TAB}")
  Case 39, 40 '* Right Arrow and Down.
   Call SendKeys("{TAB}")
 End Select
End Sub

Private Sub UserControl_LostFocus()
 If (isHotTitle = False) Then
  isFocus = False
  Call Refresh(isState)
 End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If (isHotTitle = False) And (Button = vbLeftButton) And (isEnabled = True) Then
  Call Refresh(OfficeHot)
 End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim tmpState As Integer
 
 If (isEnabled = True) And (isHotTitle = False) Then
  If (IsMouseOver = True) Then
   Call Refresh(isState)
  Else
   tmpState = isState
   Call Refresh(OfficeNormal)
   isState = tmpState
  End If
 End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
 With PropBag
  AutoSizePicture = .ReadProperty("AutoSizePicture", False)
  BackColor = .ReadProperty("BackColor", ConvertSystemColor(defBackColor))
  BorderColor = .ReadProperty("BorderColor", ConvertSystemColor(defBorderColor))
  ButtonShape = .ReadProperty("ButtonShape", defShape)
  Caption = .ReadProperty("Caption", Ambient.DisplayName)
  CaptionAlign = .ReadProperty("CaptionAlign", &H0)
  DisabledColor = .ReadProperty("DisabledColor", ConvertSystemColor(defDisabledColor))
  Enabled = .ReadProperty("Enabled", True)
  ForeColor = .ReadProperty("ForeColor", ConvertSystemColor(defForeColor))
  GrayIcon = PropBag.ReadProperty("GrayIcon", True)
  HighLightColor = .ReadProperty("HighlightColor", ConvertSystemColor(defHighLightColor))
  HotColor = .ReadProperty("HotColor", ConvertSystemColor(defHotColor))
  HotTitle = .ReadProperty("HotTitle", False)
  MultiLine = .ReadProperty("MultiLine", False)
  PictureAlign = .ReadProperty("PictureAlign", &H0)
  PictureSize = .ReadProperty("PictureSize", 16)
  SetBorder = .ReadProperty("SetBorder", False)
  SetBorderH = .ReadProperty("SetBorderH", True)
  SetGradient = .ReadProperty("SetGradient", False)
  Set g_Font = PropBag.ReadProperty("Font", Ambient.Font)
  SetHighLight = .ReadProperty("SetHighLight", True)
  Set isPicture = .ReadProperty("Picture", Nothing)
  Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
  ShadowText = .ReadProperty("ShadowText", False)
  ShowFocus = .ReadProperty("ShowFocus", False)
  SystemColor = .ReadProperty("SystemColor", True)
  TipActive = .ReadProperty("TipActive", False)
  TipBackColor = .ReadProperty("TipBackColor", vbInfoBackground)
  TipCentered = .ReadProperty("TipCentered", True)
  TipForeColor = .ReadProperty("TipForeColor", vbInfoText)
  TipIcon = .ReadProperty("TipIcon", 1)
  TipStyle = .ReadProperty("TipStyle", 1)
  TipTitle = .ReadProperty("TipTitle", "HACKPRO TM")
  TipText = .ReadProperty("TipText", "")
  UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
  XPosPicture = .ReadProperty("XPosPicture", 4)
  YPosPicture = .ReadProperty("YPosPicture", 4)
 End With
 If (Ambient.UserMode = True) Then
  bTrack = True
  bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
  If Not (bTrackUser32 = True) Then
   If Not (IsFunctionExported("_TrackMouseEvent", "Comctl32") = True) Then
    bTrack = False
   End If
  End If
  If (bTrack = True) Then '* OS supports mouse leave so subclass for it.
   '* Start subclassing the UserControl.
   Call Subclass_Start(hWnd)
   Call Subclass_AddMsg(hWnd, WM_MOUSEMOVE, MSG_AFTER)
   Call Subclass_AddMsg(hWnd, WM_MOUSELEAVE, MSG_AFTER)
   Call Subclass_AddMsg(hWnd, WM_THEMECHANGED, MSG_AFTER)
   Call Subclass_AddMsg(hWnd, WM_SYSCOLORCHANGE, MSG_AFTER)
  End If
 End If
End Sub

Private Sub UserControl_Resize()
 If (isHotTitle = False) Then Call Refresh(isState) '* Call the Refresh Sub.
End Sub

'* The control is terminating - a good place to stop the subclasser
Private Sub UserControl_Terminate()
On Error GoTo Catch
 Call TipRemove
 If (Ambient.UserMode = True) Then Call Subclass_StopAll '* Stop all subclassing.
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
 With PropBag
  Call .WriteProperty("AutoSizePicture", isAutoSizePic, False)
  Call .WriteProperty("BackColor", isBackColor, ConvertSystemColor(defBackColor))
  Call .WriteProperty("BorderColor", isBorderColor, ConvertSystemColor(defBorderColor))
  Call .WriteProperty("ButtonShape", isButtonShape, defShape)
  Call .WriteProperty("Caption", isCaption, Ambient.DisplayName)
  Call .WriteProperty("CaptionAlign", isFontAlign, &H0)
  Call .WriteProperty("DisabledColor", isDisabledColor, ConvertSystemColor(defDisabledColor))
  Call .WriteProperty("Enabled", isEnabled, True)
  Call .WriteProperty("Font", g_Font, Ambient.Font)
  Call .WriteProperty("ForeColor", isForeColor, ConvertSystemColor(defForeColor))
  Call .WriteProperty("GrayIcon", m_bGrayIcon, True)
  Call .WriteProperty("HighlightColor", isHighLightColor, ConvertSystemColor(defHighLightColor))
  Call .WriteProperty("HotColor", isHotColor, ConvertSystemColor(defHotColor))
  Call .WriteProperty("HotTitle", isHotTitle, False)
  Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
  Call .WriteProperty("MousePointer", MousePointer, vbDefault)
  Call .WriteProperty("MultiLine", isMultiLine, False)
  Call .WriteProperty("Picture", isPicture, Nothing)
  Call .WriteProperty("PictureAlign", isPictureAlign, &H0)
  Call .WriteProperty("PictureSize", isPictureSize, 16)
  Call .WriteProperty("SetBorder", isSetBorder, False)
  Call .WriteProperty("SetBorderH", isSetBorderH, True)
  Call .WriteProperty("SetGradient", isSetGradient, False)
  Call .WriteProperty("SetHighLight", isSetHighLight, True)
  Call .WriteProperty("ShadowText", isShadowText, False)
  Call .WriteProperty("ShowFocus", isShowFocus, False)
  Call .WriteProperty("SystemColor", isSystemColor, True)
  Call .WriteProperty("TipActive", ToolActive, False)
  Call .WriteProperty("TipBackColor", ToolBackColor, vbInfoBackground)
  Call .WriteProperty("TipCentered", ToolCentered, True)
  Call .WriteProperty("TipForeColor", ToolForeColor, vbInfoText)
  Call .WriteProperty("TipIcon", ToolIcon, 1)
  Call .WriteProperty("TipStyle", TOOLSTYLE, 1)
  Call .WriteProperty("TipText", ToolText, "")
  Call .WriteProperty("TipTitle", ToolTitle, "HACKPRO TM")
  Call .WriteProperty("XPosPicture", isXPos, 4)
  Call .WriteProperty("YPosPicture", isYPos, 4)
 End With
On Error GoTo 0
End Sub

'*******************************************************'
'* Private Subs and Functions.                         *'
'*******************************************************'

'* English: Paints lines in a simple and faster.
Private Sub APILine(ByVal whDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal lColor As Long)
 Dim PT As POINTAPI, hPen As Long, hPenOld As Long
 
 '* Pinta líneas de forma sencilla y rápida.
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(whDC, hPen)
 Call MoveToEx(whDC, X1, Y1, PT)
 Call LineTo(whDC, x2, y2)
 Call SelectObject(whDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

'* English: Convert Long to System Color.
Private Function ConvertSystemColor(ByVal theColor As Long) As Long
 '* Convierte un long en un color del sistema.
 Call OleTranslateColor(theColor, 0, ConvertSystemColor)
End Function

'* English: Paints a rectangle with oval border.
Private Sub DrawBox(ByVal hDC As Long, ByVal Offset As Long, ByVal Radius As Long, ByVal ColorFill As Long, ByVal ColorBorder As Long, ByVal isWidth As Long, ByVal isHeight As Long)
 Dim pRect As RECT, hPen As Long, hBrush As Long
 
 '* Crea un rectángulo con border ovalados.
On Error Resume Next
 pRect.xLeft = -4
 pRect.xRight = isWidth - IIf(isCaption = "", 1, -2)
 pRect.xTop = -3
 pRect.xBottom = isHeight - 1
 hPen = SelectObject(hDC, CreatePen(PS_SOLID, 1, ColorBorder))
 hBrush = SelectObject(hDC, CreateSolidBrush(ColorFill))
 Call InflateRect(pRect, -Offset, -Offset)
 Call RoundRect(hDC, pRect.xLeft, pRect.xTop, pRect.xRight, pRect.xBottom, Radius, Radius)
 Call InflateRect(pRect, Offset, Offset)
 Call DeleteObject(SelectObject(hDC, hPen))
 Call DeleteObject(SelectObject(hDC, hBrush))
On Error GoTo 0
End Sub

'* English: Draw the text on the Object.
Private Sub DrawCaption(ByVal iColor1 As Long, ByVal iColor2 As Long)
 Dim lColor As Long, isFAlign As Long
   
 '* Dibuja el texto sobre el Objeto.
 If (isMultiLine = True) Then lColor = DT_WORDBREAK Else lColor = DT_SINGLELINE
 Select Case isFontAlign
  Case ACenter
   isFAlign = DT_CENTER Or DT_VCENTER Or lColor Or DT_WORD_ELLIPSIS
  Case ALeft
   isFAlign = DT_VCENTER Or DT_LEFT Or lColor Or DT_WORD_ELLIPSIS
  Case ARight
   isFAlign = DT_VCENTER Or DT_RIGHT Or lColor Or DT_WORD_ELLIPSIS
  Case ATop
   isFAlign = DT_CENTER Or DT_TOP Or lColor Or DT_WORD_ELLIPSIS
  Case ABottom
   isFAlign = DT_CENTER Or DT_BOTTOM Or lColor Or DT_WORD_ELLIPSIS
 End Select
 If (isState <> OfficeDisabled) Then
  lColor = iColor2
 Else
  lColor = iColor1
 End If
 If (isShadowText = True) And ((isState = &H1) Or (isState = &H2)) Then
  isTxtRect.xLeft = isTxtRect.xLeft + 1.5
  isTxtRect.xTop = isTxtRect.xTop + 1.5
  Call SetTextColor(UserControl.hDC, ShiftColorOXP(lColor))
  Call DrawText(UserControl.hDC, isCaption, -1, isTxtRect, isFAlign)
  isTxtRect.xLeft = isTxtRect.xLeft - 1.5
  isTxtRect.xTop = isTxtRect.xTop - 1.5
 End If
 Call SetTextColor(UserControl.hDC, lColor)
 Call DrawText(UserControl.hDC, isCaption, -1, isTxtRect, isFAlign)
End Sub

'* English: Show focus of control.
Private Sub DrawFocus()
 Dim iPos As Integer
 
 '* Muestra el enfoque del control.
 If (isFocus = True) And (isShowFocus = True) Then
  If (isButtonShape = &H0) Then '* Shape Rectangle.
   Call DrawFocusRect(UserControl.hDC, RectButton)
  Else
   For iPos = RectButton.xLeft + 3 To RectButton.xRight - IIf(isCaption = "", 7, 4)
    Call SetPixel(UserControl.hDC, iPos, RectButton.xTop + 1, &H1DD6B7)
    Call SetPixel(UserControl.hDC, iPos, RectButton.xTop + isHeight - 3, &H1DD6B7)
   Next
   For iPos = RectButton.xTop + 4 To RectButton.xTop + isHeight - 5
    Call SetPixel(UserControl.hDC, RectButton.xLeft, iPos, &H1DD6B7)
    Call SetPixel(UserControl.hDC, RectButton.xRight - IIf(isCaption = "", 4, 1), iPos, &H1DD6B7)
   Next
   For iPos = RectButton.xLeft + 3 To RectButton.xRight - IIf(isCaption = "", 7, 4) Step 2
    Call SetPixel(UserControl.hDC, iPos, RectButton.xTop + 1, &H24427A)
    Call SetPixel(UserControl.hDC, iPos, RectButton.xTop + isHeight - 3, &H24427A)
   Next
   For iPos = RectButton.xTop + 4 To RectButton.xTop + isHeight - 5 Step 2
    Call SetPixel(UserControl.hDC, RectButton.xLeft, iPos, &H24427A)
    Call SetPixel(UserControl.hDC, RectButton.xRight - IIf(isCaption = "", 4, 1), iPos, &H24427A)
   Next
   Call SetPixel(UserControl.hDC, RectButton.xLeft + 1, 2, vbBlack)
   Call SetPixel(UserControl.hDC, RectButton.xRight - IIf(isCaption = "", 5, 2), 2, vbBlack)
   Call SetPixel(UserControl.hDC, RectButton.xLeft + 1, RectButton.xTop + isHeight - 4, vbBlack)
   Call SetPixel(UserControl.hDC, RectButton.xRight - IIf(isCaption = "", 5, 2), RectButton.xTop + isHeight - 4, vbBlack)
  End If
 End If
End Sub

'* English: Draws a degraded one in vertical form.
Private Sub DrawVGradient(ByVal whDC As Long, ByVal lEndColor As Long, ByVal lStartColor As Long, ByVal X As Long, ByVal Y As Long, ByVal x2 As Long, ByVal y2 As Long)
 Dim dR As Single, dG As Single, dB As Single, ni As Long
 Dim sR As Single, sG As Single, Sb As Single
 Dim eR As Single, eG As Single, eB As Single
 
 '* Dibuja un degradado en forma vertical.
 sR = (lStartColor And &HFF)
 sG = (lStartColor \ &H100) And &HFF
 Sb = (lStartColor And &HFF0000) / &H10000
 eR = (lEndColor And &HFF)
 eG = (lEndColor \ &H100) And &HFF
 eB = (lEndColor And &HFF0000) / &H10000
 dR = (sR - eR) / y2
 dG = (sG - eG) / y2
 dB = (Sb - eB) / y2
 For ni = 0 To y2
  Call APILine(whDC, X, Y + ni, x2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)))
 Next
End Sub

'* English: Draw a rectangle area with a specific color.
Private Sub DrawRectangle(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal ColorFill As Long, ByVal ColorBorder As Long, Optional ByVal SetBackground As Boolean = True)
 Dim hBrush As Long, TempRect As RECT

 '* Crea un área rectangular con un color específico.
 TempRect.xLeft = X
 TempRect.xTop = Y
 TempRect.xRight = X + Width
 TempRect.xBottom = Y + Height
 hBrush = CreateSolidBrush(ColorBorder)
 Call FrameRect(hDC, TempRect, hBrush)
 Call DeleteObject(hBrush)
 If (SetBackground = True) Then
  TempRect.xLeft = X + 1
  TempRect.xTop = Y + 1
  TempRect.xRight = X + Width - 1
  TempRect.xBottom = Y + Height - 1
  hBrush = CreateSolidBrush(ColorFill)
  Call FillRect(hDC, TempRect, hBrush)
  Call DeleteObject(hBrush)
 End If
End Sub

'* English: Draw a picture in the Object.
Private Sub DrawPicture()
 Dim isType As Long, isValue As Long
 
 '* Crea la imagen sobre el Objeto.
On Error Resume Next
 If Not (isPicture Is Nothing) Then
  If (Picture <> 0) Then
   Dim iX As Long, iY As Long
   
   If (isPictureSize <= 0) Then isPictureSize = 16
   Select Case isPicture.Type
    Case 1, 4: isType = DST_BITMAP
    Case 3:    isType = DST_ICON
   End Select
   If (isPictureAlign = &H0) Then
    iX = (isWidth - isPictureSize) / 2
    iY = (isHeight - isPictureSize) / 2
   ElseIf (isPictureAlign = &H1) Then
    iX = isXPos
    iY = (isHeight - isPictureSize) / 2
   ElseIf (isPictureAlign = &H2) Then
    iX = isWidth - isPictureSize - isXPos
    iY = (isHeight - isPictureSize) / 2
   ElseIf (isPictureAlign = &H3) Then
    iX = (isWidth - isPictureSize) / 2
    iY = isYPos
   ElseIf (isPictureAlign = &H4) Then
    iX = (isWidth - isPictureSize) / 2
    iY = isHeight - isPictureSize - isYPos
   End If
  End If
  If (isEnabled = False) Then
   isValue = DSS_DISABLED
   If (m_bGrayIcon = False) Then
    Call DrawState(UserControl.hDC, 0, 0, isPicture.handle, 0, iX, iY, isPictureSize, isPictureSize, isType Or isValue)
   Else
    Call RenderIconGrayscale(UserControl.hDC, isPicture.handle, iX, iY, isPictureSize, isPictureSize)
   End If
  Else
   isValue = DSS_NORMAL
   If (isState = OfficeHot) Then
    iX = iX - 1
    iY = iY - 1
   ElseIf (isState = OfficeHighLight) Then
    isValue = CreateSolidBrush(RGB(136, 141, 157))
    Call DrawState(UserControl.hDC, isValue, 0, isPicture.handle, 0, iX, iY, isPictureSize, isPictureSize, isType Or DSS_MONO)
    iX = iX - 2
    iY = iY - 2
    isValue = DSS_NORMAL
    Call DrawState(UserControl.hDC, 0, 0, isPicture.handle, 0, iX, iY, isPictureSize, isPictureSize, isType Or isValue)
    Call DeleteObject(isValue)
    Exit Sub
   End If
   Call RenderIconGrayscale(UserControl.hDC, isPicture.handle, iX, iY, isPictureSize, isPictureSize, False)
  End If
 End If
End Sub

'* English: Return, if the mouse is over the Object.
Private Function IsMouseOver() As Boolean
 Dim PT As POINTAPI
 
 '* Devuelve si el mouse esta sobre el objeto.
 Call GetCursorPos(PT)
 IsMouseOver = (WindowFromPoint(PT.X, PT.Y) = hWnd)
End Function

'* English: Executable file or a document file.
Public Function OpenLink(ByVal sLink As String) As Long
 '* Ejecuta un archivo ó documento cualquiera.
On Error Resume Next
 OpenLink = ShellExecute(Parent.hWnd, vbNullString, sLink, vbNullString, "C:\", SW_SHOWNORMAL)
On Error GoTo 0
End Function

'* English: Draw appearance of the control.
Private Sub Refresh(Optional ByVal State As OfficeState = 0)
 Dim lColor  As Long, lBase   As Long, iColor1 As Long
 Dim iColor2 As Long, iColor3 As Long, iColor4 As Long
 Dim iColor5 As Long, iColor6 As Long, lBase1  As Integer
 
 '* Crea la apariencia del control.
 If (isEnabled = False) Then State = OfficeDisabled
 If (isSystemColor = False) Then
  iColor1 = isBackColor
  iColor2 = isBorderColor
  iColor3 = isDisabledColor
  iColor4 = isForeColor
  iColor5 = isHighLightColor
  iColor6 = isHotColor
 Else
  iColor1 = ConvertSystemColor(defBackColor)
  iColor2 = ConvertSystemColor(defBorderColor)
  iColor3 = ConvertSystemColor(defDisabledColor)
  iColor4 = ConvertSystemColor(defForeColor)
  iColor5 = ConvertSystemColor(defHighLightColor)
  iColor6 = ConvertSystemColor(defHotColor)
 End If
 If (isEnabled = False) Then iColor2 = iColor3
 With UserControl
  isHeight = .ScaleHeight
  isWidth = .ScaleWidth
  .AutoRedraw = True
  .ScaleMode = vbPixels
  .Cls
 On Error Resume Next
  Set .Font = g_Font
  Call GetClientRect(.hWnd, RectButton)
  Call GetClientRect(.hWnd, isTxtRect)
  .BackColor = iColor1
  lBase = &HB0
  lBase1 = 1
  If Not (isButtonShape = &H0) Then lBase1 = 4
  'If (State > &H0) And (State < &H3) And (isSetGradient = True) Then State = &H0
  Select Case State
   Case &H0 '* Normal State.
    If (isSetGradient = True) Then
     'Call DrawVGradient(.hDC, iColor1, ShiftColorOXP(iColor1, &H72), 0, 0, .ScaleWidth - lBase1, .ScaleHeight - lBase1)
     Call DrawVGradient(.hDC, ShiftColorOXP(iColor1, &H72), iColor1, 0, 0, .ScaleWidth - lBase1, .ScaleHeight - lBase1)
    End If
    If (isSetBorder = True) Then
     If (isButtonShape = &H0) Then
      Call DrawRectangle(.hDC, 0, 0, isWidth, isHeight, iColor1, iColor2, IIf(isSetGradient = True, False, True))
     Else
      Call DrawBox(.hDC, 4, 5, iColor1, iColor2, RectButton.xRight + 2, RectButton.xBottom + 3)
     End If
    ElseIf (isSetGradient = False) Then
     Call DrawRectangle(.hDC, 0, 0, isWidth, isHeight, iColor1, iColor1)
    End If
   Case &H1, &H2 '* HighLight or Hot State.
    If (isSetHighLight = True) Then
     If (State = &H1) Then
      lColor = ShiftColorOXP(iColor5, &H40)
      If (isSetGradient = True) Then Call DrawVGradient(.hDC, iColor1, ShiftColorOXP(iColor5, &H122), 0, 0, .ScaleWidth - lBase1, .ScaleHeight - lBase1)
     Else
      lColor = ShiftColorOXP(iColor6, &H10)
      lBase = &H9C
      If (isSetGradient = True) Then Call DrawVGradient(.hDC, iColor1, ShiftColorOXP(iColor6, &H40), 0, 0, .ScaleWidth - lBase1, .ScaleHeight - lBase1)
     End If
    ElseIf (isSetBorderH = True) Then
     lColor = iColor1
     lBase = 0
    End If
    If (isSetBorderH = True) And (isButtonShape = &H0) Then
     Call DrawRectangle(.hDC, 0, 0, isWidth, isHeight, ShiftColorOXP(lColor, lBase), iColor2, IIf(isSetGradient = True, False, True))
    ElseIf (isSetBorderH = True) Then
     Call DrawBox(.hDC, 4, 5, ShiftColorOXP(lColor, lBase), iColor2, RectButton.xRight + 2, RectButton.xBottom + 3)
    End If
   Case &H3 '* Disabled State.
    lColor = iColor3
    If (isSetBorder = True) Then
     If (isButtonShape = &H0) Then
      Call DrawRectangle(.hDC, 0, 0, isWidth, isHeight, iColor1, lColor)
     Else
      Call DrawBox(.hDC, 4, 5, iColor1, lColor, RectButton.xRight + 2, RectButton.xBottom + 3)
     End If
    End If
  End Select
  isState = State
  If (isAutoSizePic = True) Then
   .Width = isPicture.Width
   .Height = isPicture.Height
   isHeight = .ScaleHeight
   isWidth = .ScaleWidth
  End If
  Call DrawCaption(iColor3, iColor4)
  Call DrawPicture
  If (isState <> &H3) Then Call DrawFocus
 End With
End Sub

'* English: Returns or sets a string that contains the keys that will act as the access keys (or hot keys for the control.)
Private Sub SetAccessKey(ByVal Caption As String)
 Dim AmperSandPos As Long, isText As String

 '* Devuelve ó establece una cadena que contiene las teclas que funcionarán como teclas de acceso (o teclas aceleradoras) del control.
 With UserControl
  .AccessKeys = ""
  If (Len(Caption) > 1) Then
   AmperSandPos = InStr(1, Caption, "&", vbTextCompare)
   If (AmperSandPos < Len(Caption)) And (AmperSandPos > 0) Then
    isText = Mid$(Caption, AmperSandPos + 1, 1)
    If (isText <> "&") Then
     .AccessKeys = LCase$(isText)
    Else
     AmperSandPos = InStr(AmperSandPos + 2, Caption, "&", vbTextCompare)
     isText = Mid$(Caption, AmperSandPos + 1, 1)
     If (isText <> "&") Then .AccessKeys = LCase$(isText)
    End If
   End If
  End If
 End With
End Sub

'* English: Shift a color.
Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
 Dim Red   As Long, Blue  As Long
 Dim Delta As Long, Green As Long
   
 '* Devuelve un Color con menos intensidad.
 Blue = ((theColor \ &H10000) Mod &H100)
 Green = ((theColor \ &H100) Mod &H100)
 Red = (theColor And &HFF)
 Delta = &HFF - Base
 Blue = Base + Blue * Delta \ &HFF
 Green = Base + Green * Delta \ &HFF
 Red = Base + Red * Delta \ &HFF
 If (Red > 255) Then Red = 255
 If (Green > 255) Then Green = 255
 If (Blue > 255) Then Blue = 255
 ShiftColorOXP = Red + 256& * Green + 65536 * Blue
End Function

'* ======================================================================================================
'*  UserControl private routines.
'*  Determine if the passed function is supported.
'* ======================================================================================================
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
 Dim hMod As Long, bLibLoaded As Boolean
 
 hMod = GetModuleHandleA(sModule)
 If (hMod = 0) Then
  hMod = LoadLibraryA(sModule)
  If (hMod) Then bLibLoaded = True
 End If
 If (hMod) Then
  If (GetProcAddress(hMod, sFunction)) Then IsFunctionExported = True
 End If
 If (bLibLoaded = True) Then Call FreeLibrary(hMod)
End Function

'* Track the mouse leaving the indicated window.
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
 Dim tme As TRACKMOUSEEVENT_STRUCT
 
 If (bTrack = True) Then
  With tme
   .cbSize = Len(tme)
   .dwFlags = TME_LEAVE
   .hWndTrack = lng_hWnd
  End With
  If (bTrackUser32 = True) Then
   Call TrackMouseEvent(tme)
  Else
   Call TrackMouseEventComCtl(tme)
  End If
 End If
End Sub

'* =============================================================================================================================
'*  Subclass code - The programmer may call any of the following Subclass_??? routines
'*  Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages.
'* =============================================================================================================================
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
 '* Parameters:
 '*  lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
 '*  uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
 '*  When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
 With sc_aSubData(zIdx(lng_hWnd))
  If (When) And (eMsgWhen.MSG_BEFORE) Then
   Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
  End If
  If (When) And (eMsgWhen.MSG_AFTER) Then
   Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
  End If
 End With
End Sub

'* Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
 '* Parameters:
 '*  lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table.
 '*  uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback.
 '*  When      - Whether the msg is to be removed from the before, after or both callback tables.
 With sc_aSubData(zIdx(lng_hWnd))
  If (When) And (eMsgWhen.MSG_BEFORE) Then
   Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
  End If
  If (When) And (eMsgWhen.MSG_AFTER) Then
   Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
  End If
 End With
End Sub

'* Return whether were running in the IDE.
Private Function Subclass_InIDE() As Boolean
 Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'* Start subclassing the passed window handle.
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
 '* Parameters:
 '*  lng_hWnd - The handle of the window to be subclassed.
 '*  Returns;
 '*  The sc_aSubData() index.
 Const CODE_LEN              As Long = 200
 Const FUNC_CWP              As String = "CallWindowProcA"
 Const FUNC_EBM              As String = "EbMode"
 Const FUNC_SWL              As String = "SetWindowLongA"
 Const MOD_USER              As String = "user32"
 Const MOD_VBA5              As String = "vba5"
 Const MOD_VBA6              As String = "vba6"
 Const PATCH_01              As Long = 18
 Const PATCH_02              As Long = 68
 Const PATCH_03              As Long = 78
 Const PATCH_06              As Long = 116
 Const PATCH_07              As Long = 121
 Const PATCH_0A              As Long = 186
 Static aBuf(1 To CODE_LEN)  As Byte
 Static pCWP                 As Long
 Static pEbMode              As Long
 Static pSWL                 As Long
 Dim i                       As Long
 Dim j                       As Long
 Dim nSubIdx                 As Long
 Dim sHex                    As String
 
 '* If it's the first time through here...
 If (aBuf(1) = 0) Then
  '* The hex pair machine code representation.
  sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
  '* Convert the string from hex pairs to bytes and store in the static machine code buffer.
  i = 1
  Do While (j < CODE_LEN)
   j = j + 1
   aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
   i = i + 2
  Loop
  '* Get API function addresses.
  If (Subclass_InIDE = True) Then
   aBuf(16) = &H90
   aBuf(17) = &H90
   pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
   If (pEbMode = 0) Then pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
  End If
  pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
  pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
  ReDim sc_aSubData(0 To 0) As tSubData
 Else
  nSubIdx = zIdx(lng_hWnd, True)
  If (nSubIdx = -1) Then
   nSubIdx = UBound(sc_aSubData()) + 1
   ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
  End If
  Subclass_Start = nSubIdx
 End If
 With sc_aSubData(nSubIdx)
  .hWnd = lng_hWnd
  .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
  .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)
  Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
  Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
  Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
  Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
  Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
  Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
  Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))
 End With
End Function

'* Stop all subclassing.
Private Sub Subclass_StopAll()
 Dim i As Long
 
On Error GoTo myErr
 i = UBound(sc_aSubData()) '* Get the upper bound of the subclass data array.
 Do While (i >= 0)         '* Iterate through each element.
  With sc_aSubData(i)      '* If not previously Subclass_Stop'd.
   If (.hWnd <> 0) Then Call Subclass_Stop(.hWnd)
  End With
  i = i - 1
 Loop
 Exit Sub
myErr:
End Sub

'* Stop subclassing the passed window handle.
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
 '* Parameters:
 '*  lng_hWnd - The handle of the window to stop being subclassed.
 '*  Parameters:
 '*  lng_hWnd  - The handle of the window to stop being subclassed
 With sc_aSubData(zIdx(lng_hWnd))
  Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig) '* Restore the original WndProc.
  Call zPatchVal(.nAddrSub, PATCH_05, 0)              '* Patch the Table B entry count to ensure no further 'before' callbacks.
  Call zPatchVal(.nAddrSub, PATCH_09, 0)              '* Patch the Table A entry count to ensure no further 'after' callbacks.
  Call GlobalFree(.nAddrSub)                          '* Release the machine code memory.
  .hWnd = 0                                           '* Mark the sc_aSubData element as available for re-use.
  .nMsgCntB = 0                                       '* Clear the before table.
  .nMsgCntA = 0                                       '* Clear the after table.
  Erase .aMsgTblB                                     '* Erase the before table.
  Erase .aMsgTblA                                     '* Erase the after table.
 End With
End Sub

'* ======================================================================================================
'*  These z??? routines are exclusively called by the Subclass_??? routines.
'*  Worker sub for Subclass_AddMsg.
'* ======================================================================================================
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
 Dim nEntry As Long, nOff1 As Long, nOff2 As Long
 
 If (uMsg = ALL_MESSAGES) Then
  nMsgCnt = ALL_MESSAGES
 Else
  Do While (nEntry < nMsgCnt)
   nEntry = nEntry + 1
   If (aMsgTbl(nEntry) = 0) Then
    aMsgTbl(nEntry) = uMsg
    Exit Sub
   ElseIf (aMsgTbl(nEntry) = uMsg) Then
    Exit Sub
   End If
  Loop
  nMsgCnt = nMsgCnt + 1
  ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
  aMsgTbl(nMsgCnt) = uMsg
 End If
 If (When = eMsgWhen.MSG_BEFORE) Then
  nOff1 = PATCH_04
  nOff2 = PATCH_05
 Else
  nOff1 = PATCH_08
  nOff2 = PATCH_09
 End If
 If (uMsg <> ALL_MESSAGES) Then Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
 Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

'* Return the memory address of the passed function in the passed dll.
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
 zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
 Debug.Assert zAddrFunc
End Function

'* Worker sub for Subclass_DelMsg.
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
 Dim nEntry As Long
 
 If (uMsg = ALL_MESSAGES) Then
  nMsgCnt = 0
  If (When = eMsgWhen.MSG_BEFORE) Then
   nEntry = PATCH_05
  Else
   nEntry = PATCH_09
  End If
  Call zPatchVal(nAddr, nEntry, 0)
 Else
  Do While (nEntry < nMsgCnt)
   nEntry = nEntry + 1
   If (aMsgTbl(nEntry) = uMsg) Then
    aMsgTbl(nEntry) = 0
    Exit Do
   End If
  Loop
 End If
End Sub

'* Get the sc_aSubData() array index of the passed hWnd.
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
 '* Get the upper bound of sc_aSubData() - If you get an error here, youre probably Subclass_AddMsg-ing before Subclass_Start.
 zIdx = UBound(sc_aSubData)
 Do While (zIdx >= 0)
  With sc_aSubData(zIdx)
   If (.hWnd = lng_hWnd) And Not (bAdd = True) Then
    Exit Function
   ElseIf (.hWnd = 0) And (bAdd = True) Then
    Exit Function
   End If
  End With
  zIdx = zIdx - 1
 Loop
 If Not (bAdd = True) Then Debug.Assert False
 '* If we exit here, were returning -1, no freed elements were found.
End Function

'* Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
 Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'* Patch the machine code buffer at the indicated offset with the passed value.
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
 Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'* Worker function for Subclass_InIDE.
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
 zSetTrue = True
 bValue = True
End Function
'*******************************************************'

Public Property Get TipActive() As Boolean
 '* Retrieving value of a property, Boolean responce (true/false).
 '* Syntax: BooleanVar = object.TipActive.
 TipActive = ToolActive
End Property

Public Property Let TipActive(ByVal ToolData As Boolean)
 '* If True, activate (show) ToolTip, False deactivate (hide) tool tip.
 '* Syntax: object.TipActive = True/False.
 ToolActive = ToolData
 Call PropertyChanged("TipActive")
End Property

Public Property Get TipBackColor() As OLE_COLOR
 '* Retrieving value of a property, returns RGB as Long.
 '* Syntax: LongVar = object.BackColor.
 TipBackColor = ToolBackColor
End Property

Public Property Let TipBackColor(ByVal ToolData As OLE_COLOR)
 '* Assigning a value to the property, set RGB value as Long.
 '* Syntax: object.BackColor = RGB (as Long). Since 0 is _
    Black (no RGB), and the API thinks 0 is the default _
    color ("off" yellow), we need to "fudge" Black a bit _
    (yes set bit "1" to "1",). I couldn't resist the _
    pun!. So, in module or form code, if setting to Black, _
    make it "1", if restoring the default color, make it _
    "0".
 ToolBackColor = ConvertSystemColor(ToolData)
 Call PropertyChanged("TipBackColor")
End Property

Public Property Get TipCentered() As Boolean
 '* Retrieving value of a property, returns Boolean true/false.
 '* Syntax: BooleanVar = object.TipCentered.
 TipCentered = ToolCentered
End Property

Public Property Let TipCentered(ByVal ToolData As Boolean)
 '* Assigning a value to the property, Set Boolean true/false if ToolTip. _
    Is TipCentered on the parent control.
 '* Syntax: object.TipCentered = True/False.
 ToolCentered = ToolData
 Call PropertyChanged("TipCentered")
End Property

Public Property Get TipForeColor() As OLE_COLOR
 '* Retrieving value of a property, returns RGB value as Long.
 '* Syntax: LongVar = object.ForeColor.
 TipForeColor = ToolForeColor
End Property

Public Property Let TipForeColor(ByVal ToolData As OLE_COLOR)
 '* Assigning a value to the property, set RGB value as Long.
 '* Syntax: object.ForeColor = RGB(As Long).
 '* Since 0 is Black (no RGB), and the API thinks 0 is _
    the default color ("off" yellow), we need to "fudge" _
    Black a bit (yes set bit "1" to "1",). I couldn't _
    resist the pun!. So, in module or form code, if _
    setting to Black, make it "1" if restoring _
    the default color, make it "0".
 '* Syntax: object.ForeColor = RGB(as long).
 ToolForeColor = ConvertSystemColor(ToolData)
 Call PropertyChanged("TipForeColor")
End Property

Public Property Get TipIcon() As ToolIconType
 '* Retrieving value of a property, returns string.
 '* Syntax: StringVar = object.TipIcon.
 TipIcon = ToolIcon
End Property

Public Property Let TipIcon(ByVal ToolData As ToolIconType)
 '* Assigning a value to the property, set TipIcon TipStyle with type var.
 '* Syntax: object.TipIcon = IconStyle.
 '* TipIcon Styles are: INFO, WARNING And ERROR (TipNoIcom, TipIconInfo, TipIconWarning, TipIconError).
 ToolIcon = ToolData
 Call PropertyChanged("TipIcon")
End Property

Public Property Get TipStyle() As ToolStyleEnum
 '* Retrieving value of a property, returns string.
 '* Syntax: StringVar = object.TipStyle.
 TipStyle = TOOLSTYLE
End Property

Public Property Let TipStyle(ByVal ToolData As ToolStyleEnum)
 '* Assigning a value to the property, set TipStyle param Standard or Balloon
 '* Syntax: object.TipStyle = TipStyle.
 TOOLSTYLE = ToolData
 Call PropertyChanged("TipStyle")
End Property

Public Property Get TipText() As String
 '* Retrieving value of a property, returns string..
 '* Syntax: StringVar = object.TipText.
 TipText = ToolText
End Property

Public Property Let TipText(ByVal ToolData As String)
 '* Assigning a value to the property, Set as String.
 '* Syntax: object.TipText = StringVar.
 '* Multi line Tips are enabled in the Create sub.
 '* To change lines, just add a vbCrLF between text.
 '* ex. object.TipText = "Line 1 text" & vbCrLF & _
    "Line 2 text".
 ToolText = ToolData
 Call PropertyChanged("TipText")
End Property

Public Property Get TipTitle() As String
 '* Retrieving value of a property, returns string.
 '* Syntax: StringVar = object.TipTitle.
 TipTitle = ToolTitle
End Property

Public Property Let TipTitle(ByVal ToolData As String)
 '* Assigning a value to the property, set as string.
 '* Syntax: object.TipTitle = StringVar.
 ToolTitle = ToolData
 Call PropertyChanged("TipTitle")
End Property

'* Private sub used with Create and Update subs/functions.
Private Sub CreateToolTip()
 Dim lpRect As RECT, lWinStyle As Long
 
 '* If Tool Tip already made, destroy it and reconstruct.
 Call TipRemove
 lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
 '* Create Baloon TipStyle if desired.
 If (TOOLSTYLE = StyleBalloon) Then lWinStyle = lWinStyle Or TTS_BALLOON
 '* The parent control has to be set first.
 If (UserControl.hWnd <> &H0) Then
  m_ltthWnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, UserControl.hWnd, 0&, App.hInstance, 0&)
  Call SendMessage(m_ltthWnd, TTM_ACTIVATE, CInt(ToolActive), TI)
  '* Make our ToolTip window a topmost window.
  Call SetWindowPos(m_ltthWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
  '* Get the rectangle of the parent control.
  Call GetClientRect(UserControl.hWnd, lpRect)
  '* Now set up our ToolTip info structure.
  With TI
   '* If we want it TipCentered, then set that flag.
   If (ToolCentered = True) Then
    .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
   Else
    .lFlags = TTF_SUBCLASS
   End If
   '* Set the hWnd prop to our Parent Control's hWnd.
   .lhWnd = UserControl.hWnd
   .lId = 0
   .hInstance = App.hInstance
   .lpRect = lpRect
   .lpStr = ToolText
  End With
  '* Add the ToolTip Structure.
  Call SendMessage(m_ltthWnd, TTM_ADDTOOLA, 0&, TI)
  '* Set Max Width to 32 characters, and enable Multi Line Tool Tips.
  Call SendMessage(m_ltthWnd, TTM_SETMAXTIPWIDTH, 0&, &H20)
  If (ToolIcon <> TipNoIcon) Or (ToolTitle <> vbNullString) Then
   '* If we want a TipTitle or we want an TipIcon.
   Call SendMessage(m_ltthWnd, TTM_SETTITLE, CLng(ToolIcon), ByVal ToolTitle)
  End If
  If (ToolForeColor <> Empty) Then
   '* 0 (zero) or Null is seen by the API as the default color. _
      See ForeColor property for more datails.
   Call SendMessage(m_ltthWnd, TTM_SETTIPTEXTCOLOR, ToolForeColor, 0&)
  End If
  If (ToolBackColor <> Empty) Then
   '* 0 (zero) or Null is seen by the API as the default color. _
      See BackColor property for more datails.
   Call SendMessage(m_ltthWnd, TTM_SETTIPBKCOLOR, ToolBackColor, 0&)
  End If
 End If
End Sub

Public Sub TipRemove()
 '* Kills Tool Tip Object.
 If (m_ltthWnd <> 0) Then Call DestroyWindow(m_ltthWnd)
End Sub

Private Sub UpDate()
 '* Used to update tooltip parameters that require reconfiguration of _
    subclass to envoke.
 If (ToolActive = True) Then Call CreateToolTip '* Refresh the object.
End Sub

' See post: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=58622&lngWId=1
' Thanks MArio Florez.
Private Function RenderIconGrayscale(ByVal Dest_hDC As Long, ByVal hIcon As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Dest_Height As Long, Optional ByVal Dest_Width As Long, Optional ByVal GrayC As Boolean = True) As Boolean
 Dim hBMP_Mask As Long, hBMP_Image As Long
 Dim hBMP_Prev As Long, hIcon_Temp As Long
 Dim hDC_Temp  As Long

 ' Make sure parameters passed are valid
 If (Dest_hDC = 0) Or (hIcon = 0) Then Exit Function
 ' Extract the bitmaps from the icon
 If (GetIconBitmaps(hIcon, hBMP_Mask, hBMP_Image) = False) Then Exit Function
 ' Create a memory DC to work with
 hDC_Temp = CreateCompatibleDC(0)
 If (hDC_Temp = 0) Then GoTo CleanUp
 ' Make the image bitmap gradient
 If (RenderBitmapGrayscale(hDC_Temp, hBMP_Image, 0, 0, , , GrayC) = False) Then GoTo CleanUp
 ' Extract the gradient bitmap out of the DC
 Call SelectObject(hDC_Temp, hBMP_Prev)
 ' Take the newly gradient bitmap and make a gradient icon from it
 hIcon_Temp = CreateIconFromBMP(hBMP_Mask, hBMP_Image)
 If (hIcon_Temp = 0) Then GoTo CleanUp
 ' Draw the newly created gradient icon onto the specified DC
 If (DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hIcon_Temp, Dest_Width, Dest_Height, 0, 0, &H3) <> 0) Then
  RenderIconGrayscale = True
 End If
CleanUp:
 Call DestroyIcon(hIcon_Temp): hIcon_Temp = 0
 Call DeleteDC(hDC_Temp): hDC_Temp = 0
 Call DeleteObject(hBMP_Mask): hBMP_Mask = 0
 Call DeleteObject(hBMP_Image): hBMP_Image = 0
End Function

Public Function GetIconBitmaps(ByVal hIcon As Long, ByRef Return_hBmpMask As Long, ByRef Return_hBmpImage As Long) As Boolean
 Dim TempICONINFO As ICONINFO

 If (GetIconInfo(hIcon, TempICONINFO) = 0) Then Exit Function
 Return_hBmpMask = TempICONINFO.hbmMask
 Return_hBmpImage = TempICONINFO.hbmColor
 GetIconBitmaps = True
End Function

'=============================================================================================================
Private Function RenderBitmapGrayscale(ByVal Dest_hDC As Long, ByVal hBitmap As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Srce_X As Long, Optional ByVal Srce_Y As Long, Optional ByVal GrayC As Boolean = True) As Boolean
 Dim TempBITMAP As BITMAP, hScreen   As Long
 Dim hDC_Temp   As Long, hBMP_Prev   As Long
 Dim MyCounterX As Long, MyCounterY  As Long
 Dim NewColor   As Long, hNewPicture As Long
 Dim DeletePic  As Boolean

 ' Make sure parameters passed are valid
 If (Dest_hDC = 0) Or (hBitmap = 0) Then Exit Function
 ' Get the handle to the screen DC
 hScreen = GetDC(0)
 If (hScreen = 0) Then Exit Function
 ' Create a memory DC to work with the picture
 hDC_Temp = CreateCompatibleDC(hScreen)
 If (hDC_Temp = 0) Then GoTo CleanUp
 ' If the user specifies NOT to alter the original, then make a copy of it to use
 DeletePic = False
 hNewPicture = hBitmap
 ' Select the bitmap into the DC
 hBMP_Prev = SelectObject(hDC_Temp, hNewPicture)
 ' Get the height / width of the bitmap in pixels
 If (GetObjectAPI(hNewPicture, Len(TempBITMAP), TempBITMAP) = 0) Then GoTo CleanUp
 If (TempBITMAP.bmHeight <= 0) Or (TempBITMAP.bmWidth <= 0) Then GoTo CleanUp
 ' Loop through each pixel and conver it to it's grayscale equivelant
 If (GrayC = True) Then
  For MyCounterX = 0 To TempBITMAP.bmWidth - 1
   For MyCounterY = 0 To TempBITMAP.bmHeight - 1
    NewColor = GetPixel(hDC_Temp, MyCounterX, MyCounterY)
    If (NewColor <> -1) Then
     Select Case NewColor
      ' If the color is already a grey shade, no need to convert it
      Case vbBlack, vbWhite, &H101010, &H202020, &H303030, &H404040, &H505050, &H606060, &H707070, &H808080, &HA0A0A0, &HB0B0B0, &HC0C0C0, &HD0D0D0, &HE0E0E0, &HF0F0F0
       NewColor = NewColor
      Case Else
       NewColor = 0.33 * (NewColor Mod 256) + 0.59 * ((NewColor \ 256) Mod 256) + 0.11 * ((NewColor \ 65536) Mod 256)
       NewColor = RGB(NewColor, NewColor, NewColor)
     End Select
     Call SetPixel(hDC_Temp, MyCounterX, MyCounterY, NewColor)
    End If
   Next
  Next
 End If
 ' Display the picture on the specified hDC
 Call BitBlt(Dest_hDC, Dest_X, Dest_Y, TempBITMAP.bmWidth, TempBITMAP.bmHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcCopy)
 RenderBitmapGrayscale = True
CleanUp:
 Call ReleaseDC(0, hScreen): hScreen = 0
 Call SelectObject(hDC_Temp, hBMP_Prev)
 Call DeleteDC(hDC_Temp): hDC_Temp = 0
 If (DeletePic = True) Then
  Call DeleteObject(hNewPicture)
  hNewPicture = 0
 End If
End Function

Private Function CreateIconFromBMP(ByVal hBMP_Mask As Long, ByVal hBMP_Image As Long) As Long
 Dim TempICONINFO As ICONINFO

 If (hBMP_Mask = 0) Or (hBMP_Image = 0) Then Exit Function
 TempICONINFO.fIcon = 1
 TempICONINFO.hbmMask = hBMP_Mask
 TempICONINFO.hbmColor = hBMP_Image
 CreateIconFromBMP = CreateIconIndirect(TempICONINFO)
End Function
