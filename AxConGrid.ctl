VERSION 5.00
Begin VB.UserControl AxConGrid 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   PropertyPages   =   "AxConGrid.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "AxConGrid.ctx":0011
   Begin AxConectionGrid.ucScrollbar ucScrollV 
      Height          =   3375
      Left            =   4455
      Top             =   135
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   5953
      Style           =   4
      ShowButtons     =   0   'False
      BeginProperty ThumbTooltipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AxConectionGrid.ucScrollbar ucScrollH 
      Height          =   165
      Left            =   75
      Top             =   3225
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   291
      Orientation     =   1
      Style           =   4
      ShowButtons     =   0   'False
      BeginProperty ThumbTooltipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "AxConGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxCompareGrid
'Version  : 0.01.00b
'Editor   : David Rojas [AxioUK]
'Date     : 27/09/2022
'------------------------------------
Option Explicit

Const sVersion = "0.01.0b"

Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function SetRECL Lib "user32" Alias "SetRect" (lpRect As RECTL, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal Token As Long)
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal Brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal RGBA As Long, ByRef Brush As Long) As Long
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTS, ByVal mFormat As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As eStringTrimming) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As eStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As eStringAlignment) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dx As Single, ByVal dy As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long) As Long
Private Declare Function GdipSetClipRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipCreateLineBrush Lib "gdiplus" (point1 As POINTS, point2 As POINTS, ByVal color1 As Long, ByVal Color2 As Long, ByVal WrapMd As Long, lineGradient As Long) As Long
Private Declare Function GdipSetPenBrushFill Lib "gdiplus" (ByVal pen As Long, ByVal Brush As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As Long, ByVal dStyle As DashStyle) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
Private Declare Function GdipDrawBezier Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal X3 As Single, ByVal Y3 As Single, ByVal X4 As Single, ByVal Y4 As Single) As Long
Private Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal pen As Long, ByVal StartCap As LineCap) As Long
Private Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal pen As Long, ByVal EndCap As LineCap) As Long

Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, ByRef Image As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixelOffsetMode As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
'---

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type RECTS
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type POINTS
   X As Single
   Y As Single
End Type

Private Type POINTL
    X As Long
    Y As Long
End Type

Private Enum CallOutPosition
  coLeft
  coTop
  coRight
  coBottom
End Enum

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type PicBmp
  Size As Long
  type As Long
  hBmp As Long
  hpal As Long
  Reserved As Long
End Type

Public Enum eStringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum
  
Public Enum eStringTrimming
    StringTrimmingNone = &H0
    StringTrimmingCharacter = &H1
    StringTrimmingWord = &H2
    StringTrimmingEllipsisCharacter = &H3
    StringTrimmingEllipsisWord = &H4
    StringTrimmingEllipsisPath = &H5
End Enum

Public Enum eStringFormatFlags
    StringFormatFlagsNone = &H0
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000
    StringFormatFlagsNoClip = &H4000
End Enum

Public Enum eTextAlignH
    eLeft
    eCenter
    eRight
End Enum

Public Enum eTextAlignV
    eTop
    eMiddle
    eBottom
End Enum

Public Enum eVisibleType
  scNone
  scAllways
  scOnlyActive
End Enum

Public Enum eTypeBadge
  vbNone
  vbLabel
  vbIcon
End Enum

Public Enum eImageType
  eIconFont
  eImageFile
End Enum

Private Type tCell
    Text As String
    SubText As String
    Label As String
    Tag As Variant
    BackColor As Long
    ForeColor As Long
    Font As StdFont
    Height As Single
    PointLX As Single
    PointRX As Single
    PointY As Single
    Radius As Single
    LineTo As Long
    ColumnTo As Long
    CurveTo As Long
    SideTo As tSide
    Direction As tDirection
    IconChar As Long
    IconFile As String
    CellVisible As Boolean
    LineWidth As Single
    LineStyle As DashStyle
    LineStartCap As LineCap
    LineEndCap As LineCap
    LineStartColor As OLE_COLOR
    LineEndColor As OLE_COLOR
    LineOpacity As Long
    LineVisible As Boolean
End Type

Private Type tColumn
    Header As String
    Cell() As tCell
    StartX As Long
    StartY As Long
    Width As Long
    BackColor As OLE_COLOR
    ForeColor As OLE_COLOR
    Font As StdFont
    Visible As Boolean
End Type

Public Enum tDirection
  UpSide = 0
  DownSide = 1
End Enum

Public Enum tSide
  toLeft = 0
  toRight = 1
  Auto = 2
End Enum

Public Enum LineCap
   LineCapFlat = 0
   LineCapSquare = 1
   LineCapRound = 2
   LineCapTriangle = 3
   LineCapNoAnchor = &H10         ' corresponds to flat cap
   LineCapSquareAnchor = &H11     ' corresponds to square cap
   LineCapRoundAnchor = &H12      ' corresponds to round cap
   LineCapDiamondAnchor = &H13    ' corresponds to triangle cap
   LineCapArrowAnchor = &H14      ' no correspondence
End Enum

Public Enum DashStyle
   DashStyleSolid          ' 0
   DashStyleDash           ' 1
   DashStyleDot            ' 2
   DashStyleDashDot        ' 3
   DashStyleDashDotDot     ' 4
End Enum

'Constants
Private Const CombineModeExclude As Long = &H4
Private Const WrapModeTileFlipXY = &H3
Private Const SmoothingModeHighQuality As Long = &H2
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const TLS_MINIMUM_AVAILABLE As Long = 64
Private Const IDC_HAND As Long = 32649
Private Const UnitPixel As Long = &H2&

'Define EVENTS-------------------
Public Event Click(lColumn As Long, lItem As Long, iColumnTo As Long, iLineTo As Long, iCurveTo As Long)
Public Event DblClick(lColumn As Long, lItem As Long)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Property Variables:
Private hFontCollection As Long
Private GdipToken   As Long
Private nScale      As Single
Private hGraphics   As Long
Private hCur        As Long
Private m_ImageType As eImageType

Private m_Clickable     As Boolean
Private m_Editable      As Boolean
Private m_Enabled       As Boolean

Private m_Col()         As tColumn
Private mCol            As Long
Private m_ColWidth      As Long
Private m_ActiveCol     As Long
Private m_ItemHeight      As Long
Private m_ActiveItem()  As Long
Private m_SelectItem()  As Long

Private m_BorderColor   As OLE_COLOR
Private m_BackColor     As OLE_COLOR
Private m_BoxColor      As OLE_COLOR
Private m_BorderWidth   As Long
Private m_CornerCurve   As Long
Private m_ForeColor1    As OLE_COLOR
Private m_ForeColor2    As OLE_COLOR
Private m_Font1         As StdFont
Private m_Font2         As StdFont
Private m_IconFont      As StdFont
Private m_LabelFont     As StdFont
Private m_HeaderFont    As StdFont
Private m_HeaderBackColor As OLE_COLOR
Private m_HeaderForeColor As OLE_COLOR
Private m_ColorActive   As OLE_COLOR
Private m_BorderColorActive As OLE_COLOR

Private m_CaptionAlignV As eTextAlignV
Private m_CaptionAlignH As eTextAlignH
Private m_SubTextAlignV As eTextAlignV
Private m_SubTextAlignH As eTextAlignH
Private m_FontOpacity   As Long

Private m_IconForeColor As OLE_COLOR
Private m_IconAlignV    As eTextAlignV
Private m_IconAlignH    As eTextAlignH
Private m_Bmp           As Long

Private m_SubTextVisible As eVisibleType
Private m_BadgeVisible   As eVisibleType
Private m_BadgeType      As eTypeBadge
Private m_BoxOpacity     As Long

Private m_LineWidth     As Single
Private m_LineStyle     As DashStyle
Private m_LineOpacity   As Long
Private m_LineStartCap  As LineCap
Private m_LineEndCap    As LineCap
Private m_LineStartColor As OLE_COLOR
Private m_LineEndColor  As OLE_COLOR

Private SideStart   As Long
Private SideEnd     As Long
Private InitCurve   As Long
Private EndCurve    As Long
Private mY          As Single
Private mX          As Single
Private m_X1        As Single
Private m_Y1        As Single
Private m_X2        As Single
Private m_Y2        As Single
Private m_X3        As Single
Private m_Y3        As Single
Private m_X4        As Single
Private m_Y4        As Single

Private m_DefaultSide  As tSide
Private m_ColsMoveable As Boolean
Private mRedraw        As Boolean


Public Function AddColumn(Optional sCaption As String, Optional lWidth As Long, _
                          Optional lX As Long = 0, Optional lY As Long = 0, _
                          Optional oBackColor As OLE_COLOR, Optional oForeColor As OLE_COLOR, _
                          Optional bVisible As Boolean = True) As Boolean
On Error Resume Next
Dim C As Long

C = ColCount

ReDim Preserve m_Col(C)
ReDim Preserve m_ActiveItem(C)
ReDim Preserve m_SelectItem(C)

With m_Col(C)
  .Header = sCaption
  .BackColor = IIf(oBackColor = 0, m_HeaderBackColor, oBackColor)
  .ForeColor = IIf(oForeColor = 0, m_HeaderForeColor, oForeColor)
  .Width = IIf(lWidth > 0, lWidth, m_ColWidth)
  .StartX = lX
  .StartY = lY
  .Visible = bVisible
End With

m_ActiveItem(C) = -1
m_SelectItem(C) = -1

Refresh

End Function

Public Function AddItem(ByVal lColumn As Long, ByVal eText As String, Optional eSubText As String = "", Optional eLabel As String = "", _
                        Optional eBackColor As OLE_COLOR = vbWhite, Optional eForeColor As OLE_COLOR = vbBlack, _
                        Optional lHeight As Long, Optional eTag As Variant = "", _
                        Optional eIconchar As String = "&H0", Optional eIconFile As String = "", _
                        Optional eVisible As Boolean = True) As Boolean
                        
On Error Resume Next
Dim i   As Long
  
  i = ItemCount(lColumn)
    
With m_Col(lColumn)
  ReDim Preserve .Cell(i)
  .Cell(i).Text = eText
  .Cell(i).SubText = eSubText
  .Cell(i).Label = eLabel
  .Cell(i).CellVisible = eVisible
  .Cell(i).BackColor = eBackColor
  .Cell(i).ForeColor = eForeColor
  .Cell(i).Height = IIf(lHeight > 0, lHeight, m_ItemHeight)
  .Cell(i).Tag = eTag
  .Cell(i).LineTo = -1
  .Cell(i).CurveTo = -1
  .Cell(i).LineVisible = False
  .Cell(i).IconChar = IconCharCode(eIconchar)
  .Cell(i).IconFile = eIconFile
End With

m_ActiveItem(lColumn) = -1
m_SelectItem(lColumn) = -1

If eVisible Then Refresh

End Function

Public Function AddLine(ByVal lColStart As Long, ByVal ItemStart As Long, ByVal lColEnd As Long, ByVal ItemEnd As Long, Optional lStyle As DashStyle = 0, _
                        Optional lStartCap As LineCap = 0, Optional lEndCap As LineCap = 0, _
                        Optional cStartColor As OLE_COLOR = &HFF&, Optional cEndColor As OLE_COLOR = &HC67300, _
                        Optional lLineWidth As Single = 5, Optional lOpacity As Long = 50, _
                        Optional bVisible As Boolean = True)
On Error Resume Next
Dim s As Long, C As Long, X As Long
Dim Side As tSide

If lColStart > lColEnd Then
    Side = toLeft
Else
    Side = toRight
End If

If m_Col(lColEnd).Cell(ItemEnd).LineTo = ItemStart And m_Col(lColEnd).Cell(ItemEnd).ColumnTo = lColStart Then
  With m_Col(lColEnd).Cell(ItemEnd)
    .LineStartCap = .LineEndCap
    .LineStartColor = .LineEndColor
  End With
  
Else
  With m_Col(lColStart).Cell(ItemStart)
    .ColumnTo = lColEnd
    .LineTo = ItemEnd
    .SideTo = Side
    .LineStyle = lStyle
    .LineWidth = lLineWidth
    .LineStartCap = lStartCap
    .LineEndCap = lEndCap
    .LineStartColor = cStartColor
    .LineEndColor = cEndColor
    .LineOpacity = lOpacity
    .LineVisible = bVisible
  End With

End If

If bVisible Then Refresh
End Function

Public Function AddCurve(ByVal lColumn As Long, ByVal ItemStart As Long, ByVal ItemEnd As Long, Optional lSide As tSide, _
                            Optional iRadius As Single = 50, Optional lStyle As DashStyle = 0, _
                            Optional lStartCap As LineCap = 0, Optional lEndCap As LineCap = 0, _
                            Optional cStartColor As OLE_COLOR = &HFF&, Optional cEndColor As OLE_COLOR = &HC67300, _
                            Optional lLineWidth As Single = 5, Optional lOpacity As Long = 50, _
                            Optional bVisible As Boolean = True)
On Error Resume Next

If m_Col(lColumn).Cell(ItemEnd).CurveTo = ItemStart Then
  With m_Col(lColumn).Cell(ItemEnd)
    .LineStartCap = .LineEndCap
    .LineEndColor = .LineStartColor
    .LineVisible = bVisible
  End With

Else
  With m_Col(lColumn).Cell(ItemStart)
    .CurveTo = ItemEnd
    .Direction = IIf(ItemStart > ItemEnd, 0, 1)
    .SideTo = lSide
    .Radius = iRadius
    .LineStyle = lStyle
    .LineWidth = lLineWidth
    .LineStartCap = lStartCap
    .LineEndCap = lEndCap
    .LineStartColor = cStartColor
    .LineEndColor = cEndColor
    .LineOpacity = lOpacity
    .LineVisible = bVisible
  End With
  
End If

If bVisible Then Refresh
End Function

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

Public Sub Clear()
    Erase m_Col
    Refresh
End Sub

Public Sub Refresh()
Dim C As Long
On Error Resume Next

    With ucScrollV
      .Max = ((m_ItemHeight + m_BorderWidth) * MaxItemCount) - (UserControl.ScaleHeight - 50)
      If .Max > 0 Then
        .Visible = True
        .TrackMouseWheelOnHwnd UserControl.hwnd
      Else
        .Visible = False
        .TrackMouseWheelOnHwndStop
      End If
    End With
    
    With ucScrollH
      C = UBound(m_Col)
      If C <= 0 Then Exit Sub
      .Max = (m_Col(C).StartX + m_Col(C).Width + 50) - (UserControl.ScaleWidth)
      If .Max > 0 Then
        .Visible = True
        .TrackMouseWheelOnHwnd UserControl.hwnd
      Else
        .Visible = False
        .TrackMouseWheelOnHwndStop
      End If
    End With
    
    UserControl.Cls
    Draw

End Sub

Public Sub RemoveItem(ByVal lColumn As Long, ByVal Item As Long)
On Local Error Resume Next
Dim j As Long, L As Long, C As Long

C = IIf(lColumn = 0, 1, 0)

With m_Col(lColumn)
    
  If ItemCount(lColumn) > 1 Then
      For j = Item To UBound(.Cell) - 1
         LSet .Cell(j) = .Cell(j + 1)
      Next
      ReDim Preserve .Cell(UBound(.Cell) - 1)

  Else
      Erase .Cell
  End If
  
End With

With m_Col(C)
  For j = 0 To UBound(.Cell)
     .Cell(j).LineTo = IIf(.Cell(j).LineTo >= Item, IIf(.Cell(j).LineTo = Item, -1, .Cell(j).LineTo - 1), .Cell(j).LineTo)
  Next
End With

With m_Col(lColumn)
  For j = 0 To UBound(.Cell)
     .Cell(j).CurveTo = IIf(.Cell(j).CurveTo >= Item, IIf(.Cell(j).CurveTo = Item, -1, .Cell(j).CurveTo - 1), .Cell(j).CurveTo)
  Next
End With

'*2
 m_ActiveItem(lColumn) = -1
 m_SelectItem(lColumn) = -1
 
 Refresh
 
End Sub

Private Function MousePointerHands(ByVal NewValue As Boolean)
  If NewValue Then
    If Ambient.UserMode Then
      UserControl.MousePointer = vbCustom
      UserControl.MouseIcon = GetSystemHandCursor
    End If
  Else
    If hCur Then DestroyCursor hCur: hCur = 0
    UserControl.MousePointer = vbDefault
    UserControl.MouseIcon = Nothing
  End If

End Function

Private Function GetSystemHandCursor() As Picture
  Dim Pic As PicBmp
  Dim IPic As IPicture
  Dim GUID(0 To 3) As Long
  
  If hCur Then DestroyCursor hCur: hCur = 0
  
  hCur = LoadCursor(ByVal 0&, IDC_HAND)
   
  GUID(0) = &H7BF80980
  GUID(1) = &H101ABF32
  GUID(2) = &HAA00BB8B
  GUID(3) = &HAB0C3000
  
  With Pic
    .Size = Len(Pic)
    .type = vbPicTypeIcon
    .hBmp = hCur
    .hpal = 0
  End With
  
  Call OleCreatePictureIndirect(Pic, GUID(0), 1, IPic)
  
  Set GetSystemHandCursor = IPic
End Function

Private Sub Draw()
Dim rHeader As RECTL, sHeader As RECTS
Dim CELDA As RECTL, IcoBox As RECTS
Dim cp1REC As RECTS, cp2REC As RECTS
Dim rLabel As RECTL, SLabel As RECTS
Dim IcoBoxL As RECTL
Dim i As Long, C As Long
Dim lY As Long, lX As Long
Dim TopH As Long, lMargen As Long
Dim mBorder As Long, lBorder As Long
Dim ucTH As Long
Dim pGen As Long

If ColCount <= 0 Then Exit Sub
If mRedraw = False Then Exit Sub

  GdipCreateFromHDC hdc, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

  'Valores Bordes
  lBorder = m_BorderWidth * 2 * nScale
  mBorder = m_BorderWidth * nScale

  lMargen = (m_CornerCurve / Screen.TwipsPerPixelX) + 5
  
  'Control BackColor
  UserControl.BackColor = m_BackColor
  
  'Scroll Value
  lY = -ucScrollV.Value
  lX = -ucScrollH.Value
  
  ucTH = UserControl.TextHeight("Áhj") + 10 * nScale
  
For C = 0 To ColCount - 1
'*1
  With m_Col(C)

    ''DRAW-HEADERS-------------------------------------------------------------------------------------
    SetRECL rHeader, lX + mBorder + .StartX, lY + lBorder + .StartY, .Width, m_ItemHeight - lBorder
    SetRECS sHeader, rHeader.Left, rHeader.Top, rHeader.Width, rHeader.Height
      
    DrawRoundRect hGraphics, rHeader, RGBA(.BackColor, SafeRange(m_BoxOpacity + 10, 0, 100)), RGBA(m_BorderColor, m_BoxOpacity), m_BorderWidth, m_CornerCurve, True
    DrawCaption hGraphics, m_Col(C).Header, m_HeaderFont, sHeader, RGBA(.ForeColor, m_FontOpacity), 0, eCenter, eMiddle, 0, 0, False

    ''DRAW COLUMNS---------------------------------------------------------------------------------
    i = 0
    Do While i <= ItemCount(C) - 1 And lY < UserControl.ScaleHeight
      
          pGen = (m_ItemHeight * i) + m_ItemHeight

          If TopH + m_ItemHeight > 0 Then

              'CELDA
              .Cell(i).PointLX = lX + .StartX - mBorder
              .Cell(i).PointY = lY + pGen + lBorder + (m_ItemHeight / 2) + .StartY
              .Cell(i).PointRX = lX + .StartX + (IIf(.Width >= 0, .Width, m_ColWidth)) + mBorder

              SetRECL CELDA, lX + mBorder + .StartX, lY + pGen + lBorder + .StartY, .Width, m_ItemHeight - lBorder

              'Label
              SetRECL rLabel, CELDA.Left + ((CELDA.Width / 3) * 2), CELDA.Top + 5, (CELDA.Width / 3) - 5, ucTH - 5
              SetRECS SLabel, CELDA.Left + ((CELDA.Width / 3) * 2), CELDA.Top + 5, (CELDA.Width / 3) - 5, ucTH - 5
              'IconBox
              SetRECS IcoBox, CELDA.Left + ((CELDA.Width / 3) * 2), CELDA.Top + 5, (CELDA.Width / 3) - 5, CELDA.Height - 10
              'Text
              SetRECS cp1REC, CELDA.Left + lMargen, rLabel.Top + mBorder, CELDA.Width - lMargen, CELDA.Height / 2 'ucTH / 2
              'SubText
              SetRECS cp2REC, CELDA.Left + lMargen, rLabel.Top + rLabel.Height, CELDA.Width - lMargen, CELDA.Height / 2 'CELDA.Height - ucTH
              'Draw CELDA
              DrawRoundRect hGraphics, CELDA, RGBA(IIf(i = m_ActiveItem(C), m_ColorActive, m_BoxColor), SafeRange(m_BoxOpacity + 10, 0, 100)), RGBA(IIf(i = m_SelectItem(C), m_BorderColorActive, m_BorderColor), m_BoxOpacity), m_BorderWidth, m_CornerCurve, True
              'Text
              DrawCaption hGraphics, .Cell(i).Text, m_Font1, cp1REC, RGBA(m_ForeColor1, m_FontOpacity), 0, m_CaptionAlignH, m_CaptionAlignV, 0, 0, False

              Select Case m_SubTextVisible
                Case Is = scNone
                    GoTo dSubText02
                Case Is = scAllways
                    GoTo dSubText01
                Case Is = scOnlyActive
                  If i = m_SelectItem(C) Then
                    GoTo dSubText01
                  Else
                    GoTo dSubText02
                  End If
              End Select

dSubText01:
              'SubText
              DrawCaption hGraphics, .Cell(i).SubText, m_Font2, cp2REC, RGBA(m_ForeColor2, m_FontOpacity), 0, m_SubTextAlignH, m_SubTextAlignV, 0, 0, False
dSubText02:
              Select Case m_BadgeVisible
                Case Is = scAllways
                    GoTo dDrawIcon0
                Case Is = scOnlyActive
                  If i = m_SelectItem(C) Then
                    GoTo dDrawIcon0
                  Else
                    GoTo dNoDrawIcon0
                  End If
              End Select
dDrawIcon0:
              'Label/Badge
              If m_BadgeType = vbLabel Then
                DrawRoundRect hGraphics, rLabel, RGBA(m_BorderColor, m_BoxOpacity), RGBA(m_BorderColor, m_BoxOpacity), 1, m_CornerCurve, True
                DrawCaption hGraphics, .Cell(i).Label, m_LabelFont, SLabel, RGBA(m_ForeColor1, m_FontOpacity), 0, eCenter, eMiddle, 0, 0, False
              ElseIf m_BadgeType = vbIcon Then
              'IconChar
                If m_ImageType = eIconFont Then
                  DrawCaption hGraphics, .Cell(i).IconChar, IconFont, IcoBox, RGBA(m_IconForeColor, 100), 0, eCenter, eMiddle, 0, 0, True
                Else
                  'SetRECL IcoBoxL, CLng(IcoBox.Left + (IcoBox.Width / 4)), CLng(IcoBox.Top + (IcoBox.Height / 4)), CLng(IcoBox.Width / 2), CLng(IcoBox.Height / 2)
                  SetRECL IcoBoxL, CLng(IcoBox.Left), CLng(IcoBox.Top), CLng(IcoBox.Width), CLng(IcoBox.Height)
                  LoadPictureFromFile .Cell(i).IconFile
                  Call GdipSetInterpolationMode(hGraphics, 7&)  'HIGH_QUALYTY_BICUBIC
                  Call GdipSetPixelOffsetMode(hGraphics, 4&)
                  GdipDrawImageRectI hGraphics, m_Bmp, IcoBoxL.Left, IcoBoxL.Top, IcoBoxL.Width, IcoBoxL.Height
                End If
              End If

dNoDrawIcon0:
            TopH = TopH + m_ItemHeight
          End If

      i = i + 1
    Loop
  End With
Next C

'*3'-DRAW-CONNECTION-LINES---------------------------------------------------------------------------------------
For C = 0 To ColCount - 1
  i = 0
  Do While i <= ItemCount(C) - 1 And lY < UserControl.ScaleHeight
    With m_Col(C).Cell(i)
    
      '-DRAW-LINES---------------------------------------------------------------------------------------
      If .LineTo > -1 Then
        m_X1 = IIf(.SideTo = toLeft, .PointLX, .PointRX)
        m_Y1 = .PointY + .LineWidth
        m_X2 = IIf(.SideTo = toLeft, m_Col(.ColumnTo).Cell(.LineTo).PointRX, m_Col(.ColumnTo).Cell(.LineTo).PointLX)
        m_Y2 = m_Col(.ColumnTo).Cell(.LineTo).PointY - .LineWidth
        DrawLine hGraphics, m_X1, m_Y1, m_X2, m_Y2, .LineStyle, .LineWidth, .LineStartCap, .LineEndCap, .LineStartColor, .LineEndColor, .LineOpacity
      End If
      '-DRAW-CURVES---------------------------------------------------------------------------------------
      If .CurveTo > -1 Then
        m_X1 = IIf(.SideTo = toLeft, .PointLX, .PointRX)
        m_Y1 = .PointY + .LineWidth
        m_X2 = IIf(.SideTo = toLeft, m_Col(C).Cell(.CurveTo).PointLX, m_Col(C).Cell(.CurveTo).PointRX)
        m_Y2 = m_Col(C).Cell(.CurveTo).PointY - .LineWidth
        DrawCurve hGraphics, m_X1, m_Y1, m_X2, m_Y2, .Radius, .SideTo, .Direction, .LineStyle, .LineWidth, .LineStartCap, .LineEndCap, .LineStartColor, .LineEndColor, .LineOpacity
      End If
    
      i = i + 1
    End With
  Loop
Next C

'------------------------------------------------------------------------
 GdipDeleteGraphics hGraphics
'*1----------------------------------------------------------------------
End Sub

Private Function DrawBubble(ByVal hGraphics As Long, RCT As RECTL, BorderColor As Long, BorderWidth As Long, BackColor As Long, lCurve As Long, coWidth As Long, coLen As Long, COPos As CallOutPosition) As Long
    Dim mpath As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mRound As Long
    Dim Xx As Long, Yy As Long
    Dim lMax As Long
    Dim coAngle  As Long

With RCT
        
    coAngle = coWidth / 2

    mRound = GetSafeRound(lCurve * nScale, .Width, .Height)
    
    Select Case COPos
        Case coLeft
            .Left = .Left + coLen
            .Width = .Width - coLen
            lMax = .Height - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coTop
            .Top = .Top + coLen
            .Height = .Height - coLen
            lMax = .Width - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coRight
            .Width = .Width - coLen
            lMax = .Height - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
        Case coBottom
            .Height = .Height - coLen
            lMax = .Width - (mRound * 2)
            If coWidth > lMax Then coWidth = lMax
    End Select

    If BorderWidth >= 1 Then GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, hPen
    GdipCreateSolidFill BackColor, hBrush
    Call GdipCreatePath(&H0, mpath)
                    
    GdipAddPathArcI mpath, .Left, .Top, mRound * 2, mRound * 2, 180, 90

    If COPos = coTop Then
        Xx = .Left + (.Width - coWidth) / 2
        If mRound = 0 Then GdipAddPathLineI mpath, .Left, .Top, .Left, .Top
        GdipAddPathLineI mpath, Xx, .Top, Xx + coAngle, .Top - coLen
        GdipAddPathLineI mpath, Xx + coAngle, .Top - coLen, Xx + coWidth, .Top
    End If

    GdipAddPathArcI mpath, .Left + .Width - mRound * 2, .Top, mRound * 2, mRound * 2, 270, 90

    If COPos = coRight Then
        Yy = .Top + (.Height - coWidth) / 2
        Xx = .Left + .Width
        If mRound = 0 Then GdipAddPathLineI mpath, .Left + .Width, .Top, .Left + .Width, .Top
        GdipAddPathLineI mpath, Xx, Yy, Xx + coLen, Yy + coAngle
        GdipAddPathLineI mpath, Xx + coLen, Yy + coAngle, Xx, Yy + coWidth
    End If

    GdipAddPathArcI mpath, .Left + .Width - mRound * 2, .Top + .Height - mRound * 2, mRound * 2, mRound * 2, 0, 90

    If COPos = coBottom Then
        Xx = .Left + (.Width - coWidth) / 2
        Yy = .Top + .Height
        If mRound = 0 Then GdipAddPathLineI mpath, .Left + .Width, .Top + .Height, .Left + .Width, .Top + .Height
        GdipAddPathLineI mpath, Xx + coWidth, Yy, Xx + coAngle, Yy + coLen
        GdipAddPathLineI mpath, Xx + coAngle, Yy + coLen, Xx, Yy
    End If

    GdipAddPathArcI mpath, .Left, .Top + .Height - mRound * 2, mRound * 2, mRound * 2, 90, 90
    
    If COPos = coLeft Then
        Yy = .Top + (.Height - coWidth) / 2
        If mRound = 0 Then GdipAddPathLineI mpath, .Left, .Top + .Height, .Left, .Top + .Height
        GdipAddPathLineI mpath, .Left, Yy + coWidth, .Left - coLen, Yy + coAngle
        GdipAddPathLineI mpath, .Left - coLen, Yy + coAngle, .Left, Yy
    End If
End With
        
    GdipClosePathFigures mpath
    GdipFillPath hGraphics, hBrush, mpath
    If BorderWidth >= 1 Then GdipDrawPath hGraphics, hPen, mpath
    
    Call GdipDeletePath(mpath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)

End Function

Private Function DrawCaption(ByVal hGraphics As Long, sString As Variant, oFont As StdFont, layoutRect As RECTS, _
                             TextColor As Long, mAngle As Single, HAlign As eTextAlignH, VAlign As eTextAlignV, _
                             CapX As Long, CapY As Long, Icon As Boolean) As Long
Dim hPath As Long
Dim hBrush As Long
Dim hFontFamily As Long
Dim hFormat As Long
Dim lFontSize As Long
Dim lFontStyle As GDIPLUS_FONTSTYLE
Dim newY As Long, newX As Long

On Error Resume Next

    If GdipCreatePath(&H0, hPath) = 0 Then

        If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
            GdipSetStringFormatTrimming hFormat, StringTrimmingEllipsisWord
            GdipSetStringFormatAlign hFormat, HAlign
            GdipSetStringFormatLineAlign hFormat, VAlign
        End If

        GetFontStyleAndSize oFont, lFontStyle, lFontSize

        If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
            If hFontCollection Then
                If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), hFontCollection, hFontFamily) Then
                    If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
                End If
            Else
                If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
            End If
        End If
'------------------------------------------------------------------------
        If mAngle <> 0 Then
            newY = (layoutRect.Height / 2)
            newX = (layoutRect.Width / 2)
            Call GdipTranslateWorldTransform(hGraphics, newX, newY, 0)
            Call GdipRotateWorldTransform(hGraphics, mAngle, 0)
            Call GdipTranslateWorldTransform(hGraphics, -newX, -newY, 0)
        End If
'------------------------------------------------------------------------
         layoutRect.Left = layoutRect.Left + CapX
         layoutRect.Top = layoutRect.Top + CapY
'------------------------------------------------------------------------
      If Icon Then
        GdipAddPathString hPath, StrPtr(ChrW2(sString)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
      Else
        GdipAddPathString hPath, StrPtr(sString), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
      End If
'------------------------------------------------------------------------
        GdipDeleteStringFormat hFormat
        GdipCreateSolidFill TextColor, hBrush
        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush
        If mAngle <> 0 Then GdipResetWorldTransform hGraphics
        GdipDeleteFontFamily hFontFamily
        GdipDeletePath hPath
    End If

End Function

Private Sub DrawLine(hGraphics As Long, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                     fLineStyle As DashStyle, fLineWidth As Single, _
                     fStarCap As LineCap, fEndCap As LineCap, _
                     fStartColor As OLE_COLOR, fEndColor As OLE_COLOR, fOpacity As Long)
Dim hPen As Long
Dim hBrush As Long
Dim lP1 As POINTS
Dim lP2 As POINTS
    
  lP1.X = X1: lP1.Y = Y1
  lP2.X = X2: lP2.Y = Y2

  GdipCreatePen1 RGBA(fStartColor, fOpacity), fLineWidth * nScale, UnitPixel, hPen
  GdipCreateLineBrush lP1, lP2, RGBA(fStartColor, fOpacity), RGBA(fEndColor, fOpacity), WrapModeTileFlipXY, hBrush
  
  GdipSetPenBrushFill hPen, hBrush
  GdipSetPenDashStyle hPen, fLineStyle
  GdipSetPenStartCap hPen, fStarCap
  GdipSetPenEndCap hPen, fEndCap
  
  GdipDrawLine hGraphics, hPen, lP1.X, lP1.Y, lP2.X, lP2.Y
  
  Call GdipDeleteBrush(hBrush)
  Call GdipDeletePen(hPen)
    
End Sub

Private Sub DrawCurve(hGraphics As Long, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                     iRadius As Single, lSide As tSide, lDirection As tDirection, _
                     fLineStyle As DashStyle, fLineWidth As Single, _
                     fStarCap As LineCap, fEndCap As LineCap, _
                     fStartColor As OLE_COLOR, fEndColor As OLE_COLOR, fOpacity As Long)
Dim hPen As Long
Dim hBrush As Long
Dim lP1 As POINTS
Dim lP2 As POINTS
Dim lP3 As POINTS
Dim lP4 As POINTS
    
  lP1.X = X1: lP1.Y = Y1
  lP4.X = X2: lP4.Y = Y2
  
  Dim xSide As tSide
  
  If lSide = Auto Then
    If X1 >= 0 And X1 <= UserControl.ScaleWidth - (m_ColWidth + (iRadius * 1.5)) Then
      xSide = toRight
    ElseIf X1 >= m_ColWidth + (iRadius * 1.5) And X1 <= UserControl.ScaleWidth - m_ColWidth Then
      xSide = toLeft
    End If
  Else
    xSide = lSide
  End If
  
  lP2.X = X1 + IIf(xSide = toRight, iRadius, -iRadius)
  lP2.Y = Y1 + IIf(lDirection = 0, -(iRadius / 2), (iRadius / 2))
  lP3.X = X2 + IIf(xSide = toRight, iRadius, -iRadius)
  lP3.Y = Y2 + IIf(lDirection = 0, (iRadius / 2), -(iRadius / 2))

  GdipCreatePen1 RGBA(fStartColor, fOpacity), fLineWidth * nScale, UnitPixel, hPen
  GdipCreateLineBrush lP1, lP2, RGBA(fStartColor, fOpacity), RGBA(fEndColor, fOpacity), WrapModeTileFlipXY, hBrush
  
  GdipSetPenBrushFill hPen, hBrush
  GdipSetPenDashStyle hPen, fLineStyle
  GdipSetPenStartCap hPen, fStarCap
  GdipSetPenEndCap hPen, fEndCap
  
  GdipDrawBezier hGraphics, hPen, lP1.X, lP1.Y, lP2.X, lP2.Y, lP3.X, lP3.Y, lP4.X, lP4.Y
  
  Call GdipDeleteBrush(hBrush)
  Call GdipDeletePen(hPen)
    
End Sub

Private Function DrawRoundRect(ByVal hGraphics As Long, Rect As RECTL, ByVal BackColor As Long, _
                               ByVal BorderColor As Long, ByVal BorderWidth As Long, _
                               ByVal Round As Long, Filled As Boolean) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mpath As Long
    Dim mRound As Long
    
    If m_BorderWidth > 0 Then GdipCreatePen1 BorderColor, BorderWidth * nScale, &H2, hPen
    If Filled Then GdipCreateSolidFill BackColor, hBrush
    'GdipCreateLineBrushFromRectWithAngleI Rect, BackColor, BackColor, 90, 0, WrapModeTileFlipXY, hBrush
    
    GdipCreatePath &H0, mpath   '&H0
    
    With Rect
        mRound = GetSafeRound((Round * nScale), .Width * 2, .Height * 2)
        If mRound = 0 Then mRound = 1
            GdipAddPathArcI mpath, .Left, .Top, mRound, mRound, 180, 90
            GdipAddPathArcI mpath, (.Left + .Width) - mRound, .Top, mRound, mRound, 270, 90
            GdipAddPathArcI mpath, (.Left + .Width) - mRound, (.Top + .Height) - mRound, mRound, mRound, 0, 90
            GdipAddPathArcI mpath, .Left, (.Top + .Height) - mRound, mRound, mRound, 90, 90
    End With
    
    GdipClosePathFigures mpath
    GdipFillPath hGraphics, hBrush, mpath
    GdipDrawPath hGraphics, hPen, mpath
    
    Call GdipDeletePath(mpath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)
End Function

Private Function GetFontStyleAndSize(oFont As StdFont, lFontStyle As Long, lFontSize As Long)
On Error GoTo ErrO
    Dim hdc As Long
    lFontStyle = 0
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
    
    hdc = GetDC(0&)
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hdc, LOGPIXELSY), 72)
    ReleaseDC 0&, hdc
ErrO:
End Function

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function

Private Function GetItem(lColumn As Long, ByVal Y As Single) As Long
'On Error Resume Next
    Y = (Y + ucScrollV.Value) - (m_Col(lColumn).StartY + m_ItemHeight)
    GetItem = (Y \ m_ItemHeight)
    If GetItem >= ItemCount(lColumn) Or GetItem < 0 Then GetItem = -1
End Function

Private Function GetColumn(ByVal X As Single, ByVal Y As Single) As Long
Dim C As Long

X = (X + ucScrollH.Value)
For C = 0 To UBound(m_Col)
  With m_Col(C)
    If X >= (.StartX) And X <= (.StartX + .Width) And Y >= .StartY And Y <= .StartY + (ItemCount(C) * m_ItemHeight) Then
      GetColumn = C
      Debug.Print "GetColumn:" & C
      Exit For
    Else
      GetColumn = -1
    End If
  End With
Next C

End Function

Private Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double, LPY As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    LPY = CDbl(GetDeviceCaps(hdc, LOGPIXELSY))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Function IconCharCode(ByVal New_IconCharCode As String) As Long
  If Trim$(New_IconCharCode) <> "" Or New_IconCharCode <> vbNullString Then
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not VBA.Left$(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        IconCharCode = "&H" & New_IconCharCode
    Else
        IconCharCode = New_IconCharCode
    End If
  End If
End Function

Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub

Private Function LoadPictureFromFile(ByVal FileName As String) As Boolean
Dim imgW As Long, imgH As Long
Dim BmpW As Single, BmpH As Single
Dim Bmp  As Long, Grph As Long

If m_Bmp Then
    Call GdipDisposeImage(m_Bmp)
    m_Bmp = 0
End If

imgW = IcoBoxL.Width
imgH = IcoBoxL.Height

Call GdipLoadImageFromFile(StrPtr(FileName), m_Bmp)

If m_Bmp <> 0 Then
    GdipGetImageDimension m_Bmp, BmpW, BmpH
    '-->
    If GdipCreateBitmapFromScan0(imgW, imgH, 0&, &HE200B, ByVal 0&, Bmp) = 0 Then
      If GdipGetImageGraphicsContext(Bmp, Grph) = 0 Then

          If imgW > BmpW Or imgH > BmpH Then
              Call GdipSetInterpolationMode(Grph, 5&)  '// IterpolationModeNearestNeighbor
          Else
              Call GdipSetInterpolationMode(Grph, 7&)  '//InterpolationModeHighQualityBicubic
              Call GdipSetPixelOffsetMode(Grph, 4&)
          End If

          Call GdipDrawImageRectRectI(Grph, m_Bmp, 0, 0, imgW, imgH, 0, 0, BmpW, BmpH, &H2)
          GdipDeleteGraphics Grph

          Call GdipDisposeImage(m_Bmp)
          m_Bmp = Bmp
          '...Draw Image on Control
          'Draw 'Don´t Call the Function because is called inside it...
          LoadPictureFromFile = True
      End If
    End If
Else
    LoadPictureFromFile = False
End If
End Function

Private Function ReadValue(ByVal lProp As Long, Optional Default As Long) As Long
    Dim i       As Long
    For i = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(i) = lProp Then
            ReadValue = TlsGetValue(i + 1)
            Exit Function
        End If
    Next
    ReadValue = Default
End Function

Private Function RGBA(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
  If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
  RGBA = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
  Opacity = CByte((Abs(Opacity) / 100) * 255)
  If Opacity < 128 Then
      If Opacity < 0& Then Opacity = 0&
      RGBA = RGBA Or Opacity * &H1000000
  Else
      If Opacity > 255& Then Opacity = 255&
      RGBA = RGBA Or (Opacity - 128&) * &H1000000 Or &H80000000
  End If
End Function

Private Function SafeRange(Value, Min, Max) As Long
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
    SafeRange = Value
End Function

Private Function SetRECS(lpRect As RECTS, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long) As Long
  lpRect.Left = X
  lpRect.Top = Y
  lpRect.Width = W
  lpRect.Height = H
End Function

Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub ucScrollH_Change()
  Refresh
End Sub

Private Sub ucScrollH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ucScrollV.TrackMouseWheelOnHwndStop
  ucScrollH.TrackMouseWheelOnHwnd UserControl.hwnd
End Sub

Private Sub ucScrollH_Scroll()
  Refresh
End Sub

Private Sub ucScrollV_Change()
  Refresh
End Sub

Private Sub ucScrollV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ucScrollH.TrackMouseWheelOnHwndStop
  ucScrollV.TrackMouseWheelOnHwnd UserControl.hwnd
End Sub

Private Sub ucScrollV_Scroll()
  Refresh
End Sub

Private Sub UserControl_Initialize()
    InitGDI
    nScale = GetWindowsDPI
    SideStart = -1
End Sub

Private Sub UserControl_InitProperties()
hFontCollection = ReadValue(&HFC)
  
  m_Enabled = True
  m_Clickable = False
  mRedraw = True
  m_Editable = False
  
  m_BorderColor = &HFF8080
  m_BackColor = &H8000000F
  m_BoxColor = vbRed
  m_BorderWidth = 1
  m_ForeColor1 = &HFFFFFF
  m_ForeColor2 = &HFFFFFF
  m_IconForeColor = &HFFFFFF
  Set m_Font1 = UserControl.Font
  Set m_Font2 = UserControl.Font
  Set m_IconFont = UserControl.Font
  Set m_LabelFont = UserControl.Font
  m_CaptionAlignV = 1
  m_CaptionAlignH = 1
  m_SubTextAlignV = 1
  m_SubTextAlignH = 1
  'm_IconAlignV = 1
  'm_IconAlignH = 1
  'm_ImageType = 0
  'm_BadgeVisible = 1
  m_SubTextVisible = 1
  m_ItemHeight = 60
  m_ColWidth = 150
  m_ColorActive = vbRed
  m_BadgeType = False
  m_CornerCurve = 5
  m_BoxOpacity = 90
 
 
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case 37, 38
      ucScrollV.Value = ucScrollV.Value - (m_ItemHeight / 2)
      'If m_Clickable Then m_ActiveItem = IIf(m_ActiveItem = 0, 0, m_ActiveItem - 1)
    
  Case 39, 40
      ucScrollV.Value = ucScrollV.Value + (m_ItemHeight / 2)
      'If m_Clickable Then m_ActiveItem = IIf(m_ActiveItem = ItemCount(ActiveSide) - 1, ItemCount(ActiveSide) - 1, m_ActiveItem + 1)
    
End Select
  
RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_DblClick()
On Error Resume Next

If m_Editable = True Then

  m_ActiveCol = GetColumn(mX, mY)
  m_ActiveItem(m_ActiveCol) = GetItem(m_ActiveCol, mY)

  If SideStart = -1 Then
  
    SideStart = m_ActiveCol
    InitCurve = m_ActiveItem(m_ActiveCol)
    Debug.Print "StartDraw : SideStart " & SideStart
  
  Else
  
    SideEnd = m_ActiveCol
    
    If SideStart = SideEnd Then
    
      Debug.Print "[DrawCurve] SideStart: " & SideStart & " ActiveSide: " & m_ActiveCol & " SideTo: " & IIf(m_ActiveCol > 0, "toLeft", "toRight")
      EndCurve = m_ActiveItem(m_ActiveCol)
      AddCurve SideStart, InitCurve, EndCurve, IIf(m_DefaultSide = Auto, IIf(m_ActiveCol > 0, toLeft, toRight), m_DefaultSide), 50, m_LineStyle, m_LineStartCap, m_LineEndCap, m_LineStartColor, m_LineEndColor, m_LineWidth, m_LineOpacity, True
      SideStart = -1
      
    Else
    
      Debug.Print "DrawLine"
      If m_ActiveItem(SideStart) > -1 And m_ActiveItem(m_ActiveCol) > -1 Then
        AddLine SideStart, m_ActiveItem(SideStart), SideEnd, m_ActiveItem(m_ActiveCol), m_LineStyle, m_LineStartCap, m_LineEndCap, m_LineStartColor, m_LineEndColor, m_LineWidth, m_LineOpacity, True
        SideStart = -1
      End If
  
    End If
    
  End If
  
Else

  m_ActiveItem(SideStart) = -1
  m_ActiveItem(SideEnd) = -1

End If

Refresh

RaiseEvent DblClick(m_ActiveCol, m_SelectItem(m_ActiveCol))
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

mY = Y
mX = X
mCol = GetColumn(X, Y)


If mCol > -1 Then
  m_ActiveCol = mCol
  m_SelectItem(mCol) = -1
  m_ActiveItem(mCol) = -1

  If m_Clickable Then
    m_SelectItem(mCol) = GetItem(mCol, Y)
    Debug.Print "m_SelectItem(" & mCol & ")=" & m_SelectItem(mCol)

    RaiseEvent MouseDown(Button, Shift, X, Y)
    With m_Col(mCol)
      If m_SelectItem(mCol) <> -1 Then RaiseEvent Click(mCol, GetItem(mCol, Y), .Cell(GetItem(mCol, Y)).ColumnTo, .Cell(GetItem(mCol, Y)).LineTo, .Cell(GetItem(mCol, Y)).CurveTo)
    End With
    Refresh
  End If
End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_Editable Then
  MousePointerHands True
Else
  MousePointerHands False
End If

If m_ColsMoveable Then
  On Error Resume Next
  m_Col(mCol).StartX = X - (m_Col(mCol).Width / 2)
  m_Col(mCol).StartY = Y - (m_Col(mCol).Cell(0).Height / 2)
  Refresh
End If

RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_ColsMoveable Then mCol = -1
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  m_Enabled = .ReadProperty("Enabled", True)
  m_Clickable = .ReadProperty("Clickable", False)
  m_Editable = .ReadProperty("Editable", False)
  
  m_BorderColor = .ReadProperty("BorderColor", &HFF8080)
  m_BorderColorActive = .ReadProperty("BorderColorActive", vbRed)
  m_BackColor = .ReadProperty("BackColor", &H8000000F)
  m_BoxColor = .ReadProperty("BoxColor", vbRed)
  m_BorderWidth = .ReadProperty("BorderWidth", 1)
  m_ForeColor1 = .ReadProperty("TextColor", &HFFFFFF)
  m_ForeColor2 = .ReadProperty("SubTextColor", &HFFFFFF)
  m_HeaderBackColor = .ReadProperty("HeaderBackColor", vbBlack)
  m_HeaderForeColor = .ReadProperty("HeaderForeColor", vbWhite)
  Set m_Font1 = .ReadProperty("TextFont", UserControl.Font)
  Set m_Font2 = .ReadProperty("SubTextFont", UserControl.Font)
  Set m_LabelFont = .ReadProperty("LabelFont", UserControl.Font)
  Set m_HeaderFont = .ReadProperty("HeaderFont", UserControl.Font)
  m_FontOpacity = .ReadProperty("FontOpacity", 100)
  m_CaptionAlignV = .ReadProperty("TextAlignV", 1)
  m_CaptionAlignH = .ReadProperty("TextAlignH", 1)
  m_SubTextAlignV = .ReadProperty("SubTextAlignV", 1)
  m_SubTextAlignH = .ReadProperty("SubTextAlignH", 1)
  m_IconForeColor = .ReadProperty("IconForeColor", &HFFFFFF)
  Set m_IconFont = .ReadProperty("IconFont", UserControl.Font)
  'm_IconAlignV = .ReadProperty("IconAlignV", 1)
  'm_IconAlignH = .ReadProperty("IconAlignH", 1)
  m_ImageType = .ReadProperty("ImageType", 0)
  m_BadgeVisible = .ReadProperty("BadgeVisible", 1) 'BadgeVisible
  m_SubTextVisible = .ReadProperty("SubTextVisible", 1)
  m_ItemHeight = .ReadProperty("ItemHeight", 60)
  m_ColWidth = .ReadProperty("ColWidth", 150)
  m_ColsMoveable = .ReadProperty("ColsMoveable", False)
  m_ColorActive = .ReadProperty("BackColorActive", vbRed)
  m_BadgeType = .ReadProperty("BadgeType", 0)
  m_CornerCurve = .ReadProperty("CornerCurve", 5)
  m_BoxOpacity = .ReadProperty("BoxOpacity", 90)
  m_DefaultSide = .ReadProperty("LineSideDefault", 1)
  m_LineWidth = .ReadProperty("LineWidth", 5)
  m_LineStyle = .ReadProperty("LineStyle", 0)
  m_LineOpacity = .ReadProperty("LineOpacity", 100)
  m_LineStartCap = .ReadProperty("LineStartCap", 0)
  m_LineEndCap = .ReadProperty("LineEndCap", 0)
  m_LineStartColor = .ReadProperty("LineStartColor", &HFF&)
  m_LineEndColor = .ReadProperty("LineEndColor", &HC67300)
  
End With

End Sub

Private Sub UserControl_Resize()
    ucScrollV.Move UserControl.ScaleWidth - 11, 0, 11, UserControl.ScaleHeight - 11
    ucScrollH.Move 0, UserControl.ScaleHeight - 11, UserControl.ScaleWidth - 11, 11
    If ColCount > 0 Then Refresh
End Sub

Private Sub UserControl_Show()
If ColCount > 0 Then Refresh
End Sub

Private Sub UserControl_Terminate()
TerminateGDI
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("Enabled", m_Enabled)
  Call .WriteProperty("Clickable", m_Clickable)
  Call .WriteProperty("Editable", m_Editable)
  
  Call .WriteProperty("BorderColor", m_BorderColor)
  Call .WriteProperty("BorderColorActive", m_BorderColorActive)
  Call .WriteProperty("BackColor", m_BackColor)
  Call .WriteProperty("BorderWidth", m_BorderWidth)
  Call .WriteProperty("BoxColor", m_BoxColor)
  Call .WriteProperty("TextColor", m_ForeColor1)
  Call .WriteProperty("SubTextColor", m_ForeColor2)
  Call .WriteProperty("HeaderBackColor", m_HeaderBackColor)
  Call .WriteProperty("HeaderForeColor", m_HeaderForeColor)
  Call .WriteProperty("TextFont", m_Font1)
  Call .WriteProperty("SubTextFont", m_Font2)
  Call .WriteProperty("LabelFont", m_LabelFont)
  Call .WriteProperty("HeaderFont", m_HeaderFont)
  Call .WriteProperty("FontOpacity", m_FontOpacity)
  Call .WriteProperty("TextAlignV", m_CaptionAlignV)
  Call .WriteProperty("TextAlignH", m_CaptionAlignH)
  Call .WriteProperty("SubTextAlignV", m_SubTextAlignV)
  Call .WriteProperty("SubTextAlignH", m_SubTextAlignH)
  Call .WriteProperty("IconForeColor", m_IconForeColor)
  Call .WriteProperty("IconFont", m_IconFont)
  'Call .WriteProperty("IconAlignV", m_IconAlignV)
  'Call .WriteProperty("IconAlignH", m_IconAlignH)
  Call .WriteProperty("ImageType", m_ImageType)
  Call .WriteProperty("BadgeVisible", m_BadgeVisible)  'BadgeVisible
  Call .WriteProperty("SubTextVisible", m_SubTextVisible)
  Call .WriteProperty("ItemHeight", m_ItemHeight)
  Call .WriteProperty("ColWidth", m_ColWidth)
  Call .WriteProperty("ColsMoveable", m_ColsMoveable)
  Call .WriteProperty("BackColorActive", m_ColorActive)
  Call .WriteProperty("BadgeType", m_BadgeType)
  Call .WriteProperty("CornerCurve", m_CornerCurve)
  Call .WriteProperty("BoxOpacity", m_BoxOpacity)
  Call .WriteProperty("LineSideDefault", m_DefaultSide)
  Call .WriteProperty("LineWidth", m_LineWidth)
  Call .WriteProperty("LineStyle", m_LineStyle)
  Call .WriteProperty("LineOpacity", m_LineOpacity)
  Call .WriteProperty("LineStartCap", m_LineStartCap)
  Call .WriteProperty("LineEndCap", m_LineEndCap)
  Call .WriteProperty("LineStartColor", m_LineStartColor)
  Call .WriteProperty("LineEndColor", m_LineEndColor)
End With
  
End Sub

Public Property Get BackColor() As OLE_COLOR
  BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
  m_BackColor = New_Color
  PropertyChanged "BackColor"
  Refresh
End Property

Public Property Get BackColorActive() As OLE_COLOR
  BackColorActive = m_ColorActive
End Property

Public Property Let BackColorActive(ByVal NewColorActive As OLE_COLOR)
  m_ColorActive = NewColorActive
  PropertyChanged "BackColorActive"
  Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
  m_BorderColor = NewBorderColor
  PropertyChanged "BorderColor"
  Refresh
End Property

Public Property Get BorderColorActive() As OLE_COLOR
  BorderColorActive = m_BorderColorActive
End Property

Public Property Let BorderColorActive(ByVal NewColorActive As OLE_COLOR)
  m_BorderColorActive = NewColorActive
  PropertyChanged "BorderColorActive"
  Refresh
End Property

Public Property Get BorderWidth() As Long
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  Refresh
End Property

Public Property Get BoxColor() As OLE_COLOR
  BoxColor = m_BoxColor
End Property

Public Property Let BoxColor(ByVal New_Color As OLE_COLOR)
  m_BoxColor = New_Color
  PropertyChanged "BoxColor"
  Refresh
End Property

Public Property Get BoxOpacity() As Long
BoxOpacity = m_BoxOpacity
End Property

Public Property Let BoxOpacity(ByVal newBoxOpacity As Long)
m_BoxOpacity = newBoxOpacity
PropertyChanged "BoxOpacity"
Refresh
End Property

Public Property Get Clickable() As Boolean
  Clickable = m_Clickable
End Property

Public Property Let Clickable(ByVal NewClickable As Boolean)
  m_Clickable = NewClickable
  PropertyChanged "Clickable"
End Property

Public Property Get ColWidth() As Long
  ColWidth = m_ColWidth
End Property

Public Property Let ColWidth(ByVal NewColWidth As Long)
  m_ColWidth = NewColWidth
  PropertyChanged "ColWidth"
  Refresh
End Property

Public Property Get CornerCurve() As Long
  CornerCurve = m_CornerCurve
End Property

Public Property Let CornerCurve(ByVal NewCornerCurve As Long)
  m_CornerCurve = NewCornerCurve
  PropertyChanged "CornerCurve"
  Refresh
End Property

Public Property Get Editable() As Boolean
  Editable = m_Editable
End Property

Public Property Let Editable(ByVal NEditable As Boolean)
  m_Editable = NEditable
  PropertyChanged "Editable"
End Property

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get HeaderBackColor() As OLE_COLOR
  HeaderBackColor = m_HeaderBackColor
End Property

Public Property Let HeaderBackColor(ByVal NewHeaderBackColor As OLE_COLOR)
  m_HeaderBackColor = NewHeaderBackColor
  PropertyChanged "HeaderBackColor"
  Refresh
End Property

Public Property Get HeaderFont() As StdFont
  Set HeaderFont = m_HeaderFont
End Property

Public Property Set HeaderFont(ByVal NewHeaderFont As StdFont)
  Set m_HeaderFont = NewHeaderFont
  PropertyChanged "HeaderFont"
  Refresh
End Property

Public Property Get HeaderForeColor() As OLE_COLOR
  HeaderForeColor = m_HeaderForeColor
End Property

Public Property Let HeaderForeColor(ByVal NewHeaderForeColor As OLE_COLOR)
  m_HeaderForeColor = NewHeaderForeColor
  PropertyChanged "HeaderForeColor"
  Refresh
End Property

Public Property Let Header(ByVal lColumn As Long, ByVal sHeader As String)
On Error Resume Next

With m_Col(lColumn)
  .Header = sHeader
End With

Refresh
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'Public Property Get IconAlignH() As eTextAlignH
'  IconAlignH = m_IconAlignH
'End Property
'
'Public Property Let IconAlignH(ByVal NewIconAlignH As eTextAlignH)
'  m_IconAlignH = NewIconAlignH
'  PropertyChanged "IconAlignH"
'  Refresh
'End Property
'
'Public Property Get IconAlignV() As eTextAlignV
'  IconAlignV = m_IconAlignV
'End Property
'
'Public Property Let IconAlignV(ByVal NewIconAlignV As eTextAlignV)
'  m_IconAlignV = NewIconAlignV
'  PropertyChanged "IconAlignV"
'  Refresh
'End Property

Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property

Public Property Set IconFont(New_Font As StdFont)
  Set m_IconFont = New_Font
    PropertyChanged "IconFont"
  Refresh
End Property

Public Property Get IconForeColor() As OLE_COLOR
    IconForeColor = m_IconForeColor
End Property

Public Property Let IconForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor = New_ForeColor
    PropertyChanged "IconForeColor"
    Refresh
End Property

Public Property Get BadgeVisible() As eVisibleType
BadgeVisible = m_BadgeVisible
End Property

Public Property Let BadgeVisible(ByVal newVisible As eVisibleType)
m_BadgeVisible = newVisible
PropertyChanged "BadgeVisible"
Refresh
End Property

Public Property Get ImageType() As eImageType
  ImageType = m_ImageType
End Property

Public Property Let ImageType(ByVal NewImageType As eImageType)
  m_ImageType = NewImageType
  PropertyChanged "ImageType"
  Refresh
End Property

Property Get ColCount() As Long
On Error GoTo ErrC
    ColCount = UBound(m_Col) + 1
    Exit Property
ErrC:
ColCount = 0
End Property

Property Get ItemCount(lColumn As Long) As Long
On Error Resume Next
    ItemCount = UBound(m_Col(lColumn).Cell) + 1
End Property

Public Property Get ItemHeight() As Long
  ItemHeight = m_ItemHeight
End Property

Public Property Let ItemHeight(ByVal NewSectionSpace As Long)
  m_ItemHeight = NewSectionSpace
  PropertyChanged "ItemHeight"
  Refresh
End Property

Public Property Get LabelFont() As StdFont
  Set LabelFont = m_LabelFont
End Property

Public Property Set LabelFont(ByVal New_Font As StdFont)
  Set m_LabelFont = New_Font
  PropertyChanged "LabelFont"
  Refresh
End Property

Public Property Get BadgeType() As eTypeBadge
  BadgeType = m_BadgeType
End Property

Public Property Let BadgeType(ByVal NewBadgeType As eTypeBadge)
  m_BadgeType = NewBadgeType
  PropertyChanged "BadgeType"
  Refresh
End Property

Property Get LinesCount() As Long
On Local Error Resume Next
    LinesCount = UBound(mLine) + 1
End Property

Property Get MaxItemCount() As Long
Dim Count As Long, CountF As Long, C As Long

CountF = 0
On Error Resume Next

For C = 0 To UBound(m_Col)
  Count = ItemCount(C)
  If Count > CountF Then CountF = Count
Next C

MaxItemCount = CountF

End Property

Public Property Get Redraw() As Boolean
  Redraw = mRedraw
End Property

Public Property Let Redraw(ByVal bRedraw As Boolean)
  mRedraw = bRedraw
  If mRedraw Then Refresh
End Property

Public Property Get RunMode() As Boolean
'Detect usermode even if its used inside another UC
On Error Resume Next
    RunMode = True
    RunMode = Ambient.UserMode
    RunMode = Extender.Parent.RunMode
End Property

Public Property Get SubTextAlignH() As eTextAlignH
  SubTextAlignH = m_SubTextAlignH
End Property

Public Property Let SubTextAlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_SubTextAlignH = NewCaptionAlignH
  PropertyChanged "SubTextAlignH"
  Refresh
End Property

Public Property Get SubTextAlignV() As eTextAlignV
  SubTextAlignV = m_SubTextAlignV
End Property

Public Property Let SubTextAlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_SubTextAlignV = NewCaptionAlignV
  PropertyChanged "SubTextAlignV"
  Refresh
End Property

Public Property Get SubTextColor() As OLE_COLOR
  SubTextColor = m_ForeColor2
End Property

Public Property Let SubTextColor(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor2 = NewForeColor
  PropertyChanged "SubTextColor"
  Refresh
End Property

Public Property Get SubTextFont() As StdFont
  Set SubTextFont = m_Font2
End Property

Public Property Set SubTextFont(ByVal New_Font As StdFont)
  Set m_Font2 = New_Font
  PropertyChanged "SubTextFont"
  Refresh
End Property

Public Property Get SubTextVisible() As eVisibleType
SubTextVisible = m_SubTextVisible
End Property

Public Property Let SubTextVisible(ByVal newVisible As eVisibleType)
m_SubTextVisible = newVisible
PropertyChanged "SubTextVisible"
Refresh
End Property

Public Property Get TextAlignH() As eTextAlignH
  TextAlignH = m_CaptionAlignH
End Property

Public Property Let TextAlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_CaptionAlignH = NewCaptionAlignH
  PropertyChanged "TextAlignH"
  Refresh
End Property

Public Property Get TextAlignV() As eTextAlignV
  TextAlignV = m_CaptionAlignV
End Property

Public Property Let TextAlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_CaptionAlignV = NewCaptionAlignV
  PropertyChanged "TextAlignV"
  Refresh
End Property

Public Property Get TextColor() As OLE_COLOR
  TextColor = m_ForeColor1
End Property

Public Property Let TextColor(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor1 = NewForeColor
  PropertyChanged "TextColor"
  Refresh
End Property

Public Property Get TextFont() As StdFont
  Set TextFont = m_Font1
End Property

Public Property Set TextFont(ByVal New_Font As StdFont)
  Set m_Font1 = New_Font
  PropertyChanged "TextFont"
  Refresh
End Property

Public Property Get Version() As String
Version = sVersion
End Property

Public Property Get Visible() As Boolean
  Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal newVisible As Boolean)
  Extender.Visible = newVisible
End Property

Public Property Get LineWidth() As Single
  LineWidth = m_LineWidth
End Property

Public Property Let LineWidth(ByVal NewLineWidth As Single)
  m_LineWidth = NewLineWidth
  PropertyChanged "LineWidth"
  Refresh
End Property

Public Property Get LineStyle() As DashStyle
  LineStyle = m_LineStyle
End Property

Public Property Let LineStyle(ByVal NewLineStyle As DashStyle)
  m_LineStyle = NewLineStyle
  PropertyChanged "LineStyle"
  Refresh
End Property

Public Property Get LineOpacity() As Long
  LineOpacity = m_LineOpacity
End Property

Public Property Let LineOpacity(ByVal NewLineOpacity As Long)
  m_LineOpacity = NewLineOpacity
  PropertyChanged "LineOpacity"
  Refresh
End Property

Public Property Get LineStartCap() As LineCap
  LineStartCap = m_LineStartCap
End Property

Public Property Let LineStartCap(ByVal NewLineStartCap As LineCap)
  m_LineStartCap = NewLineStartCap
  PropertyChanged "LineStartCap"
  Refresh
End Property

Public Property Get LineEndCap() As LineCap
  LineEndCap = m_LineEndCap
End Property

Public Property Let LineEndCap(ByVal NewLineEndCap As LineCap)
  m_LineEndCap = NewLineEndCap
  PropertyChanged "LineEndCap"
  Refresh
End Property

Public Property Get LineStartColor() As OLE_COLOR
  LineStartColor = m_LineStartColor
End Property

Public Property Let LineStartColor(ByVal NewLineStartColor As OLE_COLOR)
  m_LineStartColor = NewLineStartColor
  PropertyChanged "LineStartColor"
  Refresh
End Property

Public Property Get LineEndColor() As OLE_COLOR
  LineEndColor = m_LineEndColor
End Property

Public Property Let LineEndColor(ByVal NewLineEndColor As OLE_COLOR)
  m_LineEndColor = NewLineEndColor
  PropertyChanged "LineEndColor"
  Refresh
End Property

Public Property Get LineSideDefault() As tSide
  LineSideDefault = m_DefaultSide
End Property

Public Property Let LineSideDefault(ByVal NewDefaultSide As tSide)
  m_DefaultSide = NewDefaultSide
  PropertyChanged "LineSideDefault"
  Refresh
End Property

Public Property Get ColsMoveable() As Boolean
  ColsMoveable = m_ColsMoveable
End Property

Public Property Let ColsMoveable(ByVal NewColsMoveable As Boolean)
  m_ColsMoveable = NewColsMoveable
  PropertyChanged "ColsMoveable"
  Refresh
End Property

Public Property Get FontOpacity() As Long
  FontOpacity = m_FontOpacity
End Property

Public Property Let FontOpacity(ByVal NewFontOpacity As Long)
  m_FontOpacity = NewFontOpacity
  PropertyChanged "FontOpacity"
  Refresh
End Property









'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'  --- All folded content will be temporary put under this lines ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'CODEFOLD STORAGE:
'CODEFOLD STORAGE END:
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'--- If you're Subclassing: Move the CODEFOLD STORAGE up as needed ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\













