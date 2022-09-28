VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "*\AAxConnectionGrid.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   13050
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   150
      TabIndex        =   3
      Top             =   5985
      Width           =   12810
      Begin VB.TextBox txtLineWidth 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6075
         TabIndex        =   54
         Text            =   "5"
         Top             =   1950
         Width           =   390
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00E0E0E0&
         Height          =   645
         Left            =   9555
         TabIndex        =   48
         Top             =   795
         Width           =   855
      End
      Begin VB.ListBox List3 
         BackColor       =   &H00E0E0E0&
         Height          =   645
         Left            =   10560
         TabIndex        =   47
         Top             =   795
         Width           =   855
      End
      Begin VB.ListBox List5 
         Height          =   645
         Left            =   10560
         TabIndex        =   46
         Top             =   1470
         Width           =   855
      End
      Begin VB.ListBox List6 
         Height          =   645
         Left            =   9555
         TabIndex        =   45
         Top             =   1470
         Width           =   855
      End
      Begin VB.ListBox List7 
         Height          =   645
         Left            =   11505
         TabIndex        =   44
         Top             =   1470
         Width           =   1230
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   150
         ScaleHeight     =   195
         ScaleWidth      =   330
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1275
         Width           =   390
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FF8080&
         Height          =   255
         Index           =   1
         Left            =   150
         ScaleHeight     =   195
         ScaleWidth      =   330
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1530
         Width           =   390
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   150
         ScaleHeight     =   195
         ScaleWidth      =   330
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1800
         Width           =   390
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00404040&
         Height          =   255
         Index           =   3
         Left            =   2190
         ScaleHeight     =   195
         ScaleWidth      =   330
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1560
         Width           =   390
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00404040&
         Height          =   255
         Index           =   4
         Left            =   2190
         ScaleHeight     =   195
         ScaleWidth      =   330
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1815
         Width           =   390
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   3915
         ScaleHeight     =   195
         ScaleWidth      =   330
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1860
         Width           =   390
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Index           =   6
         Left            =   3915
         ScaleHeight     =   195
         ScaleWidth      =   330
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1305
         Width           =   390
      End
      Begin VB.TextBox txtBorderWidth 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6075
         TabIndex        =   20
         Text            =   "1"
         Top             =   1620
         Width           =   390
      End
      Begin VB.TextBox txtCurve 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6075
         TabIndex        =   19
         Text            =   "80"
         Top             =   960
         Width           =   390
      End
      Begin VB.TextBox txtColWidth 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6075
         TabIndex        =   18
         Text            =   "150"
         Top             =   300
         Width           =   390
      End
      Begin VB.TextBox txtItemSize 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6075
         TabIndex        =   17
         Text            =   "45"
         Top             =   630
         Width           =   390
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00C67300&
         Height          =   255
         Index           =   7
         Left            =   2190
         ScaleHeight     =   195
         ScaleWidth      =   330
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1290
         Width           =   390
      End
      Begin VB.CheckBox chkEditable 
         Caption         =   "Editable ?"
         Height          =   225
         Left            =   4275
         TabIndex        =   15
         Top             =   435
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.TextBox txtOpacity 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6075
         TabIndex        =   14
         Text            =   "60"
         Top             =   1290
         Width           =   390
      End
      Begin VB.CheckBox chkLabelVisible 
         Caption         =   "Label Visible?"
         Height          =   195
         Left            =   4275
         TabIndex        =   13
         Top             =   915
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.PictureBox pColor 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   3915
         ScaleHeight     =   195
         ScaleWidth      =   330
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1575
         Width           =   390
      End
      Begin VB.CommandButton cmdClearGrid 
         Caption         =   "Clear Grid"
         Height          =   405
         Left            =   2970
         TabIndex        =   11
         Top             =   195
         Width           =   1080
      End
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "Remove Item"
         Height          =   405
         Left            =   135
         TabIndex        =   10
         Top             =   645
         Width           =   2760
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Add Item"
         Height          =   405
         Left            =   150
         TabIndex        =   9
         Top             =   195
         Width           =   2760
      End
      Begin VB.CommandButton cmdReLoad 
         Caption         =   "ReLoad"
         Height          =   405
         Left            =   2970
         TabIndex        =   8
         Top             =   645
         Width           =   1080
      End
      Begin VB.CheckBox chkClickable 
         Caption         =   "Clickable ?"
         Height          =   225
         Left            =   4275
         TabIndex        =   7
         Top             =   195
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   7725
         TabIndex        =   6
         Top             =   450
         Width           =   855
      End
      Begin VB.CheckBox chkColsMoveable 
         Caption         =   "Cols Moveable?"
         Height          =   195
         Left            =   4275
         TabIndex        =   5
         Top             =   675
         Width           =   1575
      End
      Begin VB.ListBox List4 
         Height          =   645
         Left            =   7770
         TabIndex        =   4
         Top             =   1575
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LineWidth"
         Height          =   165
         Left            =   6510
         TabIndex        =   55
         Top             =   1995
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TextAlign Horizontal"
         Height          =   540
         Left            =   9585
         TabIndex        =   53
         Top             =   360
         Width           =   840
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TextAlign Vertical"
         Height          =   540
         Left            =   10620
         TabIndex        =   52
         Top             =   360
         Width           =   720
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text"
         Height          =   195
         Left            =   3375
         TabIndex        =   51
         Top             =   675
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subtext"
         Height          =   195
         Left            =   3345
         TabIndex        =   50
         Top             =   1365
         Width           =   570
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visible?"
         Height          =   195
         Left            =   2190
         TabIndex        =   49
         Top             =   165
         Width           =   720
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ColWidth"
         Height          =   195
         Left            =   6510
         TabIndex        =   43
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BackColor"
         Height          =   195
         Left            =   585
         TabIndex        =   42
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BorderColor"
         Height          =   195
         Left            =   585
         TabIndex        =   41
         Top             =   1575
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BorderColorActive"
         Height          =   195
         Left            =   585
         TabIndex        =   40
         Top             =   1845
         Width           =   1305
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption1 Color"
         Height          =   195
         Left            =   2625
         TabIndex        =   39
         Top             =   1605
         Width           =   1065
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption2 Color"
         Height          =   195
         Left            =   2625
         TabIndex        =   38
         Top             =   1860
         Width           =   1065
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Icon Color"
         Height          =   195
         Left            =   4380
         TabIndex        =   37
         Top             =   1890
         Width           =   735
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header BackColor"
         Height          =   195
         Left            =   4380
         TabIndex        =   36
         Top             =   1350
         Width           =   1275
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BorderWidth"
         Height          =   165
         Left            =   6510
         TabIndex        =   35
         Top             =   1692
         Width           =   975
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CornerCurve"
         Height          =   195
         Left            =   6510
         TabIndex        =   34
         Top             =   1026
         Width           =   930
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ItemSize"
         Height          =   195
         Left            =   6510
         TabIndex        =   33
         Top             =   693
         Width           =   615
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BoxColor"
         Height          =   195
         Left            =   2625
         TabIndex        =   32
         Top             =   1335
         Width           =   645
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opacity"
         Height          =   195
         Left            =   6510
         TabIndex        =   31
         Top             =   1359
         Width           =   555
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header ForeColor"
         Height          =   195
         Left            =   4380
         TabIndex        =   30
         Top             =   1620
         Width           =   1275
      End
      Begin VB.Label lblLineSide 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line Side default"
         Height          =   195
         Left            =   7725
         TabIndex        =   29
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label lblLabel1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BadgeType"
         Height          =   195
         Left            =   7785
         TabIndex        =   28
         Top             =   1305
         Width           =   810
      End
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   1005
      Top             =   8625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin AxConectionGrid.AxConGrid axConGrid 
      Height          =   5580
      Left            =   150
      TabIndex        =   0
      Top             =   360
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   9843
      Enabled         =   -1  'True
      Clickable       =   -1  'True
      Editable        =   -1  'True
      BorderColor     =   11566073
      BorderColorActive=   255
      BackColor       =   16777215
      BorderWidth     =   1
      BoxColor        =   255
      TextColor       =   16777215
      SubTextColor    =   16777215
      HeaderBackColor =   0
      HeaderForeColor =   65535
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontOpacity     =   100
      TextAlignV      =   1
      TextAlignH      =   1
      SubTextAlignV   =   1
      SubTextAlignH   =   1
      IconForeColor   =   16777215
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IcoFont"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageType       =   0
      BadgeVisible    =   1
      SubTextVisible  =   1
      ItemHeight      =   40
      ColWidth        =   150
      ColsMoveable    =   0   'False
      BackColorActive =   255
      BadgeType       =   0
      CornerCurve     =   10
      BoxOpacity      =   90
      LineSideDefault =   2
      LineWidth       =   5
      LineStyle       =   2
      LineOpacity     =   100
      LineStartCap    =   18
      LineEndCap      =   20
      LineStartColor  =   16384
      LineEndColor    =   49344
   End
   Begin VB.Label lblColumn 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Columna:0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   180
      TabIndex        =   2
      Top             =   30
      Width           =   2550
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item Subtext"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2820
      TabIndex        =   1
      Top             =   30
      Width           =   7560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lCol As Long, Item As Long



Private Sub axConGrid_Click(lColumn As Long, lItem As Long, iColumnTo As Long, iLineTo As Long, iCurveTo As Long)
If iLineTo = -1 Then
  lblInfo.Caption = "El Item:" & lItem & " de la Columna:" & lColumn & " no tiene Linea"
Else
  lblInfo.Caption = "Linea desde Item:" & lItem & " de Columna " & lColumn & " a Item:" & iLineTo & " de Columna " & iColumnTo
End If

If iCurveTo = -1 Then
  lblInfo.Caption = lblInfo.Caption & " y no tiene Curva"
Else
  lblInfo.Caption = lblInfo.Caption & " y tiene curva al item:" & iCurveTo
End If

lblColumn.Caption = "Columna: " & lColumn & "  Item: " & lItem

cmdRemoveItem.Caption = "Remove Item " & lItem & vbCrLf & _
                   "Columna " & lColumn
                   
cmdAddItem.Caption = "Add Item to" & vbCrLf & _
                   "Columna " & lColumn

lCol = lColumn
Item = lItem
End Sub

Private Sub chkEditable_Click()
axConGrid.Editable = chkEditable.Value
End Sub

Private Sub chkLabelVisible_Click()
axConGrid.BadgeVisible = chkLabelVisible.Value
End Sub


Private Sub chkClickable_Click()
axConGrid.Clickable = chkClickable.Value
End Sub

Private Sub chkColsMoveable_Click()
axConGrid.ColsMoveable = chkColsMoveable.Value
End Sub

Private Sub cmdClearGrid_Click()
axConGrid.Clear
End Sub

Private Sub cmdReLoad_Click()
axConGrid.Clear
Form_Load
End Sub

Private Sub cmdRemoveItem_Click()

Call axConGrid.RemoveItem(lCol, Item)

End Sub

Private Sub cmdAddItem_Click()

axConGrid.AddItem lCol, "NewItem_X", "--------", "000", , , , , , , True

End Sub

Private Sub Form_Load()
Dim I As Long


With axConGrid
  
  .Redraw = False

  .AddColumn "Column1", 180, 0, 0, vbBlack, vbRed, True
  .AddColumn "Column2", 180, 350, 0, vbRed, vbBlack, True
  .AddColumn "Column3", 150, 650, 100, vbGreen, vbBlue, True
  
  For I = 0 To 6
      .AddItem 0, "Itemtext_" & I, "SubText_" & I, "Item " & I, , , , , "eec" & IIf(I <= 9, I, I - 9), , True
  Next I
  
  For I = 0 To 9
      .AddItem 1, "Item0_" & I, "SubText_" & I, "Item " & I, , , , , "eec" & IIf(I <= 9, I, I - 9), , True
  Next I
    
  For I = 0 To 4
      .AddItem 2, "Itemtext_" & I, "SubText_" & I, "Item " & I, , , , , "eec" & IIf(I <= 9, I, I - 9), , True
  Next I
    
  'Para agregar Líneas por código descomenta las siguientes lineas...
  ' ADDLINE(itemLeft, itemRight, LineStyle, StartAnchor, EndAnchor, StartColor, EndColor, LineWidth, Opacity, Visible)
  .AddLine 0, 1, 1, 3, DashStyleDot, LineCapSquareAnchor, LineCapArrowAnchor, vbBlack, vbRed, 5, 100, True
  .AddLine 0, 0, 1, 0, DashStyleDot, LineCapArrowAnchor, LineCapArrowAnchor, vbBlack, vbRed, 5, 100, True
  .AddLine 1, 2, 0, 5, DashStyleDot, LineCapRoundAnchor, LineCapArrowAnchor, vbBlack, vbRed, 5, 100, True
  .AddLine 2, 2, 1, 2, DashStyleDot, LineCapRoundAnchor, LineCapArrowAnchor, vbBlack, vbRed, 5, 50, True
  .AddLine 2, 1, 1, 1, DashStyleDot, LineCapRoundAnchor, LineCapArrowAnchor, vbGreen, vbBlue, 5, 50, True

  List1.AddItem "toLeft"
  List1.AddItem "toRight"
  List1.AddItem "Auto"
  
  List2.AddItem "eLeft"
  List2.AddItem "eCenter"
  List2.AddItem "eRight"
  
  List3.AddItem "eTop"
  List3.AddItem "eMiddle"
  List3.AddItem "eBottom"
          
  List4.AddItem "vbNone"
  List4.AddItem "vbLabel"
  List4.AddItem "vbIcon"
  
  List5.AddItem "eTop"
  List5.AddItem "eMiddle"
  List5.AddItem "eBottom"
  
  List6.AddItem "eLeft"
  List6.AddItem "eCenter"
  List6.AddItem "eRight"
  
  List7.AddItem "scNone"
  List7.AddItem "scAllways"
  List7.AddItem "scOnlyActive"
  List7.AddItem "scAllPastPoint"

  .BorderWidth = txtBorderWidth.Text
  .ColWidth = txtColWidth.Text
  .Clickable = chkClickable.Value
  .ItemHeight = txtItemSize.Text
  .BoxOpacity = txtOpacity.Text
  .TextAlignH = eLeft
  .SubTextAlignH = eLeft
  .SubTextAlignV = eTop
  .BadgeVisible = True
  .BackColor = pColor(0).BackColor
  .BorderColor = pColor(1).BackColor
  .BorderColorActive = pColor(2).BackColor
  .BoxColor = pColor(7).BackColor
  
  .Redraw = True
End With

 
End Sub

Private Sub Form_Resize()
lblColumn.Move 50, 50
lblInfo.Move 100 + lblColumn.Width, 50
Frame1.Move 50, Me.ScaleHeight - (Frame1.Height)
axConGrid.Move 50, lblColumn.Height + 100, Me.ScaleWidth - 100, Me.ScaleHeight - (Frame1.Height + lblColumn.Height + 100)

End Sub

Private Sub List1_Click()
axConGrid.LineSideDefault = List1.ListIndex
End Sub

Private Sub List2_Click()
axConGrid.TextAlignH = List2.ListIndex
End Sub

Private Sub List3_Click()
axConGrid.TextAlignV = List3.ListIndex
End Sub


Private Sub List4_Click()
axConGrid.BadgeType = List4.ListIndex
End Sub

Private Sub List5_Click()
axConGrid.SubTextAlignV = List5.ListIndex
End Sub

Private Sub List6_Click()
axConGrid.SubTextAlignH = List6.ListIndex
End Sub

Private Sub List7_Click()
axConGrid.SubTextVisible = List7.ListIndex
End Sub

Private Sub pColor_Click(Index As Integer)
With cDialog
  .ShowColor
  pColor(Index).BackColor = .Color
End With
With axConGrid
  .BackColor = pColor(0).BackColor
  .BorderColor = pColor(1).BackColor
  .BorderColorActive = pColor(2).BackColor
  .HeaderBackColor = pColor(6).BackColor
  .HeaderForeColor = pColor(8).BackColor
  .BoxColor = pColor(7).BackColor
End With
End Sub

Private Sub txtLineWidth_Change()
On Error Resume Next
axConGrid.LineWidth = txtLineWidth.Text
End Sub

Private Sub txtBorderWidth_Change()
On Error Resume Next
axConGrid.BorderWidth = txtBorderWidth.Text
End Sub

Private Sub txtCurve_Change()
On Error Resume Next
axConGrid.CornerCurve = txtCurve.Text
End Sub

Private Sub txtOpacity_Change()
On Error Resume Next
axConGrid.BoxOpacity = txtOpacity.Text
End Sub

Private Sub txtColWidth_Change()
On Error Resume Next
axConGrid.ColWidth = txtColWidth.Text
End Sub

Private Sub txtItemSize_Change()
On Error Resume Next
axConGrid.ItemHeight = txtItemSize.Text
End Sub

