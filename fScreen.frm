VERSION 5.00
Begin VB.Form fScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zooming API regions"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   ClipControls    =   0   'False
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   523
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrSelection 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill &preview"
      Height          =   780
      Left            =   6480
      TabIndex        =   11
      Top             =   6090
      Width           =   1200
   End
   Begin VB.ComboBox cbZoom 
      Height          =   315
      ItemData        =   "fScreen.frx":0000
      Left            =   690
      List            =   "fScreen.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   6540
      Width           =   1155
   End
   Begin VB.CommandButton cmdInvert 
      Caption         =   "&Invert"
      Height          =   375
      Left            =   4830
      TabIndex        =   8
      Top             =   6495
      Width           =   1200
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4830
      TabIndex        =   7
      Top             =   6090
      Width           =   1200
   End
   Begin VB.OptionButton optRegion 
      Caption         =   "&Polygon"
      Height          =   210
      Index           =   2
      Left            =   3150
      TabIndex        =   6
      ToolTipText     =   "DblClick to close"
      Top             =   6615
      Width           =   1230
   End
   Begin VB.OptionButton optRegion 
      Caption         =   "&Ellipse"
      Height          =   210
      Index           =   1
      Left            =   3150
      TabIndex        =   5
      Top             =   6345
      Width           =   1230
   End
   Begin VB.OptionButton optRegion 
      Caption         =   "&Rectangle"
      Height          =   210
      Index           =   0
      Left            =   3150
      TabIndex        =   4
      Top             =   6090
      Value           =   -1  'True
      Width           =   1230
   End
   Begin VB.ComboBox cbMode 
      Height          =   315
      ItemData        =   "fScreen.frx":006D
      Left            =   690
      List            =   "fScreen.frx":0077
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6015
      Width           =   1155
   End
   Begin VB.PictureBox iScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   150
      MouseIcon       =   "fScreen.frx":008A
      MousePointer    =   99  'Custom
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   150
      Width           =   7530
      Begin VB.Line shpLine 
         BorderColor     =   &H00008000&
         Index           =   0
         Visible         =   0   'False
         X1              =   158
         X2              =   306
         Y1              =   175
         Y2              =   236
      End
      Begin VB.Shape shpEllp 
         BorderColor     =   &H00008000&
         Height          =   1590
         Left            =   1095
         Shape           =   2  'Oval
         Top             =   1500
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.Shape shpRect 
         BorderColor     =   &H00008000&
         Height          =   2055
         Left            =   2100
         Top             =   2010
         Visible         =   0   'False
         Width           =   2745
      End
   End
   Begin VB.Label lblZoom 
      Caption         =   "Zoom"
      Height          =   315
      Left            =   180
      TabIndex        =   10
      Top             =   6600
      Width           =   570
   End
   Begin VB.Label lblRegion 
      Caption         =   "Region type"
      Height          =   240
      Left            =   2115
      TabIndex        =   3
      Top             =   6075
      Width           =   990
   End
   Begin VB.Label lblMode 
      Caption         =   "Mode"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   6075
      Width           =   735
   End
End
Attribute VB_Name = "fScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private oRegion         As cRegion
Private m_lFactor       As Long

Private m_uStartPt      As POINTAPI
Private m_lPolyCoords() As Long
Private m_bDblClicked   As Boolean



Private Sub Form_Load()

    Set oRegion = New cRegion
    oRegion.InitRegion 0, 0, iScreen.ScaleWidth, iScreen.ScaleHeight
    
    cbMode.ListIndex = 0
    cbZoom.ListIndex = 0
    
    m_lFactor = 1
    ReDim m_lPolyCoords(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oRegion = Nothing
    Set fScreen = Nothing
End Sub

Private Sub RefreshView()
    Call oRegion.DrawRegion(iScreen.hDC)
    Call iScreen.Refresh
End Sub

'//

Private Sub iScreen_DblClick()
    m_bDblClicked = True
    tmrSelection.Enabled = True
End Sub

Private Sub iScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (Button = vbLeftButton) Then
        tmrSelection.Enabled = False
        Call iScreen.Cls
        Call RefreshView
        If (optRegion(2).Value = False) Then
            m_uStartPt.x = x
            m_uStartPt.y = y
        End If
    End If
End Sub

Private Sub iScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim x1 As Long, y1 As Long
  Dim x2 As Long, y2 As Long
  Dim lSwap As Long
    
    '-- Readjust to "real" pixel
    x1 = (m_uStartPt.x \ m_lFactor) * m_lFactor
    y1 = (m_uStartPt.y \ m_lFactor) * m_lFactor
    x2 = (x \ m_lFactor) * m_lFactor
    y2 = (y \ m_lFactor) * m_lFactor
    
    '-- In case of Rect. or Ellip. region, invert shape coord.
    If (optRegion(2).Value = False) Then
        If (x2 < x1) Then lSwap = x1: x1 = x2: x2 = lSwap
        If (y2 < y1) Then lSwap = y1: y1 = y2: y2 = lSwap
    End If
    
    If (Button = vbLeftButton Or optRegion(2).Value = True) Then
    
        Select Case True
            Case optRegion(0) '-- Rectangle
                On Error Resume Next
                shpRect.Visible = -1
                shpRect.Move x1, y1, x2 - x1, y2 - y1
                
            Case optRegion(1) '-- Ellipse
                On Error Resume Next
                shpEllp.Visible = -1
                shpEllp.Move x1, y1, x2 - x1, y2 - y1
                                                                                   
            Case optRegion(2) '-- Polygon
                If (UBound(m_lPolyCoords) > 1) Then
                    shpLine(shpLine.Count - 1).Visible = True
                    With shpLine(shpLine.Count - 1)
                        .x1 = x1: .x2 = x2: .y1 = y1: .y2 = y2
                    End With
                  Else
                    m_lPolyCoords(0) = x1
                    m_lPolyCoords(1) = y1
                End If
        End Select
    End If
    On Error GoTo 0
End Sub

Private Sub iScreen_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim ellIncSx As Long, ellIncSy As Long
  Dim ellIncEx As Long, ellIncEy As Long
  Dim lIdx As Long

    If (Button = vbLeftButton) Then

        '-- Adjust to zoom factor
        x = x \ m_lFactor
        y = y \ m_lFactor
        m_uStartPt.x = m_uStartPt.x \ m_lFactor
        m_uStartPt.y = m_uStartPt.y \ m_lFactor
    
        '-- Ellipse adjust
        If (optRegion(1)) Then
            If (x < m_uStartPt.x) Then ellIncSx = 1: ellIncEx = 0 Else ellIncSx = 0: ellIncEx = 1
            If (y < m_uStartPt.y) Then ellIncSy = 1: ellIncEy = 0 Else ellIncSy = 0: ellIncEy = 1
        End If
    
        Select Case True
            Case optRegion(0) '-- Rectangle
                oRegion.RectRegion m_uStartPt.x, m_uStartPt.y, x, y, Choose(cbMode.ListIndex + 1, RGN_Add, RGN_Subtract)
                shpRect.Visible = 0
        
            Case optRegion(1) '-- Ellipse
                oRegion.EllipseRegion m_uStartPt.x + ellIncSx, m_uStartPt.y + ellIncSy, x + ellIncEx, y + ellIncEy, Choose(cbMode.ListIndex + 1, RGN_Add, RGN_Subtract)
                shpEllp.Visible = 0
        
            Case optRegion(2) '-- Polygon
                m_lPolyCoords(UBound(m_lPolyCoords) - 1) = x
                m_lPolyCoords(UBound(m_lPolyCoords) - 0) = y
            
            If (m_bDblClicked) Then
            
                m_bDblClicked = 0
                For lIdx = 1 To shpLine.Count - 1
                    Unload shpLine(lIdx)
                Next lIdx
                oRegion.PolyRegion m_lPolyCoords(), Choose(cbMode.ListIndex + 1, RGN_Add, RGN_Subtract)
                
                shpLine(0).Visible = False
                ReDim m_lPolyCoords(1)
                
                Call iScreen.Cls
                Call RefreshView
              Else
                m_uStartPt.x = x * m_lFactor
                m_uStartPt.y = y * m_lFactor
                
                ReDim Preserve m_lPolyCoords(UBound(m_lPolyCoords) + 2)
                VB.Load shpLine(shpLine.Count)
            End If
        End Select
    
        If (optRegion(2).Value = False) Then
            Call iScreen.Cls
            Call RefreshView
            tmrSelection.Enabled = True
        End If
    End If
End Sub

Private Sub iScreen_LostFocus()

  Dim lIdx As Long

    If (optRegion(2).Value = False) Then
        
        For lIdx = 1 To shpLine.Count - 1
            Unload shpLine(lIdx)
        Next lIdx
        shpLine(0).Visible = False
        ReDim m_lPolyCoords(1)
        
        m_bDblClicked = False
    End If
End Sub

'//

Private Sub cbZoom_Click()
    m_lFactor = cbZoom.ListIndex + 1
    oRegion.ZoomFactor = m_lFactor
    Call iScreen.Cls
    Call RefreshView
End Sub

Private Sub cmdClear_Click()
    Call oRegion.Clear
    Call oRegion.InitRegion(0, 0, iScreen.ScaleWidth, iScreen.ScaleHeight)
    Call iScreen.Cls
    Call RefreshView
End Sub

Private Sub cmdInvert_Click()
    Call oRegion.InvertRegion
    Call iScreen.Cls
    Call RefreshView
End Sub

Private Sub cmdFill_Click()

  Dim hBrush As Long

    hBrush = CreateSolidBrush(&HFFFFFF)
    Call FillRgn(iScreen.hDC, oRegion.Region, hBrush)
    Call DeleteObject(hBrush)
    Call RefreshView
End Sub

'//

Private Sub tmrSelection_Timer()
    Call oRegion.RotateBrush
    Call RefreshView
End Sub

