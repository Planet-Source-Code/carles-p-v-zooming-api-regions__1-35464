VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' cRegion.cls
' Carles P.V. - 2002
'================================================

Option Explicit

'Private Type XFORM
'    eM11 As Double
'    eM12 As Double
'    eM21 As Double
'    eM22 As Double
'    eDx  As Double
'    eDy  As Double
'End Type
'
'Private Type RGNDATAHEADER
'    dwSize   As Long
'    iType    As Long
'    nCount   As Long
'    nRgnSize As Long
'    rcBound  As RECT2 '-> Region bound rectangle
'End Type
'
'Private Type RGNDATA
'    rdh      As RGNDATAHEADER
'    Buffer() As Byte '-> Rectangles coord. (Left-Top -> Right-Bottom)
'End Type

'Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As XFORM, ByVal nCount As Long, lpRgnData As RGNDATA) As Long
'Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As RGNDATA) As Long

Public Enum eCombineRgnMode
    [RGN_Add] = 2      '= RGN_OR
    [RGN_Subtract] = 4 '= RGN_DIFF
End Enum

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const ALTERNATE As Long = 1
Private Const WINDING   As Long = 2
                         
Private Const RGN_AND   As Long = 1
Private Const RGN_OR    As Long = 2
Private Const RGN_XOR   As Long = 3
Private Const RGN_DIFF  As Long = 4
Private Const RGN_COPY  As Long = 5
Private Const RGN_MAX   As Long = RGN_COPY
Private Const RGN_MIN   As Long = RGN_AND

Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal dx As Long, ByVal dy As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT2) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
       
Private m_hRegionMain      As Long  'Main region
Private m_hRegion          As Long  'Original region
Private m_hRegionScaled    As Long  'Scaled region

Private m_ZoomFactor       As Long  'Scale factor (>= 1)

Private m_BuffData()       As Long  'Original region data (1:1)
Private m_BuffDataScaled() As Long  'Scaled region data   (ZoomFactor:1)

Private m_hBrush(7)        As Long  'Countour region...
Private m_nBrush           As Long  'Current brush



Private Sub Class_Initialize()

  Dim lId1           As Long
  Dim lId2           As Long
  Dim hBitmap        As Long
  Dim tBytes(1 To 8) As Integer
  Dim bBytes(1 To 8) As Integer

    tBytes(1) = &HF0
    tBytes(2) = &H78
    tBytes(3) = &H3C
    tBytes(4) = &H1E
    tBytes(5) = &HF
    tBytes(6) = &H87
    tBytes(7) = &HC3
    tBytes(8) = &HE1

    hBitmap = CreateBitmap(8, 8, 1, 1, tBytes(1))
    m_hBrush(0) = CreatePatternBrush(hBitmap)
    Call DeleteObject(hBitmap)

    For lId1 = 7 To 1 Step -1
        For lId2 = 1 To 8
            If (lId1 + lId2 < 9) Then
                bBytes(lId2) = tBytes(lId1 + lId2)
              Else
                bBytes(lId2) = tBytes(lId1 + lId2 - 8)
            End If
        Next lId2
        hBitmap = CreateBitmap(8, 8, 1, 1, bBytes(1))
        m_hBrush(lId1) = CreatePatternBrush(hBitmap)
        Call DeleteObject(hBitmap)
    Next lId1

    Erase tBytes
    Erase bBytes
    
    m_ZoomFactor = 1
End Sub

Private Sub Class_Terminate()

  Dim lIdx As Long
  
    For lIdx = 0 To 7
        If (m_hBrush(lIdx) <> 0) Then
            Call DeleteObject(m_hBrush(lIdx))
            m_hBrush(lIdx) = 0
        End If
    Next lIdx
    Call Clear
End Sub

'//

Public Sub InitRegion(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    
    '-- Invert/Crop
    Call Clear
    m_hRegionMain = CreateRectRgn(x1, y1, x2, y2)
End Sub


Public Sub RectRegion(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Mode As eCombineRgnMode, Optional ByVal Crop As Boolean = 0)
    
  Dim hTmpRgn As Long

    If (m_hRegion <> 0) Then
        hTmpRgn = CreateRectRgn(x1, y1, x2, y2)
        Call CombineRgn(m_hRegion, m_hRegion, hTmpRgn, Mode)
      Else
        If (Mode <> RGN_Subtract) Then
            m_hRegion = CreateRectRgn(x1, y1, x2, y2)
        End If
    End If
    If (hTmpRgn <> 0) Then
        Call DeleteObject(hTmpRgn)
    End If
    If (Crop) Then
        Call CombineRgn(m_hRegion, m_hRegion, m_hRegionMain, RGN_AND)
    End If
    
    Call pvGetRgnData
    Call pvScaleRgn(m_ZoomFactor)
End Sub

Public Sub EllipseRegion(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Mode As eCombineRgnMode, Optional ByVal Crop As Boolean = 0)

  Dim hTmpRgn As Long
    
    If (m_hRegion <> 0) Then
        hTmpRgn = CreateEllipticRgn(x1, y1, x2, y2)
        Call CombineRgn(m_hRegion, m_hRegion, hTmpRgn, Mode)
      Else
        If (Mode <> RGN_Subtract) Then
            m_hRegion = CreateEllipticRgn(x1, y1, x2, y2)
        End If
    End If
    If (hTmpRgn <> 0) Then
        Call DeleteObject(hTmpRgn)
    End If
    If (Crop) Then
        Call CombineRgn(m_hRegion, m_hRegion, m_hRegionMain, RGN_AND)
    End If
    
    Call pvGetRgnData
    Call pvScaleRgn(m_ZoomFactor)
End Sub

Public Sub PolyRegion(lPoints() As Long, ByVal Mode As eCombineRgnMode, Optional ByVal Crop As Boolean = 0)

  Dim hTmpRgn As Long

    If (UBound(lPoints) > 1) Then
    
        If (m_hRegion <> 0) Then
            hTmpRgn = CreatePolygonRgn(lPoints(0), (UBound(lPoints) + 1) \ 2, WINDING)
            Call CombineRgn(m_hRegion, m_hRegion, hTmpRgn, Mode)
          Else
            If (Mode <> RGN_Subtract) Then
                m_hRegion = CreatePolygonRgn(lPoints(0), (UBound(lPoints) + 1) \ 2, WINDING)
            End If
        End If
        If (hTmpRgn <> 0) Then
            Call DeleteObject(hTmpRgn)
        End If
        If (Crop) Then
            Call CombineRgn(m_hRegion, m_hRegion, m_hRegionMain, RGN_AND)
        End If
    
        Call pvGetRgnData
        Call pvScaleRgn(m_ZoomFactor)
    End If
End Sub

Public Sub InvertRegion()

  Dim rgnRct1 As RECT2
  Dim rgnRct2 As RECT2

    Call GetRgnBox(m_hRegion, rgnRct1)
    Call GetRgnBox(m_hRegionMain, rgnRct2)
    
    If (IsRectEmpty(rgnRct1)) Then
        With rgnRct2
            m_hRegion = CreateRectRgn(.x1, .y1, .x2, .y2)
        End With
      Else
        Call CombineRgn(m_hRegion, m_hRegionMain, m_hRegion, RGN_XOR)
    End If
    
    Call pvGetRgnData
    Call pvScaleRgn(m_ZoomFactor)
End Sub

Public Sub Clear()

    If (m_hRegionMain <> 0) Then
        Call DeleteObject(m_hRegionMain)
        m_hRegionMain = 0
    End If
    If (m_hRegion <> 0) Then
        Call DeleteObject(m_hRegion)
        m_hRegion = 0
    End If
    If (m_hRegionScaled <> 0) Then
        Call DeleteObject(m_hRegionScaled)
        m_hRegionScaled = 0
    End If
    
    Erase m_BuffData
    Erase m_BuffDataScaled
End Sub

Public Sub DrawRegion(ByVal hDC As Long, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0)
    
    Call OffsetRgn(m_hRegionScaled, xOffset, yOffset)
    Call FrameRgn(hDC, m_hRegionScaled, m_hBrush(m_nBrush), 1, 1)
    Call OffsetRgn(m_hRegionScaled, -xOffset, -yOffset)
End Sub

Public Sub RotateBrush()
    If (m_nBrush < 7) Then
        m_nBrush = m_nBrush + 1
      Else
        m_nBrush = 0
    End If
End Sub

'//

Public Property Let ZoomFactor(ByVal Factor As Long)
    If (Factor < 1) Then
        m_ZoomFactor = 1
      Else
        m_ZoomFactor = Factor
    End If
    Call pvScaleRgn(m_ZoomFactor)
End Property

Public Property Get ZoomFactor() As Long
    ZoomFactor = m_ZoomFactor
End Property

Public Property Get Region() As Long
    Region = m_hRegionScaled
End Property

Public Property Get RegionMain() As Long
    RegionMain = m_hRegionMain
End Property

Public Function GetBoundRect(x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Boolean

  Dim lType  As Long
  Dim rgnRct As RECT2

   'lType = GetRgnBox(m_hRegion, rgnRct)

    If (IsRectEmpty(rgnRct)) Then '~ lType = NULLREGION ?
        GetBoundRect = False
      Else
        With rgnRct
            x1 = .x1 '= m_BuffData(4)
            y1 = .y1 '= m_BuffData(5)
            x2 = .x2 '= m_BuffData(6)
            y2 = .y2 '= m_BuffData(7)
            GetBoundRect = True
        End With
    End If
End Function

'//

Private Sub pvGetRgnData()

  Dim lCount As Long

    If (m_hRegion <> 0) Then
        lCount = GetRegionData(m_hRegion, 0, ByVal 0)
        ReDim m_BuffData(lCount \ 4 - 1)
        GetRegionData m_hRegion, lCount, m_BuffData(0)
    End If
End Sub

Private Sub pvScaleRgn(ByVal Factor As Long)

  Dim lIdx As Long
  
    If (m_hRegion <> 0) Then
        
        '-- Get data
        m_BuffDataScaled() = m_BuffData()
        
        '-- Scale
        If (Factor > 1) Then
            '0 - 3   : Header
            '4 - 7   : rcBound
            '8 - End : Region rects. collection
            For lIdx = 4 To UBound(m_BuffDataScaled)
                m_BuffDataScaled(lIdx) = m_BuffData(lIdx) * Factor
            Next lIdx
        End If
        
        '-- Delete previous
        If (m_hRegionScaled <> 0) Then
            Call DeleteObject(m_hRegionScaled)
            m_hRegionScaled = 0
        End If
        
        '-- Create scaled region from data
        m_hRegionScaled = ExtCreateRegion(ByVal 0, 4 * (UBound(m_BuffDataScaled) + 1), m_BuffDataScaled(0))
    End If
End Sub
