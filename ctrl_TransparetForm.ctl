VERSION 5.00
Begin VB.UserControl ctrl_TransparetForm 
   BackStyle       =   0  'Transparent
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   ScaleHeight     =   510
   ScaleWidth      =   510
End
Attribute VB_Name = "ctrl_TransparetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type
    
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Const WINDING = 2
Const PS_SOLID = 0

Public Sub ShapeForm(m_Form As Form, m_SkinPath As String, m_Active As Boolean)
    Dim v_iLoop As Integer
    Dim v_iPtsCount As Integer
    Dim v_lStatus As Long
    Dim a_tPts() As POINTAPI
    Dim v_lRgn As Long
    Dim v_lOld_rgn As Long
    Dim v_lPen As Long
    Dim v_lOld_pen As Long
    Dim v_lRtn As Long
    Dim v_sString As String
    Dim v_sTemp As String
    
    If m_Active = True Then
    
    v_sString = Space(255)
    v_lRtn = GetPrivateProfileString("TRANAPARENY", "PtsCount", "", v_sString, Len(v_sString), m_SkinPath & "\Transparency.ini")
    v_iPtsCount = Val(TrimString(v_sString))
    
    ReDim Preserve a_tPts(v_iPtsCount)
    For v_iLoop = 1 To v_iPtsCount
        v_sString = Space(255)
        v_lRtn = GetPrivateProfileString("TRANAPARENY", "X(" & v_iLoop & ")", "", v_sString, Len(v_sString), m_SkinPath & "\Transparency.ini")
        a_tPts(v_iLoop).x = Val(TrimString(v_sString))
        
        v_sString = Space(255)
        v_lRtn = GetPrivateProfileString("TRANAPARENY", "Y(" & v_iLoop & ")", "", v_sString, Len(v_sString), m_SkinPath & "\Transparency.ini")
        a_tPts(v_iLoop).y = Val(TrimString(v_sString))
    Next v_iLoop

    ' Set the form region.
    v_lRgn = CreatePolygonRgn(a_tPts(1), v_iPtsCount, WINDING)
    v_lOld_rgn = SetWindowRgn(m_Form.hWnd, v_lRgn, True)

    ' Create a pen to draw the region edge.
    v_lPen = CreatePen(PS_SOLID, 2, vbBlack)
    v_lOld_pen = SelectObject(m_Form.hdc, v_lPen)
    
    v_lStatus = Polygon(m_Form.hdc, a_tPts(1), v_iPtsCount)
    
    v_lPen = SelectObject(m_Form.hdc, v_lOld_pen)
    v_lStatus = DeleteObject(v_lPen)
    
    Else
    
    ReDim Preserve a_tPts(4)
    a_tPts(1).x = 0
    a_tPts(1).y = 0
    a_tPts(2).x = m_Form.Width / Screen.TwipsPerPixelX
    a_tPts(2).y = 0
    a_tPts(3).x = m_Form.Width / Screen.TwipsPerPixelX
    a_tPts(3).y = m_Form.Height / Screen.TwipsPerPixelY
    a_tPts(4).x = 0
    a_tPts(4).y = m_Form.Height / Screen.TwipsPerPixelY
    
    ' Set the form region.
    v_lRgn = CreatePolygonRgn(a_tPts(1), v_iPtsCount, WINDING)
    v_lOld_rgn = SetWindowRgn(m_Form.hWnd, v_lRgn, True)

    ' Create a pen to draw the region edge.
    v_lPen = CreatePen(PS_SOLID, 2, vbBlack)
    v_lOld_pen = SelectObject(m_Form.hdc, v_lPen)
    
    v_lStatus = Polygon(m_Form.hdc, a_tPts(1), v_iPtsCount)
    
    v_lPen = SelectObject(m_Form.hdc, v_lOld_pen)
    v_lStatus = DeleteObject(v_lPen)
    
    End If
End Sub

Public Function TrimString(m_Str As String) As String
    m_Str = RTrim$(m_Str)
    m_Str = Left(m_Str, Len(m_Str) - 1)
    TrimString = m_Str
End Function

