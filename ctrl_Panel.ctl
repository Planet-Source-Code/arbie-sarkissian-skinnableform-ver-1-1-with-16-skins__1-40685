VERSION 5.00
Begin VB.UserControl ctrl_Panel 
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   ScaleHeight     =   1695
   ScaleWidth      =   1980
   Begin VB.PictureBox pic_RightBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   180
      TabIndex        =   4
      Top             =   1200
      Width           =   180
   End
   Begin VB.PictureBox pic_LeftBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   180
      TabIndex        =   3
      Top             =   1200
      Width           =   180
   End
   Begin VB.PictureBox pic_DownBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox pic_UpBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.PictureBox pic_Source 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "ctrl_Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Const DefBackColor = 0

Dim v_oBackColor As OLE_COLOR
Dim v_sSkinPath As String

Public Sub DrawPanel()
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    With UserControl
        .pic_Source.Picture = LoadPicture(SkinPath & "\img_Panel.bmp")
        .pic_UpBorder.Width = .Width
        .pic_UpBorder.Height = 150
        
        .pic_UpBorder.Cls
        v_lRtn = BitBlt(.pic_UpBorder.hDC, 0, 0, 20, 10, .pic_Source.hDC, 0, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 20)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_UpBorder.hDC, v_iLoop * 20, 0, 20, 10, .pic_Source.hDC, 20, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_UpBorder.hDC, (.Width / Screen.TwipsPerPixelX) - 20, 0, 20, 10, .pic_Source.hDC, 53, 0, SRCCOPY)
        pic_UpBorder.Refresh
    
        .pic_LeftBorder.Cls
        .pic_LeftBorder.Width = 150
        .pic_LeftBorder.Height = .Height - 300
        .pic_LeftBorder.Top = 150
        .pic_RightBorder.Cls
        .pic_RightBorder.Width = 150
        .pic_RightBorder.Height = .Height - 300
        .pic_RightBorder.Top = 150
        .pic_RightBorder.Left = .Width - 150
        v_iCenterImgFrequency = Abs((.Height / Screen.TwipsPerPixelY) / 10)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 0 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_LeftBorder.hDC, 0, v_iLoop * 10, 20, 10, .pic_Source.hDC, 0, 10, SRCCOPY)
                v_lRtn = BitBlt(.pic_RightBorder.hDC, 0, v_iLoop * 10, 20, 10, .pic_Source.hDC, 63, 10, SRCCOPY)
            Next v_iLoop
        End If
        .pic_LeftBorder.Refresh
        .pic_RightBorder.Refresh
        
        .pic_DownBorder.Cls
        .pic_DownBorder.Width = .Width
        .pic_DownBorder.Height = 150
        .pic_DownBorder.Top = .Height - .pic_DownBorder.Height
        v_lRtn = BitBlt(.pic_DownBorder.hDC, 0, 0, 20, 10, .pic_Source.hDC, 0, 65, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 20)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_DownBorder.hDC, v_iLoop * 20, 0, 20, 10, .pic_Source.hDC, 20, 65, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_DownBorder.hDC, (.Width / Screen.TwipsPerPixelX) - 20, 0, 20, 10, .pic_Source.hDC, 53, 65, SRCCOPY)
        .pic_DownBorder.Refresh
    End With
End Sub

Public Sub Refresh()
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    With UserControl
        .pic_UpBorder.Width = .Width
        .pic_UpBorder.Height = 150
        
        v_lRtn = BitBlt(.pic_UpBorder.hDC, 0, 0, 10, 10, .pic_Source.hDC, 0, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 10)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_UpBorder.hDC, v_iLoop * 10, 0, 10, 10, .pic_Source.hDC, 10, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_UpBorder.hDC, (.Width / Screen.TwipsPerPixelX) - 10, 0, 10, 10, .pic_Source.hDC, 63, 0, SRCCOPY)
    
        .pic_LeftBorder.Width = 150
        .pic_LeftBorder.Height = .Height - 300
        .pic_LeftBorder.Top = 150
        .pic_RightBorder.Width = 150
        .pic_RightBorder.Height = .Height - 300
        .pic_RightBorder.Top = 150
        .pic_RightBorder.Left = .Width - 150
        v_iCenterImgFrequency = Abs((.Height / Screen.TwipsPerPixelY) / 10)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 0 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_LeftBorder.hDC, 0, v_iLoop * 10, 10, 10, .pic_Source.hDC, 0, 10, SRCCOPY)
                v_lRtn = BitBlt(.pic_RightBorder.hDC, 0, v_iLoop * 10, 10, 10, .pic_Source.hDC, 63, 10, SRCCOPY)
            Next v_iLoop
        End If
        
        .pic_DownBorder.Width = .Width
        .pic_DownBorder.Height = 150
        .pic_DownBorder.Top = .Height - .pic_DownBorder.Height
        v_lRtn = BitBlt(.pic_DownBorder.hDC, 0, 0, 10, 10, .pic_Source.hDC, 0, 65, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 10)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_DownBorder.hDC, v_iLoop * 10, 0, 10, 10, .pic_Source.hDC, 10, 65, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_DownBorder.hDC, (.Width / Screen.TwipsPerPixelX) - 10, 0, 10, 10, .pic_Source.hDC, 63, 65, SRCCOPY)
    End With
End Sub

Public Property Get SkinPath() As String
    SkinPath = v_sSkinPath
End Property

Public Property Let SkinPath(ByVal m_SkinPath As String)
    v_sSkinPath = m_SkinPath
    PropertyChanged "SkinPath"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = v_oBackColor
End Property

Public Property Let BackColor(ByVal m_BackColor As OLE_COLOR)
    v_oBackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

Private Sub UserControl_InitProperties()
    v_sSkinPath = App.Path & "\Skins\Titanium"
    v_oBackColor = DefBackColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    v_sSkinPath = PropBag.ReadProperty("SkinPath", App.Path & "\Skins\Titanium")
    Call DrawPanel
    
    v_oBackColor = PropBag.ReadProperty("BackColor", DefBackColor)
    UserControl.BackColor = v_oBackColor
End Sub

Private Sub UserControl_Resize()
    Call Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SkinPath", v_sSkinPath, App.Path & "\Skins\Titanium")
    Call PropBag.WriteProperty("BackColor", v_oBackColor, DefBackColor)
 End Sub
