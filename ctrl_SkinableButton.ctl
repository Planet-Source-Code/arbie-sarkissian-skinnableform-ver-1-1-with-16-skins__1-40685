VERSION 5.00
Begin VB.UserControl ctrl_SkinableButton 
   BackStyle       =   0  'Transparent
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1290
   ScaleHeight     =   360
   ScaleWidth      =   1290
   Begin VB.PictureBox pic_Button 
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
      Begin VB.Label lbl_Caption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   570
      End
   End
   Begin VB.PictureBox pic_Buttons 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "ctrl_SkinableButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Const DefCaption = "Caption"
Const DefForeColor = &HFFFFFF
Const DefEnabled = 1

Dim v_sSkinPath As String
Dim v_sCaption As String
Dim v_oForeColor As OLE_COLOR
Dim v_bEnabled As Boolean

Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Public Sub LoadSkin()
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    With UserControl
        .pic_Buttons.Picture = LoadPicture(SkinPath & "\img_Buttons.bmp")
        .pic_Button.Width = .Width
        .pic_Button.Height = 360
        
        .pic_Button.Cls
        v_lRtn = BitBlt(.pic_Button.hDC, 0, 0, 15, 24, .pic_Buttons.hDC, 0, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 15)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_Button.hDC, v_iLoop * 15, 0, 15, 24, .pic_Buttons.hDC, 15, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_Button.hDC, (.Width / Screen.TwipsPerPixelX) - 16, 0, 16, 24, .pic_Buttons.hDC, 55, 0, SRCCOPY)
        pic_Button.Refresh
        
        .lbl_Caption.Width = .Width
        .lbl_Caption.Top = 60
    End With
End Sub

Public Sub Refresh()
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    With UserControl
        .pic_Button.Width = .Width
        .pic_Button.Height = 360
        
        .pic_Button.Cls
        v_lRtn = BitBlt(.pic_Button.hDC, 0, 0, 15, 24, .pic_Buttons.hDC, 0, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 15)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_Button.hDC, v_iLoop * 15, 0, 15, 24, .pic_Buttons.hDC, 15, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_Button.hDC, (.Width / Screen.TwipsPerPixelX) - 16, 0, 16, 24, .pic_Buttons.hDC, 55, 0, SRCCOPY)
        
        .lbl_Caption.Width = .Width
        .lbl_Caption.Top = 60
        .lbl_Caption.ForeColor = ForeColor
    End With
End Sub

Public Property Get SkinPath() As String
    SkinPath = v_sSkinPath
End Property

Public Property Let SkinPath(ByVal m_SkinPath As String)
    v_sSkinPath = m_SkinPath
    PropertyChanged "SkinPath"
End Property

Public Property Get Caption() As String
    Caption = v_sCaption
End Property

Public Property Let Caption(ByVal m_Caption As String)
    v_sCaption = m_Caption
    PropertyChanged "Caption"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = v_oForeColor
End Property

Public Property Let ForeColor(ByVal m_ForeColor As OLE_COLOR)
    v_oForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
    Enabled = v_bEnabled
End Property

Public Property Let Enabled(ByVal m_Enabled As Boolean)
    v_bEnabled = m_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub lbl_Caption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    If Button = 1 Then
    
    With UserControl
        .pic_Button.Cls
        v_lRtn = BitBlt(.pic_Button.hDC, 0, 0, 15, 24, .pic_Buttons.hDC, 144, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 15)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_Button.hDC, v_iLoop * 15, 0, 15, 24, .pic_Buttons.hDC, 159, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_Button.hDC, (.Width / Screen.TwipsPerPixelX) - 16, 0, 16, 24, .pic_Buttons.hDC, 202, 0, SRCCOPY)
        
        .lbl_Caption.Width = .Width
        .lbl_Caption.Top = 75
        .lbl_Caption.ForeColor = ForeColor
    End With
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub lbl_Caption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call pic_Button_MouseMove(Button, Shift, X, Y)
        RaiseEvent MouseUp(Button, Shift, X, Y)
        RaiseEvent Click
    End If
End Sub

Private Sub pic_Button_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer
    
    If Enabled = True Then
    
    With UserControl
        .pic_Button.Cls
        v_lRtn = BitBlt(.pic_Button.hDC, 0, 0, 15, 24, .pic_Buttons.hDC, 72, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 15)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_Button.hDC, v_iLoop * 15, 0, 15, 24, .pic_Buttons.hDC, 83, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_Button.hDC, (.Width / Screen.TwipsPerPixelX) - 16, 0, 16, 24, .pic_Buttons.hDC, 128, 0, SRCCOPY)
        
        .lbl_Caption.Width = .Width
        .lbl_Caption.Top = 75
        .lbl_Caption.ForeColor = ForeColor
    End With
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    End If
End Sub

Private Sub UserControl_InitProperties()
    v_sSkinPath = App.Path & "\Skins\Titanium"
    v_sCaption = DefCaption
    v_oForeColor = DefForeColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    v_sSkinPath = PropBag.ReadProperty("SkinPath", App.Path & "\Skins\Titanium")
    Call LoadSkin
    
    v_sCaption = PropBag.ReadProperty("Caption", DefCaption)
    UserControl.lbl_Caption.Caption = v_sCaption
    
    v_oForeColor = PropBag.ReadProperty("ForeColor", DefForeColor)
    UserControl.lbl_Caption.ForeColor = v_oForeColor

    v_bEnabled = PropBag.ReadProperty("Enabled", DefEnabled)
    If v_bEnabled = True Then
        Call Refresh
    Else
        UserControl.lbl_Caption.Enabled = False
    End If
End Sub

Private Sub UserControl_Resize()
    Call Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SkinPath", v_sSkinPath, App.Path & "\Skins\Titanium")
    Call PropBag.WriteProperty("Caption", v_sCaption, DefCaption)
    Call PropBag.WriteProperty("ForeColor", v_oForeColor, DefForeColor)
    Call PropBag.WriteProperty("Enabled", v_bEnabled, DefEnabled)
End Sub
