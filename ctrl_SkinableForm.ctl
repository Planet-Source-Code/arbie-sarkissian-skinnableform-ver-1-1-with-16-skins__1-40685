VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctrl_SkinableForm 
   BackStyle       =   0  'Transparent
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   ScaleHeight     =   750
   ScaleWidth      =   750
   Begin VB.PictureBox pic_LeftCaption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   1200
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox pic_DownBorder 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   720
      ScaleHeight     =   150
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pic_RightBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   150
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pic_Borders 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pic_LeftBorder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pic_RightCaption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   2400
      ScaleHeight     =   720
      ScaleWidth      =   1440
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      Begin VB.Image img_MinimizeBtn 
         Height          =   300
         Left            =   810
         Top             =   0
         Width           =   285
      End
      Begin VB.Image img_MaximizeBtn 
         Height          =   300
         Left            =   540
         Top             =   0
         Width           =   285
      End
      Begin VB.Image img_RestoreBtn 
         Height          =   300
         Left            =   270
         Top             =   0
         Width           =   285
      End
      Begin VB.Image img_CloseBtn 
         Height          =   300
         Left            =   0
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox pic_CenterCaption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   1200
      ScaleHeight     =   720
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label lbl_Caption 
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
         TabIndex        =   6
         Top             =   0
         Width           =   555
      End
   End
   Begin MSComctlLib.ImageList iml_Skin 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Image img_Logo 
      Height          =   750
      Left            =   0
      Picture         =   "ctrl_SkinableForm.ctx":0000
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "ctrl_SkinableForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Const DefMaximizeBtn = 1
Const DefMinimizeBtn = 1
Const DefCaption = "Caption"
Const DefBackColor = 0
Const DefForeColor = 0
Const DefCaptionTop = 195
Const DefCaptionColor = 0

Dim v_bMaximizeBtn As Boolean
Dim v_bMinimizeBtn As Boolean
Dim v_sCaption As String
Dim v_sSkinPath As String
Dim v_oBackColor As OLE_COLOR
Dim v_oForeColor As OLE_COLOR
Dim v_iCaptionTop As Integer
Dim v_oCaptionColor As OLE_COLOR
Dim v_iMouseX, v_iMouseY As Integer
Dim v_oForm As Form

Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Public Sub LoadSkin(m_Form As Form)
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer
    Dim v_lRtn As Long

    Set v_oForm = m_Form
    With UserControl
        .Width = m_Form.Width
        .Height = m_Form.Height
        m_Form.BackColor = v_oBackColor
        m_Form.Caption = Caption
        
        .pic_LeftCaption.Visible = True
        .pic_CenterCaption.Visible = True
        .pic_RightCaption.Visible = True
        .pic_LeftBorder.Visible = True
        .pic_RightBorder.Visible = True
        .pic_DownBorder.Visible = True
        .img_Logo.Visible = False
    
        .iml_Skin.ListImages.Add 1, , LoadPicture(SkinPath & "\img_Caption_Left.bmp")
        .iml_Skin.ListImages.Add 2, , LoadPicture(SkinPath & "\img_Caption_Center.bmp")
        .iml_Skin.ListImages.Add 3, , LoadPicture(SkinPath & "\img_Caption_Right.bmp")
        .iml_Skin.ListImages.Add 4, , LoadPicture(SkinPath & "\img_Button_Close.gif")
        .iml_Skin.ListImages.Add 5, , LoadPicture(SkinPath & "\img_Button_Restore.gif")
        .iml_Skin.ListImages.Add 6, , LoadPicture(SkinPath & "\img_Button_Maximize.gif")
        .iml_Skin.ListImages.Add 7, , LoadPicture(SkinPath & "\img_Button_Minimize.gif")
        .iml_Skin.ListImages.Add 8, , LoadPicture(SkinPath & "\img_Borders.bmp")
        
        .pic_LeftCaption.Cls
        .pic_LeftCaption.Picture = .iml_Skin.ListImages(1).Picture
        .pic_LeftCaption.Refresh
        .pic_LeftCaption.Top = 0
        
        .pic_RightCaption.Cls
        .pic_RightCaption.Picture = .iml_Skin.ListImages(3).Picture
        .pic_RightCaption.Refresh
        .pic_RightCaption.Left = .Width - .pic_RightCaption.Width
        
        .pic_CenterCaption.Picture = .iml_Skin.ListImages(2).Picture
        .pic_CenterCaption.Left = .pic_LeftCaption.Width
        .pic_CenterCaption.Refresh
        .pic_CenterCaption.Width = .Width - .pic_LeftCaption.Width - .pic_RightCaption.Width
        v_iCenterImgFrequency = Abs((.pic_CenterCaption.Width / Screen.TwipsPerPixelX) / 50)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_CenterCaption.hDC, v_iLoop * 50, 0, 100, 48, .pic_CenterCaption.hDC, 0, 0, SRCCOPY)
            Next v_iLoop
        End If
        .lbl_Caption.Width = .pic_CenterCaption.Width
                        
        .img_CloseBtn.Picture = .iml_Skin.ListImages(4).Picture
        .img_CloseBtn.Left = .pic_RightCaption.Width - .img_CloseBtn.Width - 75
        .img_CloseBtn.Top = 45
    
        .img_RestoreBtn.Picture = .iml_Skin.ListImages(5).Picture
        .img_RestoreBtn.Left = .pic_RightCaption.Width - .img_RestoreBtn.Width - .img_CloseBtn.Width - 75
        .img_RestoreBtn.Top = 45
    
        .img_MaximizeBtn.Picture = .iml_Skin.ListImages(6).Picture
        .img_MaximizeBtn.Left = .pic_RightCaption.Width - .img_MaximizeBtn.Width - .img_CloseBtn.Width - 75
        .img_MaximizeBtn.Top = 45
    
        .img_MinimizeBtn.Picture = .iml_Skin.ListImages(7).Picture
        .img_MinimizeBtn.Left = .pic_RightCaption.Width - .img_MinimizeBtn.Width - .img_MaximizeBtn.Width - .img_CloseBtn.Width - 75
        .img_MinimizeBtn.Top = 45
    
        .pic_Borders.Picture = .iml_Skin.ListImages(8).Picture
        .pic_LeftBorder.Cls
        .pic_LeftBorder.Top = .pic_LeftCaption.Height
        .pic_LeftBorder.Height = .Height - .pic_LeftCaption.Height
        .pic_RightBorder.Cls
        .pic_RightBorder.Refresh
        .pic_RightBorder.Left = .Width - 150
        .pic_RightBorder.Top = .pic_RightCaption.Height
        .pic_RightBorder.Height = m_Form.Height - .pic_RightCaption.Height
        v_iCenterImgFrequency = Abs(((m_Form.Height - .pic_LeftCaption.Height) / Screen.TwipsPerPixelY) / 10)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 0 To v_iCenterImgFrequency - 1
                v_lRtn = BitBlt(.pic_LeftBorder.hDC, 0, v_iLoop * 10, 10, 10, .pic_Borders.hDC, 0, 0, SRCCOPY)
                v_lRtn = BitBlt(.pic_RightBorder.hDC, 0, v_iLoop * 10, 10, 10, .pic_Borders.hDC, 30, 0, SRCCOPY)
            Next v_iLoop
        End If
        .pic_LeftBorder.Refresh
        .pic_RightBorder.Refresh
        
        .pic_DownBorder.Cls
        .pic_DownBorder.Left = 0
        .pic_DownBorder.Top = m_Form.Height - 150
        .pic_DownBorder.Width = m_Form.Width
        .pic_DownBorder.Height = 150
        v_iCenterImgFrequency = Abs((m_Form.Width / Screen.TwipsPerPixelX) / 9)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 0 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_DownBorder.hDC, v_iLoop * 9, 0, 9, 10, .pic_Borders.hDC, 20, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_DownBorder.hDC, 0, 0, 10, 10, .pic_Borders.hDC, 10, 0, SRCCOPY)
        v_lRtn = BitBlt(.pic_DownBorder.hDC, (m_Form.Width / Screen.TwipsPerPixelX) - 10, 0, 10, 10, .pic_Borders.hDC, 40, 0, SRCCOPY)
        .pic_DownBorder.Refresh
        .lbl_Caption.Top = CaptionTop
        .lbl_Caption.ForeColor = CaptionColor
    End With
End Sub

Public Property Get MaximizeBtn() As Boolean
    MaximizeBtn = v_bMaximizeBtn
End Property

Public Property Let MaximizeBtn(ByVal m_MaximizeBtn As Boolean)
    v_bMaximizeBtn = m_MaximizeBtn
    PropertyChanged "Maximize"
End Property

Public Property Get MinimizeBtn() As Boolean
    MinimizeBtn = v_bMinimizeBtn
End Property

Public Property Let MinimizeBtn(ByVal m_MinimizeBtn As Boolean)
    v_bMinimizeBtn = m_MinimizeBtn
    PropertyChanged "Minimize"
End Property

Public Property Get Caption() As String
    Caption = v_sCaption
End Property

Public Property Let Caption(ByVal m_Caption As String)
    v_sCaption = m_Caption
    PropertyChanged "Caption"
End Property

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

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = v_oForeColor
End Property

Public Property Let ForeColor(ByVal m_ForeColor As OLE_COLOR)
    v_oForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get CaptionTop() As Integer
    CaptionTop = v_iCaptionTop
End Property

Public Property Let CaptionTop(ByVal m_CaptionTop As Integer)
    v_iCaptionTop = m_CaptionTop
    PropertyChanged "CaptionTop"
End Property

Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = v_oCaptionColor
End Property

Public Property Let CaptionColor(ByVal m_CaptionColor As OLE_COLOR)
    v_oCaptionColor = m_CaptionColor
    PropertyChanged "CaptionColor"
End Property

Private Sub img_CloseBtn_Click()
    Unload Screen.ActiveForm
End Sub

Private Sub img_MaximizeBtn_Click()
    Screen.ActiveForm.WindowState = 2
    UserControl.img_MaximizeBtn.Visible = False
    UserControl.img_RestoreBtn.Visible = True
    Call LoadSkin(v_oForm)
End Sub

Private Sub img_MinimizeBtn_Click()
    Screen.ActiveForm.WindowState = 1
End Sub

Private Sub img_RestoreBtn_Click()
    Screen.ActiveForm.WindowState = 0
    UserControl.img_MaximizeBtn.Visible = True
    UserControl.img_RestoreBtn.Visible = False
    Call LoadSkin(v_oForm)
End Sub

Private Sub lbl_Caption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        v_iMouseX = X
        v_iMouseY = Y
    End If
End Sub

Private Sub lbl_Caption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (v_oForm.WindowState <> 2) Then
        Screen.ActiveForm.Left = Screen.ActiveForm.Left + X - v_iMouseX
        Screen.ActiveForm.Top = Screen.ActiveForm.Top + Y - v_iMouseY
    End If
End Sub

Private Sub pic_CenterCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        v_iMouseX = X
        v_iMouseY = Y
    End If
End Sub

Private Sub pic_CenterCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (v_oForm.WindowState <> 2) Then
        Screen.ActiveForm.Left = Screen.ActiveForm.Left + X - v_iMouseX
        Screen.ActiveForm.Top = Screen.ActiveForm.Top + Y - v_iMouseY
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
    v_bMaximizeBtn = DefMaximizeBtn
    v_bMinimizeBtn = DefMinimizeBtn
    v_sCaption = DefCaption
    v_sSkinPath = App.Path & "\Skins\Titanium"
    v_oBackColor = DefBackColor
    v_oForeColor = DefForeColor
    v_oCaptionColor = DefCaptionColor
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    v_bMaximizeBtn = PropBag.ReadProperty("MaximizeBtn", DefMaximizeBtn)
    If v_bMaximizeBtn = False Then
        UserControl.img_MaximizeBtn.Visible = False
        UserControl.img_RestoreBtn.Visible = False
        UserControl.img_MinimizeBtn.Left = UserControl.pic_RightCaption.Width - UserControl.img_MinimizeBtn.Width - UserControl.img_CloseBtn.Width - 75
    Else
        UserControl.img_MaximizeBtn.Visible = True
        UserControl.img_RestoreBtn.Visible = False
        UserControl.img_MinimizeBtn.Left = UserControl.pic_RightCaption.Width - UserControl.img_MinimizeBtn.Width - UserControl.img_CloseBtn.Width - UserControl.img_MaximizeBtn.Width - 75
    End If

    v_bMinimizeBtn = PropBag.ReadProperty("MinimizeBtn", DefMinimizeBtn)
    If v_bMinimizeBtn = False Then
        UserControl.img_MinimizeBtn.Visible = False
    Else
        UserControl.img_MinimizeBtn.Visible = True
    End If
    
    v_sCaption = PropBag.ReadProperty("Caption", DefCaption)
    UserControl.lbl_Caption.Caption = v_sCaption
    
    v_sSkinPath = PropBag.ReadProperty("SkinPath", App.Path & "\Skins\Titanium")
    v_oBackColor = PropBag.ReadProperty("BackColor", DefBackColor)
    
    v_oForeColor = PropBag.ReadProperty("ForeColor", DefForeColor)
    UserControl.lbl_Caption.ForeColor = v_oForeColor
    
    v_iCaptionTop = PropBag.ReadProperty("CaptionTop", DefCaptionTop)
    UserControl.lbl_Caption.Top = v_iCaptionTop

    v_oCaptionColor = PropBag.ReadProperty("CaptionColor", DefCaptionColor)
    UserControl.lbl_Caption.ForeColor = v_oCaptionColor
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("MaximizeBtn", v_bMaximizeBtn, DefMaximizeBtn)
    Call PropBag.WriteProperty("MinimizeBtn", v_bMinimizeBtn, DefMinimizeBtn)
    Call PropBag.WriteProperty("Caption", v_sCaption, DefCaption)
    Call PropBag.WriteProperty("SkinPath", v_sSkinPath, App.Path & "\Skins\Titanium")
    Call PropBag.WriteProperty("BackColor", v_oBackColor, DefBackColor)
    Call PropBag.WriteProperty("ForeColor", v_oForeColor, DefForeColor)
    Call PropBag.WriteProperty("CaptionTop", v_iCaptionTop, DefCaptionTop)
    Call PropBag.WriteProperty("CaptionColor", v_oCaptionColor, DefCaptionColor)
End Sub
