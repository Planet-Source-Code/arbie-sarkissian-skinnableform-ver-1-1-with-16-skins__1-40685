VERSION 5.00
Begin VB.UserControl ctrl_ChannelBar 
   BackStyle       =   0  'Transparent
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   ScaleHeight     =   720
   ScaleWidth      =   2835
   Begin VB.PictureBox pic_PDMenu 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      Begin VB.PictureBox pic_SubMouseMove 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   1215
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Label lbl_SubMouseMove 
            BackStyle       =   0  'Transparent
            Caption         =   "SubItem"
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
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.PictureBox pic_MouseMove 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   1215
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Label lbl_MouseMove 
            BackStyle       =   0  'Transparent
            Caption         =   "Item"
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
            TabIndex        =   4
            Top             =   0
            Visible         =   0   'False
            Width           =   330
         End
      End
      Begin VB.Label lbl_SubItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
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
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lbl_Item 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
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
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.PictureBox pic_PullDownMenu 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "ctrl_ChannelBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Const DefForeColor = 0
Const DefMouseMoveColor = 0
Const DefSubMouseDownColor = &HFFFFFF
Const DefSubMouseMoveColor = &HFFFFFF
Const DefMouseDownColor = 0
Const DefSubItemTop = 395

Dim v_oForeColor As OLE_COLOR
Dim v_oMouseMoveColor As OLE_COLOR
Dim v_oMouseDownColor As OLE_COLOR
Dim v_oSubMouseMoveColor As OLE_COLOR
Dim v_oSubMouseDownColor As OLE_COLOR
Dim v_sSkinPath As String
Dim v_iSubItemTop As Integer
Dim v_iItemCount As Integer
Dim v_iSubItemCount As Integer
Dim v_iLastItem As Integer
Dim v_iLastSubItem As Integer

Event Click(Index As Integer)
Event SubClick(Index As Integer, SubIndex As Integer)
Event ItemMouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event SubItemMouseMove(ItemIndex As Integer, SubItemIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Sub DrawMenu()
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    With UserControl
        .pic_PullDownMenu.Picture = LoadPicture(SkinPath & "\img_ChannelBar.bmp")
        .pic_PDMenu.Width = .Width
        .pic_PDMenu.Height = 720
        
        .pic_PDMenu.Cls
        .lbl_Item(0).Left = -210
        .lbl_SubItem(0).Left = -210
        v_lRtn = BitBlt(.pic_PDMenu.hDC, 0, 0, 8, 24, .pic_PullDownMenu.hDC, 0, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 8)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_PDMenu.hDC, v_iLoop * 8, 0, 8, 24, .pic_PullDownMenu.hDC, 80, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_PDMenu.hDC, (.Width / Screen.TwipsPerPixelX) - 8, 0, 8, 24, .pic_PullDownMenu.hDC, 223, 0, SRCCOPY)
    
        v_lRtn = BitBlt(.pic_PDMenu.hDC, 0, 24, 8, 48, .pic_PullDownMenu.hDC, 0, 24, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 8)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_PDMenu.hDC, v_iLoop * 8, 24, 8, 24, .pic_PullDownMenu.hDC, 80, 24, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_PDMenu.hDC, (.Width / Screen.TwipsPerPixelX) - 8, 24, 8, 24, .pic_PullDownMenu.hDC, 223, 24, SRCCOPY)
    End With
End Sub

Public Sub Refresh()
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    With UserControl
        .pic_PDMenu.Width = .Width
        .pic_PDMenu.Height = 720
        
        v_lRtn = BitBlt(.pic_PDMenu.hDC, 0, 0, 8, 24, .pic_PullDownMenu.hDC, 0, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 8)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_PDMenu.hDC, v_iLoop * 8, 0, 8, 24, .pic_PullDownMenu.hDC, 80, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_PDMenu.hDC, (.Width / Screen.TwipsPerPixelX) - 8, 0, 8, 24, .pic_PullDownMenu.hDC, 223, 0, SRCCOPY)
    
        v_lRtn = BitBlt(.pic_PDMenu.hDC, 0, 24, 8, 48, .pic_PullDownMenu.hDC, 0, 24, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 8)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_PDMenu.hDC, v_iLoop * 8, 24, 8, 24, .pic_PullDownMenu.hDC, 80, 24, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_PDMenu.hDC, (.Width / Screen.TwipsPerPixelX) - 8, 24, 8, 24, .pic_PullDownMenu.hDC, 223, 24, SRCCOPY)
    End With
End Sub

Public Sub AddItem(m_Item As String)
    With UserControl
        v_iItemCount = v_iItemCount + 1
        Load .lbl_Item(v_iItemCount)
        .lbl_Item(v_iItemCount).Caption = m_Item
        .lbl_Item(v_iItemCount).Width = TextWidth(.lbl_Item(v_iItemCount).Caption) + 300
        .lbl_Item(v_iItemCount).Top = 75
        .lbl_Item(v_iItemCount).Left = .lbl_Item(v_iItemCount - 1).Left + .lbl_Item(v_iItemCount - 1).Width + 180
        .lbl_Item(v_iItemCount).Visible = True
    End With
End Sub

Public Sub AddSubItem(m_SubItem As String)
    With UserControl
        v_iSubItemCount = v_iSubItemCount + 1
        Load .lbl_SubItem(v_iSubItemCount)
        .lbl_SubItem(v_iSubItemCount).Caption = m_SubItem
        .lbl_SubItem(v_iSubItemCount).Width = TextWidth(.lbl_SubItem(v_iSubItemCount).Caption) + 300
        .lbl_SubItem(v_iSubItemCount).Top = SubItemTop '395
        .lbl_SubItem(v_iSubItemCount).Left = .lbl_SubItem(v_iSubItemCount - 1).Left + .lbl_SubItem(v_iSubItemCount - 1).Width + 180
        .lbl_SubItem(v_iSubItemCount).Visible = True
    End With
End Sub

Private Sub UnloadItems()
    Dim v_iLoop As Integer
    
    For v_iLoop = 1 To v_iItemCount
        Unload UserControl.lbl_Item(v_iLoop)
    Next v_iLoop
    v_iItemCount = 0
End Sub

Private Sub UnloadSubItems()
    Dim v_iLoop As Integer
    
    For v_iLoop = 1 To v_iSubItemCount
        Unload UserControl.lbl_SubItem(v_iLoop)
    Next v_iLoop
    v_iSubItemCount = 0
End Sub

Public Property Get SkinPath() As String
    SkinPath = v_sSkinPath
End Property

Public Property Let SkinPath(ByVal m_SkinPath As String)
    v_sSkinPath = m_SkinPath
    PropertyChanged "SkinPath"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = v_oForeColor
End Property

Public Property Let ForeColor(ByVal m_ForeColor As OLE_COLOR)
    v_oForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get MouseMoveColor() As OLE_COLOR
    MouseMoveColor = v_oMouseMoveColor
End Property

Public Property Let MouseMoveColor(ByVal m_MouseMoveColor As OLE_COLOR)
    v_oMouseMoveColor = m_MouseMoveColor
    PropertyChanged "MouseMoveColor"
End Property

Public Property Get MouseDownColor() As OLE_COLOR
    MouseDownColor = v_oMouseDownColor
End Property

Public Property Let MouseDownColor(ByVal m_MouseDownColor As OLE_COLOR)
    v_oMouseDownColor = m_MouseDownColor
    PropertyChanged "MouseDownColor"
End Property

Public Property Get SubMouseMoveColor() As OLE_COLOR
    SubMouseMoveColor = v_oSubMouseMoveColor
End Property

Public Property Let SubMouseMoveColor(ByVal m_SubMouseMoveColor As OLE_COLOR)
    v_oSubMouseMoveColor = m_SubMouseMoveColor
    PropertyChanged "SubMouseMoveColor"
End Property

Public Property Get SubMouseDownColor() As OLE_COLOR
    SubMouseDownColor = v_oSubMouseDownColor
End Property

Public Property Let SubMouseDownColor(ByVal m_SubMouseDownColor As OLE_COLOR)
    v_oSubMouseDownColor = m_SubMouseDownColor
    PropertyChanged "SubMouseDownColor"
End Property

Public Property Get SubItemTop() As Integer
    SubItemTop = v_iSubItemTop
End Property

Public Property Let SubItemTop(ByVal m_SubItemTop As Integer)
    v_iSubItemTop = m_SubItemTop
    PropertyChanged "SubItemTop"
End Property

Private Sub lbl_Item_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    RaiseEvent ItemMouseMove(Index, Button, Shift, X, Y)
    With UserControl
        .pic_MouseMove.Left = .lbl_Item(Index).Left - 90
        .pic_MouseMove.Width = .lbl_Item(Index).Width + 180
        .pic_MouseMove.Height = 360
        
        .pic_MouseMove.Cls
        v_lRtn = BitBlt(.pic_MouseMove.hDC, 0, 0, 22, 24, .pic_PullDownMenu.hDC, 154, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.pic_MouseMove.Width / Screen.TwipsPerPixelX) / 22)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_MouseMove.hDC, v_iLoop * 22, 0, 22, 24, .pic_PullDownMenu.hDC, 172, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_MouseMove.hDC, (.pic_MouseMove.Width / Screen.TwipsPerPixelX) - 8, 0, 8, 24, .pic_PullDownMenu.hDC, 212, 0, SRCCOPY)
        
        .lbl_MouseMove.Caption = .lbl_Item(Index).Caption
        .lbl_MouseMove.ForeColor = MouseMoveColor
        .lbl_MouseMove.Width = .lbl_Item(Index).Width
        .lbl_MouseMove.Top = 75
        .lbl_MouseMove.Left = 240
        .lbl_MouseMove.Visible = True
        .pic_MouseMove.Visible = True
    End With
    v_iLastItem = Index
End Sub

Private Sub lbl_MouseMove_Click()
    Call UnloadSubItems
    Call Refresh
    UserControl.pic_SubMouseMove.Visible = False
    RaiseEvent Click(v_iLastItem)
End Sub

Private Sub lbl_MouseMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    If Button = 1 Then

    With UserControl
        .pic_MouseMove.Cls
        .lbl_MouseMove.ForeColor = MouseDownColor
        v_lRtn = BitBlt(.pic_MouseMove.hDC, 0, 0, 22, 24, .pic_PullDownMenu.hDC, 7, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.pic_MouseMove.Width / Screen.TwipsPerPixelX) / 8)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_MouseMove.hDC, v_iLoop * 22, 0, 22, 24, .pic_PullDownMenu.hDC, 30, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_MouseMove.hDC, (.pic_MouseMove.Width / Screen.TwipsPerPixelX) - 8, 0, 8, 24, .pic_PullDownMenu.hDC, 70, 0, SRCCOPY)
    End With
    
    End If
End Sub

Private Sub lbl_SubItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    RaiseEvent SubItemMouseMove(v_iLastItem, Index, Button, Shift, X, Y)
    With UserControl
        .pic_SubMouseMove.Left = .lbl_SubItem(Index).Left - 90
        .pic_SubMouseMove.Width = .lbl_SubItem(Index).Width + 180
        .pic_SubMouseMove.Height = 360
        
        .pic_SubMouseMove.Cls
        v_lRtn = BitBlt(.pic_SubMouseMove.hDC, 0, 0, 22, 24, .pic_PullDownMenu.hDC, 154, 24, SRCCOPY)
        v_iCenterImgFrequency = Abs((.pic_SubMouseMove.Width / Screen.TwipsPerPixelX) / 22)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_SubMouseMove.hDC, v_iLoop * 22, 0, 22, 24, .pic_PullDownMenu.hDC, 172, 24, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_SubMouseMove.hDC, (.pic_SubMouseMove.Width / Screen.TwipsPerPixelX) - 8, 0, 8, 24, .pic_PullDownMenu.hDC, 212, 24, SRCCOPY)
        
        .lbl_SubMouseMove.Caption = .lbl_SubItem(Index).Caption
        .lbl_SubMouseMove.ForeColor = SubMouseMoveColor
        .lbl_SubMouseMove.Width = .lbl_SubItem(Index).Width
        .lbl_SubMouseMove.Top = SubItemTop - 365 '30
        .lbl_SubMouseMove.Left = 210
        .lbl_SubMouseMove.Visible = True
        .pic_SubMouseMove.Visible = True
    End With
    v_iLastSubItem = Index
End Sub

Private Sub lbl_SubMouseMove_Click()
    RaiseEvent SubClick(v_iLastItem, v_iLastSubItem)
End Sub

Private Sub lbl_SubMouseMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    If Button = 1 Then

    With UserControl
        .pic_SubMouseMove.Cls
        .lbl_SubMouseMove.ForeColor = SubMouseDownColor
        v_lRtn = BitBlt(.pic_SubMouseMove.hDC, 0, 0, 22, 24, .pic_PullDownMenu.hDC, 7, 24, SRCCOPY)
        v_iCenterImgFrequency = Abs((.pic_SubMouseMove.Width / Screen.TwipsPerPixelX) / 8)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_SubMouseMove.hDC, v_iLoop * 22, 0, 22, 24, .pic_PullDownMenu.hDC, 30, 24, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_SubMouseMove.hDC, (.pic_SubMouseMove.Width / Screen.TwipsPerPixelX) - 8, 0, 8, 24, .pic_PullDownMenu.hDC, 70, 24, SRCCOPY)
    End With
    
    End If
End Sub

Private Sub UserControl_InitProperties()
    v_sSkinPath = App.Path & "\Skins\Titanium"
    v_oForeColor = DefForeColor
    v_oMouseMoveColor = DefMouseMoveColor
    v_oMouseDownColor = DefMouseDownColor
    v_oSubMouseMoveColor = DefSubMouseMoveColor
    v_oSubMouseDownColor = DefSubMouseDownColor
    v_iSubItemTop = DefSubItemTop
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    v_sSkinPath = PropBag.ReadProperty("SkinPath", App.Path & "\Skins\Titanium")
    Call DrawMenu
    
    v_oForeColor = PropBag.ReadProperty("ForeColor", DefForeColor)
    UserControl.lbl_Item(0).ForeColor = v_oForeColor

    v_oMouseMoveColor = PropBag.ReadProperty("MouseMoveColor", DefMouseMoveColor)
    UserControl.lbl_MouseMove.ForeColor = v_oMouseMoveColor
    
    v_oMouseDownColor = PropBag.ReadProperty("MouseDownColor", DefMouseDownColor)
    
    v_iSubItemTop = PropBag.ReadProperty("SubItemTop", DefSubItemTop)

    v_oSubMouseMoveColor = PropBag.ReadProperty("SubMouseMoveColor", DefSubMouseMoveColor)
    
    v_oSubMouseDownColor = PropBag.ReadProperty("SubMouseDownColor", DefSubMouseDownColor)
End Sub

Private Sub UserControl_Resize()
    Call Refresh
End Sub

Private Sub UserControl_Terminate()
    Call UnloadItems
    Call UnloadSubItems
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SkinPath", v_sSkinPath, App.Path & "\Skins\Titanium")
    Call PropBag.WriteProperty("ForeColor", v_oForeColor, DefForeColor)
    Call PropBag.WriteProperty("MouseMoveColor", v_oMouseMoveColor, DefMouseMoveColor)
    Call PropBag.WriteProperty("MouseDownColor", v_oMouseDownColor, DefMouseDownColor)
    Call PropBag.WriteProperty("SubItemTop", v_iSubItemTop, DefSubItemTop)
    Call PropBag.WriteProperty("SubMouseMoveColor", v_oSubMouseMoveColor, DefSubMouseMoveColor)
    Call PropBag.WriteProperty("SubMouseDownColor", v_oSubMouseDownColor, DefSubMouseDownColor)
End Sub
