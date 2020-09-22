VERSION 5.00
Begin VB.UserControl ctrl_ListObject 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2100
   ScaleHeight     =   2985
   ScaleWidth      =   2100
   Begin VB.PictureBox pic_Viewport 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   600
      Width           =   1215
      Begin VB.PictureBox pic_MouseMove 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   1215
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Label lbl_MouseMove 
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
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   330
         End
      End
      Begin VB.Label lbl_Item 
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
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
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
      Top             =   1200
      Width           =   1215
      Begin VB.Image img_MoveDown 
         Height          =   360
         Left            =   0
         Top             =   0
         Width           =   300
      End
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
      Begin VB.Image img_MoveUp 
         Height          =   360
         Left            =   0
         Top             =   0
         Width           =   300
      End
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
Attribute VB_Name = "ctrl_ListObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Const DefForeColor = 0
Const DefMouseMoveColor = 0
Const DefMouseDownColor = 0

Dim v_sSkinPath As String
Dim v_oForeColor As OLE_COLOR
Dim v_oMouseMoveColor As OLE_COLOR
Dim v_oMouseDownColor As OLE_COLOR
Dim v_iItemCount As Integer
Dim v_iLastItem As Integer

Event Click(Index As Integer)
Event MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Sub DrawMenu()
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer
    Dim v_iCurrentY As Integer

    With UserControl
        .pic_Source.Picture = LoadPicture(SkinPath & "\img_ListObject.bmp")
        .pic_UpBorder.Width = .Width
        .pic_UpBorder.Height = 360
        
        .pic_UpBorder.Cls
        v_lRtn = BitBlt(.pic_UpBorder.hDC, 0, 0, 20, 24, .pic_Source.hDC, 0, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 20)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_UpBorder.hDC, v_iLoop * 20, 0, 20, 24, .pic_Source.hDC, 23, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_UpBorder.hDC, (.Width / Screen.TwipsPerPixelX) - 23, 0, 23, 24, .pic_Source.hDC, 44, 0, SRCCOPY)
        .pic_UpBorder.Refresh
            
        .pic_DownBorder.Cls
        .pic_DownBorder.Width = .Width
        .pic_DownBorder.Height = 360
        .pic_DownBorder.Top = .Height - .pic_DownBorder.Height
        v_lRtn = BitBlt(.pic_DownBorder.hDC, 0, 0, 20, 24, .pic_Source.hDC, 0, 96, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 20)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_DownBorder.hDC, v_iLoop * 20, 0, 20, 24, .pic_Source.hDC, 23, 96, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_DownBorder.hDC, (.Width / Screen.TwipsPerPixelX) - 23, 0, 23, 24, .pic_Source.hDC, 44, 96, SRCCOPY)
        .pic_DownBorder.Refresh
    
        .pic_Viewport.Top = .pic_UpBorder.Height
        .pic_Viewport.Width = .Width
        .pic_Viewport.Height = .Height - .pic_UpBorder.Height - .pic_DownBorder.Height
        
        .pic_Viewport.Cls
        v_iCurrentY = 0
        While (v_iCurrentY * 15) < (.Height - 720)
            v_lRtn = BitBlt(.pic_Viewport.hDC, 0, v_iCurrentY, 20, 24, .pic_Source.hDC, 0, 24, SRCCOPY)
            v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 20)
            If v_iCenterImgFrequency > 0 Then
                For v_iLoop = 1 To v_iCenterImgFrequency
                    v_lRtn = BitBlt(.pic_Viewport.hDC, v_iLoop * 20, v_iCurrentY, 20, 24, .pic_Source.hDC, 23, 24, SRCCOPY)
                Next v_iLoop
            End If
            v_lRtn = BitBlt(.pic_Viewport.hDC, (.Width / Screen.TwipsPerPixelX) - 23, v_iCurrentY, 23, 24, .pic_Source.hDC, 44, 24, SRCCOPY)
            v_iCurrentY = v_iCurrentY + 24
        Wend
        .pic_Viewport.Refresh
    End With
End Sub

Public Sub Refresh()
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer
    Dim v_iCurrentY As Integer

    With UserControl
        .pic_UpBorder.Width = .Width
        .pic_UpBorder.Height = 360
        
        v_lRtn = BitBlt(.pic_UpBorder.hDC, 0, 0, 20, 24, .pic_Source.hDC, 0, 0, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 20)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_UpBorder.hDC, v_iLoop * 20, 0, 20, 24, .pic_Source.hDC, 20, 0, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_UpBorder.hDC, (.Width / Screen.TwipsPerPixelX) - 23, 0, 23, 24, .pic_Source.hDC, 44, 0, SRCCOPY)
        pic_UpBorder.Refresh
            
        .pic_DownBorder.Width = .Width
        .pic_DownBorder.Height = 360
        .pic_DownBorder.Top = .Height - .pic_DownBorder.Height
        v_lRtn = BitBlt(.pic_DownBorder.hDC, 0, 0, 20, 24, .pic_Source.hDC, 0, 96, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 20)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_DownBorder.hDC, v_iLoop * 20, 0, 20, 24, .pic_Source.hDC, 20, 96, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_DownBorder.hDC, (.Width / Screen.TwipsPerPixelX) - 23, 0, 23, 24, .pic_Source.hDC, 44, 96, SRCCOPY)
    
        .pic_Viewport.Top = .pic_UpBorder.Height
        .pic_Viewport.Width = .Width
        .pic_Viewport.Height = .Height - .pic_UpBorder.Height - .pic_DownBorder.Height
        
        v_iCurrentY = 0
        While (v_iCurrentY * 15) < (.Height - 720)
            v_lRtn = BitBlt(.pic_Viewport.hDC, 0, v_iCurrentY, 20, 24, .pic_Source.hDC, 0, 24, SRCCOPY)
            v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 20)
            If v_iCenterImgFrequency > 0 Then
                For v_iLoop = 1 To v_iCenterImgFrequency
                    v_lRtn = BitBlt(.pic_Viewport.hDC, v_iLoop * 20, v_iCurrentY, 20, 24, .pic_Source.hDC, 20, 24, SRCCOPY)
                Next v_iLoop
            End If
            v_lRtn = BitBlt(.pic_Viewport.hDC, (.Width / Screen.TwipsPerPixelX) - 23, v_iCurrentY, 23, 24, .pic_Source.hDC, 44, 24, SRCCOPY)
            v_iCurrentY = v_iCurrentY + 24
        Wend
    End With
End Sub

Public Sub AddItem(m_Item As String)
    With UserControl
        If v_iItemCount <> 0 Then
            Load .lbl_Item(v_iItemCount)
        End If
        .lbl_Item(v_iItemCount).Width = .Width
        .lbl_Item(v_iItemCount).Top = 360 * v_iItemCount + 75
        .lbl_Item(v_iItemCount).Caption = m_Item
        .lbl_Item(v_iItemCount).Visible = True
        v_iItemCount = v_iItemCount + 1
    End With
End Sub

Private Sub UnloadItems()
    Dim v_iLoop As Integer
    
    For v_iLoop = 1 To v_iItemCount - 1
        Unload UserControl.lbl_Item(v_iLoop)
    Next v_iLoop
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
    Dim v_iLoop As Integer
    
    v_oForeColor = m_ForeColor
    PropertyChanged "ForeColor"
    
    For v_iLoop = 0 To v_iItemCount - 1
        UserControl.lbl_Item(v_iLoop).ForeColor = v_oForeColor
    Next v_iLoop
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

Private Sub img_MoveDown_Click()
    Dim v_iLoop As Integer
    
    For v_iLoop = 0 To v_iItemCount - 1
        UserControl.lbl_Item(v_iLoop).Top = UserControl.lbl_Item(v_iLoop).Top - 360
        UserControl.pic_MouseMove.Visible = False
    Next v_iLoop
End Sub

Private Sub img_MoveUp_Click()
    Dim v_iLoop As Integer
    
    For v_iLoop = 0 To v_iItemCount - 1
        UserControl.lbl_Item(v_iLoop).Top = UserControl.lbl_Item(v_iLoop).Top + 360
        UserControl.pic_MouseMove.Visible = False
    Next v_iLoop
End Sub

Private Sub lbl_Item_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    RaiseEvent MouseMove(Index, Button, Shift, X, Y)
    v_iLastItem = Index
    With UserControl
        .pic_MouseMove.Width = .Width
        .pic_MouseMove.Height = 360
    
        .pic_MouseMove.Cls
        v_lRtn = BitBlt(.pic_MouseMove.hDC, 0, 0, 20, 24, .pic_Source.hDC, 0, 48, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 20)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_MouseMove.hDC, v_iLoop * 20, 0, 20, 24, .pic_Source.hDC, 23, 48, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_MouseMove.hDC, (.Width / Screen.TwipsPerPixelX) - 23, 0, 23, 24, .pic_Source.hDC, 44, 48, SRCCOPY)
        
        .pic_MouseMove.Top = .lbl_Item(Index).Top - 75
        .lbl_MouseMove.Caption = .lbl_Item(Index).Caption
        .lbl_MouseMove.ForeColor = MouseMoveColor
        .lbl_MouseMove.Width = .Width
        .lbl_MouseMove.Top = 75
        .pic_MouseMove.Visible = True
    End With
End Sub

Private Sub lbl_MouseMove_Click()
    RaiseEvent Click(v_iLastItem)
End Sub

Private Sub lbl_MouseMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim v_lRtn As Long
    Dim v_iCenterImgFrequency As Integer
    Dim v_iLoop As Integer

    With UserControl
        .pic_MouseMove.Width = .Width
        .pic_MouseMove.Height = 360
    
        .pic_MouseMove.Cls
        .lbl_MouseMove.ForeColor = MouseDownColor
        v_lRtn = BitBlt(.pic_MouseMove.hDC, 0, 0, 20, 24, .pic_Source.hDC, 0, 72, SRCCOPY)
        v_iCenterImgFrequency = Abs((.Width / Screen.TwipsPerPixelX) / 20)
        If v_iCenterImgFrequency > 0 Then
            For v_iLoop = 1 To v_iCenterImgFrequency
                v_lRtn = BitBlt(.pic_MouseMove.hDC, v_iLoop * 20, 0, 20, 24, .pic_Source.hDC, 23, 72, SRCCOPY)
            Next v_iLoop
        End If
        v_lRtn = BitBlt(.pic_MouseMove.hDC, (.Width / Screen.TwipsPerPixelX) - 23, 0, 23, 24, .pic_Source.hDC, 44, 72, SRCCOPY)
    End With
End Sub

Private Sub UserControl_InitProperties()
    v_sSkinPath = App.Path & "\Skins\Titanium"
    v_oForeColor = DefForeColor
    v_oMouseMoveColor = DefMouseMoveColor
    v_oMouseDownColor = DefMouseDownColor
End Sub

Private Sub UserControl_Resize()
    Call Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    v_sSkinPath = PropBag.ReadProperty("SkinPath", App.Path & "\Skins\Titanium")
    Call DrawMenu
    
    v_oForeColor = PropBag.ReadProperty("ForeColor", DefForeColor)
    UserControl.lbl_Item(0).ForeColor = v_oForeColor

    v_oMouseMoveColor = PropBag.ReadProperty("MouseMoveColor", DefMouseMoveColor)
    UserControl.lbl_MouseMove.ForeColor = v_oMouseMoveColor

    v_oMouseDownColor = PropBag.ReadProperty("MouseDownColor", DefMouseDownColor)
End Sub

Private Sub UserControl_Terminate()
    Call UnloadItems
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SkinPath", v_sSkinPath, App.Path & "\Skins\Titanium")
    Call PropBag.WriteProperty("ForeColor", v_oForeColor, DefForeColor)
    Call PropBag.WriteProperty("MouseMoveColor", v_oMouseMoveColor, DefMouseMoveColor)
    Call PropBag.WriteProperty("MouseDownColor", v_oMouseDownColor, DefMouseDownColor)
End Sub
