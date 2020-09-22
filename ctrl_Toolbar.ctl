VERSION 5.00
Begin VB.UserControl ctrl_Toolbar 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   360
   ScaleWidth      =   4800
   Begin VB.PictureBox pic_TbrButton 
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
      Begin VB.Image img_Icon 
         Height          =   180
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   180
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
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "ctrl_Toolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Const DefBackColor = 0
Const DefIconLeft = 90
Const DefIconTop = 90

Dim v_oBackColor As OLE_COLOR
Dim v_sSkinPath As String
Dim v_iIconLeft As Integer
Dim v_iIconTop As Integer
Dim v_iItemCount As Integer
Dim v_iLastItem As Integer
Dim v_bRefreshed As Boolean

Event Click(Index As Integer)
Event MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Sub DrawToolbar()
    With UserControl
        .Height = 360
        .pic_TbrButton.BackColor = BackColor
        .pic_Source.Picture = LoadPicture(SkinPath & "\img_ToolbarBtns.bmp")
        .pic_TbrButton.Width = .Width
        .pic_TbrButton.Height = Height
        v_bRefreshed = False
    End With
End Sub

Public Sub AddButton(m_Picture As IPictureDisp)
    Dim v_lRtn As Long

    With UserControl
        v_lRtn = BitBlt(.pic_TbrButton.hDC, (v_iItemCount * 24) + 5 * v_iItemCount, 0, 24, 24, .pic_Source.hDC, 0, 0, SRCCOPY)
        .img_Icon(v_iItemCount).Picture = m_Picture
        .img_Icon(v_iItemCount).Top = IconTop
        .img_Icon(v_iItemCount).Left = (v_iItemCount * 360) + v_iIconLeft + 75 * v_iItemCount
        .img_Icon(v_iItemCount).Visible = True
        
        v_iItemCount = v_iItemCount + 1
        Load .img_Icon(v_iItemCount)
    End With
End Sub

Public Sub AddTooltipText(m_Index As Integer, m_Tooltip As String)
    UserControl.img_Icon(m_Index).ToolTipText = m_Tooltip
End Sub

Public Sub Refresh()
    Dim v_lRtn As Long
    Dim v_iLoop As Integer

    If v_bRefreshed = False Then
    
    With UserControl
        .pic_TbrButton.Cls
        For v_iLoop = 0 To v_iItemCount - 1
            v_lRtn = BitBlt(.pic_TbrButton.hDC, (v_iLoop * 24) + v_iLoop * 5, 0, 24, 24, .pic_Source.hDC, 0, 0, SRCCOPY)
        Next v_iLoop
        UserControl.pic_TbrButton.Refresh
    End With
    v_bRefreshed = True
    
    Else
    End If
End Sub

Public Sub UnloadButtons()
    Dim v_iLoop As Integer
    
    For v_iLoop = 1 To v_iItemCount
        Unload UserControl.img_Icon(v_iLoop)
        UserControl.img_Icon(0).Visible = False
    Next v_iLoop
    v_iItemCount = 0
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

Public Property Get IconLeft() As Integer
    IconLeft = v_iIconLeft
End Property

Public Property Let IconLeft(ByVal m_IconLeft As Integer)
    v_iIconLeft = m_IconLeft
    PropertyChanged "IconLeft"
End Property

Public Property Get IconTop() As Integer
    IconTop = v_iIconTop
End Property

Public Property Let IconTop(ByVal m_IconTop As Integer)
    v_iIconTop = m_IconTop
    PropertyChanged "IconTop"
End Property

Private Sub img_Icon_Click(Index As Integer)
    RaiseEvent Click(v_iLastItem)
End Sub

Private Sub img_Icon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Refresh
    End If
End Sub

Private Sub img_Icon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim v_lRtn As Long
    Dim v_iLoop As Integer

    If Button = 0 Then

    RaiseEvent MouseMove(Index, Button, Shift, X, Y)
    With UserControl
        .pic_TbrButton.Cls
        For v_iLoop = 0 To v_iItemCount - 1
            If v_iLoop <> Index Then
                v_lRtn = BitBlt(.pic_TbrButton.hDC, (v_iLoop * 24) + 5 * v_iLoop, 0, 24, 24, .pic_Source.hDC, 0, 0, SRCCOPY)
            Else
                v_lRtn = BitBlt(.pic_TbrButton.hDC, (v_iLoop * 24) + 5 * v_iLoop, 0, 24, 24, .pic_Source.hDC, 24, 0, SRCCOPY)
            End If
        Next v_iLoop
    End With
    v_iLastItem = Index
    v_bRefreshed = False
    
    End If
End Sub

Private Sub UserControl_InitProperties()
    v_sSkinPath = App.Path & "\Skins\Titanium"
    v_oBackColor = DefBackColor
    v_iIconLeft = DefIconLeft
    v_iIconTop = DefIconTop
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    v_sSkinPath = PropBag.ReadProperty("SkinPath", App.Path & "\Skins\Titanium")
    Call DrawToolbar
    
    v_oBackColor = PropBag.ReadProperty("BackColor", DefBackColor)
    UserControl.pic_TbrButton.BackColor = v_oBackColor

    v_iIconLeft = PropBag.ReadProperty("IconLeft", DefIconLeft)
    UserControl.img_Icon(0).Left = v_iIconLeft
    
    v_iIconTop = PropBag.ReadProperty("IconTop", DefIconTop)
    UserControl.img_Icon(0).Top = v_iIconTop
End Sub

Private Sub UserControl_Resize()
    Call Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SkinPath", v_sSkinPath, App.Path & "\Skins\Titanium")
    Call PropBag.WriteProperty("BackColor", v_oBackColor, DefBackColor)
    Call PropBag.WriteProperty("IconLeft", v_iIconLeft, DefIconLeft)
    Call PropBag.WriteProperty("IconTop", v_iIconTop, DefIconTop)
End Sub
