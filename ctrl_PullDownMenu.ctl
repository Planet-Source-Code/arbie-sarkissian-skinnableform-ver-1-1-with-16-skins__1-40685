VERSION 5.00
Begin VB.UserControl ctrl_PullDownMenu 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   ScaleHeight     =   360
   ScaleWidth      =   5325
   Begin VB.Line lin_Line 
      Visible         =   0   'False
      X1              =   2040
      X2              =   3240
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape shp_MouseMove 
      Height          =   255
      Left            =   1200
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl_Item 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape shp_Border 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "ctrl_PullDownMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const DefForeColor = 0
Const DefBackColor = 0
Const DefHideBorder = 0

Dim v_oForeColor As OLE_COLOR
Dim v_oBackColor As OLE_COLOR
Dim v_bHideBorder As Boolean
Dim v_iItemCount As Integer

Public pSelectionLeft, pSelectionBottom As Integer

Event Click(Index As Integer)
Event MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub lbl_Item_Click(Index As Integer)
    RaiseEvent Click(Index)
End Sub

Private Sub lbl_Item_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With UserControl
        RaiseEvent MouseMove(Index, Button, Shift, X, Y)
        .shp_MouseMove.Left = .lbl_Item(Index).Left - 15
        .shp_MouseMove.Width = .lbl_Item(Index).Width + 30
        .shp_MouseMove.Visible = True
        
        pSelectionLeft = .shp_MouseMove.Left
        pSelectionBottom = .shp_MouseMove.Top + .shp_MouseMove.Height
    End With
End Sub

Private Sub UserControl_Initialize()
    With UserControl
        .shp_Border.Width = .Width
        .shp_Border.Height = 360
        .lin_Line.X1 = 0
        .lin_Line.Y1 = .Height - 15
        .lin_Line.X2 = .Width
        .lin_Line.Y2 = .lin_Line.Y1
        .lbl_Item(0).Left = -260
        
        .shp_MouseMove.Top = 45
        .shp_MouseMove.Height = 260
    End With
End Sub

Private Sub UserControl_Resize()
    Call UserControl_Initialize
End Sub

Public Sub AddItem(m_Item As String)
    With UserControl
        v_iItemCount = v_iItemCount + 1
        Load .lbl_Item(v_iItemCount)
        .lbl_Item(v_iItemCount).Caption = m_Item
        .lbl_Item(v_iItemCount).ForeColor = .shp_Border.BorderColor
        .lbl_Item(v_iItemCount).Width = TextWidth(m_Item) + 150
        .lbl_Item(v_iItemCount).Left = .lbl_Item(v_iItemCount - 1).Left + .lbl_Item(v_iItemCount - 1).Width + 75
        .lbl_Item(v_iItemCount).Top = 75
        .lbl_Item(v_iItemCount).Visible = True
    End With
End Sub

Public Sub Refresh()
    Dim v_iLoop As Integer

    UserControl.BackColor = BackColor
    For v_iLoop = 1 To v_iItemCount
        UserControl.lbl_Item(v_iLoop).ForeColor = ForeColor
    Next v_iLoop
    UserControl.shp_Border.BorderColor = ForeColor
    UserControl.shp_MouseMove.BorderColor = ForeColor
    UserControl.lin_Line.BorderColor = ForeColor
End Sub

Private Sub UnloadItems()
    Dim v_iLoop As Integer
    
    For v_iLoop = 1 To v_iItemCount - 1
        Unload UserControl.lbl_Item(v_iLoop)
    Next v_iLoop
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = v_oForeColor
End Property

Public Property Let ForeColor(ByVal m_ForeColor As OLE_COLOR)
    v_oForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = v_oBackColor
End Property

Public Property Let BackColor(ByVal m_BackColor As OLE_COLOR)
    v_oBackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get HideBorder() As Boolean
    HideBorder = v_bHideBorder
End Property

Public Property Let HideBorder(ByVal m_HideBorder As Boolean)
    v_bHideBorder = m_HideBorder
    PropertyChanged "HideBorder"
End Property

Private Sub UserControl_InitProperties()
    v_oForeColor = DefForeColor
    v_oBackColor = DefBackColor
    v_bHideBorder = DefHideBorder
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    v_oForeColor = PropBag.ReadProperty("ForeColor", DefForeColor)
    UserControl.shp_Border.BorderColor = v_oForeColor
    UserControl.shp_MouseMove.BorderColor = v_oForeColor
    UserControl.lin_Line.BorderColor = v_oForeColor
    
    v_oBackColor = PropBag.ReadProperty("BackColor", DefBackColor)
    UserControl.BackColor = v_oBackColor
    
    v_bHideBorder = PropBag.ReadProperty("HideBorder", DefHideBorder)
    If v_bHideBorder = True Then
        UserControl.shp_Border.Visible = False
        UserControl.lin_Line.Visible = True
    Else
        UserControl.shp_Border.Visible = True
        UserControl.lin_Line.Visible = False
    End If
End Sub

Private Sub UserControl_Terminate()
    Call UnloadItems
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ForeColor", v_oForeColor, DefForeColor)
    Call PropBag.WriteProperty("BackColor", v_oBackColor, DefBackColor)
    Call PropBag.WriteProperty("HideBorder", v_bHideBorder, DefHideBorder)
End Sub

