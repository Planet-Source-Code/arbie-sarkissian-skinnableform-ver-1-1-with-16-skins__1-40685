VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Main 
   BorderStyle     =   0  'None
   ClientHeight    =   6525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin SkinableForm.ctrl_TransparetForm ctrl_TransparetForm 
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.PictureBox pic_Viewport 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   3240
      ScaleHeight     =   3015
      ScaleWidth      =   4215
      TabIndex        =   10
      Top             =   1920
      Width           =   4215
      Begin VB.TextBox tbx_Text 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   3975
      End
   End
   Begin SkinableForm.ctrl_ChannelBar ctrl_ChannelBar 
      Height          =   735
      Left            =   225
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1296
      MouseDownColor  =   16777215
   End
   Begin SkinableForm.ctrl_SkinableButton ctrl_btn_Previous 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Previous"
   End
   Begin MSComctlLib.ImageList iml_Toolbar 
      Left            =   720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":0370
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":06E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":0A71
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":11AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1540
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":18D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1C5B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SkinableForm.ctrl_Toolbar ctrl_Toolbar 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1170
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   635
      BackColor       =   4934218
   End
   Begin SkinableForm.ctrl_PullDownMenu ctrl_PullDownMenu 
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   661
      ForeColor       =   16777215
      BackColor       =   4934218
      HideBorder      =   -1  'True
   End
   Begin SkinableForm.ctrl_Panel ctrl_Panel 
      Height          =   3255
      Left            =   3120
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5741
   End
   Begin SkinableForm.ctrl_ListObject ctrl_ListObject 
      Height          =   3975
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   7011
      MouseDownColor  =   16777215
   End
   Begin SkinableForm.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
      _ExtentX        =   7223
      _ExtentY        =   2990
      Caption         =   "SkinableForm copyright(c) 2002 by Arbie Sarkissian"
      BackColor       =   4934218
   End
   Begin SkinableForm.ctrl_SkinableButton ctrl_btn_Next 
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Next"
   End
   Begin SkinableForm.ctrl_SkinableButton ctrl_btn_Exit 
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Exit"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   105
      X2              =   8025
      Y1              =   6075
      Y2              =   6075
   End
   Begin VB.Label lbl_Statusbar 
      BackStyle       =   0  'Transparent
      Caption         =   "SkinableForm project version 1.1 copyight(c) 2002 by Arbie Sarkissian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   345
      TabIndex        =   5
      Top             =   6120
      Width           =   7545
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctrl_btn_Exit_Click()
    Unload frm_Main
End Sub

Private Sub ctrl_btn_Next_Click()
    If pIndex < 8 Then
        pIndex = pIndex + 1
    Else
        pIndex = 0
    End If
    Call Showtutorials(pIndex)
End Sub

Private Sub ctrl_btn_Previous_Click()
    If pIndex > 0 Then
        pIndex = pIndex - 1
    Else
        pIndex = 8
    End If
    Call Showtutorials(pIndex)
End Sub

Private Sub ctrl_ChannelBar_Click(Index As Integer)
    Select Case Index
        Case 1:
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Wecome")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Exit")
        Case 2:
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Pulldown Menu")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("ChannelBar")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Toolbox")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Status")
        Case 3:
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Titanium")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Default")
        Case 4:
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Form")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Button")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Menu")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("ChannelBar")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("ListObject")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Toolbar")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Panel")
        Case 5:
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Introduction")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("Coming Soon")
            Call frm_Main.ctrl_ChannelBar.AddSubItem("About")
    End Select
End Sub

Private Sub ctrl_ChannelBar_SubClick(Index As Integer, SubIndex As Integer)
    Select Case Index
        Case 1:
            If SubIndex = 1 Then
                Open App.Path & "\Welcome.txt" For Input As #1
                    frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
                Close #1
            Else
                Unload frm_Main
            End If
        Case 2:
            If (SubIndex = 1) Or (SubIndex = 2) Or (SubIndex = 3) Then
                frm_Menu.ppi_ShowChannelBar.Checked = False
                frm_Main.ctrl_PullDownMenu.Visible = True
                frm_Main.ctrl_Toolbar.Visible = True
                frm_Main.ctrl_ChannelBar.Visible = False
            Else
                If frm_Main.lbl_Statusbar.Visible = True Then
                    frm_Main.lbl_Statusbar.Visible = False
                Else
                    frm_Main.lbl_Statusbar.Visible = True
                End If
            End If
        Case 3:
            If SubIndex = 1 Then
                Call ChangeSkinToTitanium
            Else
                Call ChangeSkinToDefault
            End If
        Case 4:
            If SubIndex = 1 Then
                pIndex = 1
                Call Showtutorials(pIndex)
            ElseIf SubIndex = 2 Then
                pIndex = 2
                Call Showtutorials(pIndex)
            ElseIf SubIndex = 3 Then
                pIndex = 3
                Call Showtutorials(pIndex)
            ElseIf SubIndex = 4 Then
                pIndex = 4
                Call Showtutorials(pIndex)
            ElseIf SubIndex = 5 Then
                pIndex = 5
                Call Showtutorials(pIndex)
            ElseIf SubIndex = 6 Then
                pIndex = 6
                Call Showtutorials(pIndex)
            ElseIf SubIndex = 7 Then
                pIndex = 7
                Call Showtutorials(pIndex)
            End If
        Case 5:
            If SubIndex = 1 Then
                pIndex = 0
                Call Showtutorials(pIndex)
            ElseIf SubIndex = 2 Then
            ElseIf SubIndex = 3 Then
                frm_About.Show 1
            End If
    End Select
End Sub

Private Sub ctrl_ListObject_Click(Index As Integer)
    pIndex = Index
    Call Showtutorials(Index)
End Sub

Private Sub ctrl_ListObject_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Index
        Case 0:
            frm_Main.lbl_Statusbar.Caption = "A brief introduction about this project"
        Case 1:
            frm_Main.lbl_Statusbar.Caption = "A brief tutorial about SkinableForm ActiveX control"
        Case 2:
            frm_Main.lbl_Statusbar.Caption = "A brief tutorial about SkinableButton ActiveX control"
        Case 3:
            frm_Main.lbl_Statusbar.Caption = "A brief tutorial about Pulldown Menu ActiveX control"
        Case 4:
            frm_Main.lbl_Statusbar.Caption = "A brief tutorial about ChannelBar ActiveX control"
        Case 5:
            frm_Main.lbl_Statusbar.Caption = "A brief tutorial about ListObject ActiveX control"
        Case 6:
            frm_Main.lbl_Statusbar.Caption = "A brief tutorial about Toolbar ActiveX control"
        Case 7:
            frm_Main.lbl_Statusbar.Caption = "A brief tutorial about Panel ActiveX control"
        Case 8:
            frm_Main.lbl_Statusbar.Caption = "That you'll see soon, in next version of SkinableForm project"
    End Select
End Sub

Private Sub ctrl_PullDownMenu_Click(Index As Integer)
    Select Case Index
        Case 1:
            PopupMenu frm_Menu.ppm_Start, , frm_Main.ctrl_PullDownMenu.Left + frm_Main.ctrl_PullDownMenu.pSelectionLeft, frm_Main.ctrl_PullDownMenu.Top + frm_Main.ctrl_PullDownMenu.pSelectionBottom
        Case 2:
            PopupMenu frm_Menu.ppm_View, , frm_Main.ctrl_PullDownMenu.Left + frm_Main.ctrl_PullDownMenu.pSelectionLeft, frm_Main.ctrl_PullDownMenu.Top + frm_Main.ctrl_PullDownMenu.pSelectionBottom
        Case 3:
            PopupMenu frm_Menu.ppm_Skins, , frm_Main.ctrl_PullDownMenu.Left + frm_Main.ctrl_PullDownMenu.pSelectionLeft, frm_Main.ctrl_PullDownMenu.Top + frm_Main.ctrl_PullDownMenu.pSelectionBottom
        Case 4:
            PopupMenu frm_Menu.ppm_Tutorials, , frm_Main.ctrl_PullDownMenu.Left + frm_Main.ctrl_PullDownMenu.pSelectionLeft, frm_Main.ctrl_PullDownMenu.Top + frm_Main.ctrl_PullDownMenu.pSelectionBottom
        Case 5:
            PopupMenu frm_Menu.ppm_Help, , frm_Main.ctrl_PullDownMenu.Left + frm_Main.ctrl_PullDownMenu.pSelectionLeft, frm_Main.ctrl_PullDownMenu.Top + frm_Main.ctrl_PullDownMenu.pSelectionBottom
    End Select
End Sub

Private Sub ctrl_Toolbar_Click(Index As Integer)
    Select Case Index
        Case 0:
            Call ctrl_btn_Previous_Click
        Case 1:
            Call ctrl_btn_Next_Click
        Case 2:
            Open App.Path & "\Welcome.txt" For Input As #1
                frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
            Close #1
        Case 3:
            Call Showtutorials(pIndex)
        Case 4:
            If frm_Menu.ppi_ShowPulldownMenu.Checked = True Then
                frm_Menu.ppi_ShowPulldownMenu.Checked = False
                frm_Main.ctrl_PullDownMenu.Visible = False
            Else
                frm_Menu.ppi_ShowPulldownMenu.Checked = True
                frm_Main.ctrl_PullDownMenu.Visible = True
            End If
        Case 5:
            If frm_Menu.ppi_ShowToolbar.Checked = True Then
                frm_Menu.ppi_ShowToolbar.Checked = False
                frm_Main.ctrl_Toolbar.Visible = False
            Else
                frm_Menu.ppi_ShowToolbar.Checked = True
                frm_Main.ctrl_Toolbar.Visible = True
            End If
        Case 6:
            If frm_Menu.ppi_ShowStatus.Checked = True Then
                frm_Menu.ppi_ShowStatus.Checked = False
                frm_Main.Line1.Visible = False
                frm_Main.lbl_Statusbar.Visible = False
            Else
                frm_Menu.ppi_ShowStatus.Checked = True
                frm_Main.Line1.Visible = True
                frm_Main.lbl_Statusbar.Visible = True
            End If
        Case 7:
            frm_About.Show 1
        Case 8:
            Unload frm_Main
    End Select
End Sub

Private Sub ctrl_Toolbar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Index
        Case 0:
            frm_Main.lbl_Statusbar.Caption = "Goto previous tutorial"
        Case 1:
            frm_Main.lbl_Statusbar.Caption = "Goto Next tutorial"
        Case 2:
            frm_Main.lbl_Statusbar.Caption = "See welcome screen"
        Case 3:
            frm_Main.lbl_Statusbar.Caption = "Refresh textbox"
        Case 4:
            frm_Main.lbl_Statusbar.Caption = "Show/Hide pulldown menu"
        Case 5:
            frm_Main.lbl_Statusbar.Caption = "Show/Hide toolbar"
        Case 6:
            frm_Main.lbl_Statusbar.Caption = "Show/Hide statusbar"
        Case 7:
            frm_Main.lbl_Statusbar.Caption = "About SkinableForm project"
        Case 8:
            frm_Main.lbl_Statusbar.Caption = "Exit"
    End Select
End Sub

Private Sub Form_Load()
    Call Initialize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    frm_Main.ctrl_Toolbar.Refresh

    frm_Main.ctrl_btn_Previous.Refresh
    frm_Main.ctrl_btn_Next.Refresh
    frm_Main.ctrl_btn_Exit.Refresh
End Sub
