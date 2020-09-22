VERSION 5.00
Begin VB.Form frm_Menu 
   Caption         =   "Menus"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu ppm_Start 
      Caption         =   "&Start"
      Begin VB.Menu ppi_Welcome 
         Caption         =   "&Welcome"
      End
      Begin VB.Menu Seperator01 
         Caption         =   "-"
      End
      Begin VB.Menu ppi_Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu ppm_View 
      Caption         =   "&View"
      Begin VB.Menu ppi_ShowPulldownMenu 
         Caption         =   "Pulldown Menu"
         Checked         =   -1  'True
      End
      Begin VB.Menu ppi_ShowChannelBar 
         Caption         =   "ChannelBar"
      End
      Begin VB.Menu ppi_ShowToolbar 
         Caption         =   "Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu ppi_ShowStatus 
         Caption         =   "Show Status"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu ppm_Skins 
      Caption         =   "&Skins"
      Begin VB.Menu ppi_Titanium 
         Caption         =   "Titanium"
      End
      Begin VB.Menu ppi_Default 
         Caption         =   "Default"
      End
      Begin VB.Menu ppi_Blue 
         Caption         =   "Blue"
      End
      Begin VB.Menu ppi_Deco 
         Caption         =   "Deco"
      End
      Begin VB.Menu ppi_Holograph 
         Caption         =   "Holograph"
      End
      Begin VB.Menu ppi_TreasureChest 
         Caption         =   "TreasureChest"
      End
      Begin VB.Menu ppi_ALPI 
         Caption         =   "ALPI"
      End
      Begin VB.Menu ppi_Doesnt_Suck 
         Caption         =   "Doesnt_Suck"
      End
      Begin VB.Menu ppi_SteelBlade 
         Caption         =   "SteelBlade"
      End
      Begin VB.Menu ppi_Wazoo 
         Caption         =   "Wazoo"
      End
      Begin VB.Menu ppi_SteelRain 
         Caption         =   "SteelRain"
      End
      Begin VB.Menu ppi_Coupe 
         Caption         =   "Coupe"
      End
      Begin VB.Menu ppi_BoilerRoom 
         Caption         =   "BoilerRoom"
      End
      Begin VB.Menu ppi_Executive 
         Caption         =   "Executive"
      End
      Begin VB.Menu ppi_Weaponx 
         Caption         =   "Weaponx"
      End
      Begin VB.Menu ppi_WinXP 
         Caption         =   "WinXP"
      End
   End
   Begin VB.Menu ppm_Tutorials 
      Caption         =   "T&utorials"
      Begin VB.Menu ppi_SkinableForm 
         Caption         =   "SkinableForm"
      End
      Begin VB.Menu ppi_SkinableButton 
         Caption         =   "SkinableButton"
      End
      Begin VB.Menu ppi_PulldownMenu 
         Caption         =   "PulldownMenu"
      End
      Begin VB.Menu ppi_ChannelBar 
         Caption         =   "ChannelBar"
      End
      Begin VB.Menu ppi_ListObject 
         Caption         =   "ListObject"
      End
      Begin VB.Menu ppi_Toolbar 
         Caption         =   "Toolbar"
      End
      Begin VB.Menu ppi_Panel 
         Caption         =   "Panel"
      End
   End
   Begin VB.Menu ppm_Help 
      Caption         =   "&Help"
      Begin VB.Menu ppi_Introduction 
         Caption         =   "Introduction"
      End
      Begin VB.Menu ppi_ComingSoon 
         Caption         =   "Coming Soon"
      End
      Begin VB.Menu Seperator02 
         Caption         =   "-"
      End
      Begin VB.Menu ppi_About 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ppi_About_Click()
    frm_About.Show 1
End Sub

Private Sub ppi_ALPI_Click()
    Call ChangeSkinToALPI
End Sub

Private Sub ppi_Blue_Click()
    Call ChangeSkinToBlue
End Sub

Private Sub ppi_BoilerRoom_Click()
    Call ChangeSkinToBoilerRoom
End Sub

Private Sub ppi_ChannelBar_Click()
    pIndex = 4
    Call Showtutorials(pIndex)
End Sub

Private Sub ppi_ComingSoon_Click()
    pIndex = 8
    Call Showtutorials(pIndex)
End Sub

Private Sub ppi_Coupe_Click()
    Call ChangeSkinToCoupe
End Sub

Private Sub ppi_Deco_Click()
    Call ChangeSkinToDeco
End Sub

Private Sub ppi_Default_Click()
    Call ChangeSkinToDefault
End Sub

Private Sub ppi_Doesnt_Suck_Click()
    Call ChangeSkinToDoesnt_Suck
End Sub

Private Sub ppi_Executive_Click()
    Call ChangeSkinToExecutive
End Sub

Private Sub ppi_Exit_Click()
    Unload frm_Main
    Unload frm_Menu
End Sub

Private Sub ppi_Holograph_Click()
    Call ChangeSkinToHolograph
End Sub

Private Sub ppi_Introduction_Click()
    pIndex = 0
    Call Showtutorials(pIndex)
End Sub

Private Sub ppi_ListObject_Click()
    pIndex = 5
    Call Showtutorials(pIndex)
End Sub

Private Sub ppi_Panel_Click()
    pIndex = 7
    Call Showtutorials(pIndex)
End Sub

Private Sub ppi_PulldownMenu_Click()
    pIndex = 3
    Call Showtutorials(pIndex)
End Sub

Private Sub ppi_ShowChannelBar_Click()
    If frm_Menu.ppi_ShowChannelBar.Checked = True Then
        frm_Menu.ppi_ShowChannelBar.Checked = False
        frm_Main.ctrl_PullDownMenu.Visible = True
        frm_Main.ctrl_Toolbar.Visible = True
        frm_Main.ctrl_ChannelBar.Visible = False
    Else
        frm_Menu.ppi_ShowChannelBar.Checked = True
        frm_Main.ctrl_PullDownMenu.Visible = False
        frm_Main.ctrl_Toolbar.Visible = False
        frm_Main.ctrl_ChannelBar.Visible = True
    End If
End Sub

Private Sub ppi_ShowPulldownMenu_Click()
    If frm_Menu.ppi_ShowPulldownMenu.Checked = True Then
        frm_Menu.ppi_ShowPulldownMenu.Checked = False
        frm_Main.ctrl_PullDownMenu.Visible = False
    Else
        frm_Menu.ppi_ShowPulldownMenu.Checked = True
        frm_Main.ctrl_PullDownMenu.Visible = True
    End If
End Sub

Private Sub ppi_ShowStatus_Click()
    If frm_Menu.ppi_ShowStatus.Checked = True Then
        frm_Menu.ppi_ShowStatus.Checked = False
        frm_Main.Line1.Visible = False
        frm_Main.lbl_Statusbar.Visible = False
    Else
        frm_Menu.ppi_ShowStatus.Checked = True
        frm_Main.Line1.Visible = True
        frm_Main.lbl_Statusbar.Visible = True
    End If
End Sub

Private Sub ppi_ShowToolbar_Click()
    If frm_Menu.ppi_ShowToolbar.Checked = True Then
        frm_Menu.ppi_ShowToolbar.Checked = False
        frm_Main.ctrl_Toolbar.Visible = False
    Else
        frm_Menu.ppi_ShowToolbar.Checked = True
        frm_Main.ctrl_Toolbar.Visible = True
    End If
End Sub

Private Sub ppi_SkinableButton_Click()
    pIndex = 2
    Call Showtutorials(pIndex)
End Sub

Private Sub ppi_SkinableForm_Click()
    pIndex = 1
    Call Showtutorials(pIndex)
End Sub

Private Sub ppi_SteelBlade_Click()
    Call ChangeSkinToSteelBlade
End Sub

Private Sub ppi_SteelRain_Click()
    Call ChangeSkinToSteelRain
End Sub

Private Sub ppi_Titanium_Click()
    Call ChangeSkinToTitanium
End Sub

Private Sub ppi_Toolbar_Click()
    pIndex = 6
    Call Showtutorials(pIndex)
End Sub

Private Sub ppi_TreasureChest_Click()
    Call ChangeSkinToTreasureChest
End Sub

Private Sub ppi_Wazoo_Click()
    Call ChangeSkinToWazoo
End Sub

Private Sub ppi_Weaponx_Click()
    Call ChangeSkinToWeaponx
End Sub

Private Sub ppi_Welcome_Click()
    Open App.Path & "\Welcome.txt" For Input As #1
        frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
    Close #1
End Sub

Private Sub ppi_WinXP_Click()
    Call ChangeSkinToWinXP
End Sub
