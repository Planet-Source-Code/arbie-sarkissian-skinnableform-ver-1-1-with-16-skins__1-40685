Attribute VB_Name = "mdul_Main"
Option Explicit

Public pIndex As Integer

Sub Main()
    frm_Main.Show
End Sub

Public Sub Initialize()
    Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
    Call frm_Main.ctrl_PullDownMenu.AddItem("Start")
    Call frm_Main.ctrl_PullDownMenu.AddItem("View")
    Call frm_Main.ctrl_PullDownMenu.AddItem("Skins")
    Call frm_Main.ctrl_PullDownMenu.AddItem("Tutorials")
    Call frm_Main.ctrl_PullDownMenu.AddItem("Help")
        
    Call frm_Main.ctrl_ListObject.AddItem("Introduction")
    Call frm_Main.ctrl_ListObject.AddItem("Skinable Form")
    Call frm_Main.ctrl_ListObject.AddItem("Skinable Button")
    Call frm_Main.ctrl_ListObject.AddItem("Pulldown Menu")
    Call frm_Main.ctrl_ListObject.AddItem("Channel Bar")
    Call frm_Main.ctrl_ListObject.AddItem("List Object")
    Call frm_Main.ctrl_ListObject.AddItem("Toolbar Control")
    Call frm_Main.ctrl_ListObject.AddItem("Panel Control")
    Call frm_Main.ctrl_ListObject.AddItem("Coming Soon")
    
    Call frm_Main.ctrl_Panel.DrawPanel
    
    Call frm_Main.ctrl_Toolbar.DrawToolbar
    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
    Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
    Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
    Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
    Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
    Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Pulldown Menu")
    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
    Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Toolbar")
    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
    Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Statusbar")
    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
    Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
    Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
    Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
    
    Call frm_Main.ctrl_ChannelBar.AddItem("Start")
    Call frm_Main.ctrl_ChannelBar.AddItem("View")
    Call frm_Main.ctrl_ChannelBar.AddItem("Skins")
    Call frm_Main.ctrl_ChannelBar.AddItem("Tutorials")
    Call frm_Main.ctrl_ChannelBar.AddItem("Help")

    Open App.Path & "\Welcome.txt" For Input As #1
        frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
    Close #1
End Sub

Public Sub Showtutorials(m_Index As Integer)
    Select Case m_Index
        Case 0:
            Open App.Path & "\Introduction.txt" For Input As #1
                frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
            Close #1
        Case 1:
            Open App.Path & "\SkinableForm.txt" For Input As #1
                frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
            Close #1
        Case 2:
            Open App.Path & "\SkinableButton.txt" For Input As #1
                frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
            Close #1
        Case 3:
            Open App.Path & "\PulldownMenu.txt" For Input As #1
                frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
            Close #1
        Case 4:
            Open App.Path & "\ChannelBar.txt" For Input As #1
                frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
            Close #1
        Case 5:
            Open App.Path & "\ListObject.txt" For Input As #1
                frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
            Close #1
        Case 6:
            Open App.Path & "\Toolbar.txt" For Input As #1
                frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
            Close #1
        Case 7:
            Open App.Path & "\Panel.txt" For Input As #1
                frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
            Close #1
        Case 8:
            Open App.Path & "\ComingSoon.txt" For Input As #1
                frm_Main.tbx_Text.Text = Input$(LOF(1), #1)
            Close #1
    End Select
End Sub

Public Sub ChangeSkinToDefault()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Default"
        .ctrl_SkinableForm.BackColor = &HCECECE
        .ctrl_SkinableForm.CaptionTop = 360
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Default"
        .ctrl_btn_Previous.ForeColor = &H0&
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Default"
        .ctrl_btn_Next.ForeColor = &H0&
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Default"
        .ctrl_btn_Exit.ForeColor = &H0&
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\Default"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &HC0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\Default\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 60
        .ctrl_Toolbar.IconTop = 60
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Pulldown Menu")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Toolbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Statusbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Default"
        .ctrl_Toolbar.BackColor = &HCECECE
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\Default"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &HCECECE
        .ctrl_PullDownMenu.ForeColor = &H0&
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Default"
        .ctrl_ChannelBar.SubItemTop = 370
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &H0&
        .ctrl_ChannelBar.SubMouseMoveColor = &H0&
        .ctrl_ChannelBar.SubMouseDownColor = &H0&
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &H0&
        .lbl_Statusbar.ForeColor = &H0&
        
        .pic_Viewport.BackColor = &H0&
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H0&
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToTitanium()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Titanium"
        .ctrl_SkinableForm.BackColor = &H4B4A4A
        .ctrl_SkinableForm.CaptionTop = 195
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Titanium"
        .ctrl_btn_Previous.ForeColor = &HFFFFFF
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Titanium"
        .ctrl_btn_Next.ForeColor = &HFFFFFF
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Titanium"
        .ctrl_btn_Exit.ForeColor = &HFFFFFF
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\Titanium"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &H0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\Titanium\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 90
        .ctrl_Toolbar.IconTop = 90
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Titanium"
        .ctrl_Toolbar.BackColor = &H4B4A4A
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\Titanium"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &H4B4A4A
        .ctrl_PullDownMenu.ForeColor = &HFFFFFF
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Titanium"
        .ctrl_ChannelBar.SubItemTop = 395
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &HFFFFFF
        .lbl_Statusbar.ForeColor = &HFFFFFF
        
        .pic_Viewport.BackColor = &H0&
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H0&
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToBlue()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Blue"
        .ctrl_SkinableForm.BackColor = &HBD6E06
        .ctrl_SkinableForm.CaptionTop = 250
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Blue"
        .ctrl_btn_Previous.ForeColor = &H0&
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Blue"
        .ctrl_btn_Next.ForeColor = &H0&
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Blue"
        .ctrl_btn_Exit.ForeColor = &H0&
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\Blue"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &H0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\Blue\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 90
        .ctrl_Toolbar.IconTop = 90
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Blue"
        .ctrl_Toolbar.BackColor = &HBD6E06
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\Blue"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &HBD6E06
        .ctrl_PullDownMenu.ForeColor = &H0&
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Blue"
        .ctrl_ChannelBar.SubItemTop = 440
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &H0&
        .lbl_Statusbar.ForeColor = &H0&
        
        .pic_Viewport.BackColor = &H571B02
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H571B02
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToDeco()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_SkinableForm.BackColor = &HCECECE
        .ctrl_SkinableForm.CaptionTop = 300
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_btn_Previous.ForeColor = &H0&
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_btn_Next.ForeColor = &H0&
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_btn_Exit.ForeColor = &H0&
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &HC0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\Deco\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 60
        .ctrl_Toolbar.IconTop = 60
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Pulldown Menu")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Toolbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Statusbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_Toolbar.BackColor = &HCECECE
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &HCECECE
        .ctrl_PullDownMenu.ForeColor = &H0&
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_ChannelBar.SubItemTop = 400
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &H0&
        .ctrl_ChannelBar.SubMouseMoveColor = &H0&
        .ctrl_ChannelBar.SubMouseDownColor = &H0&
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &H0&
        .lbl_Statusbar.ForeColor = &H0&
        
        .pic_Viewport.BackColor = &H968A7B
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H968A7B
        .tbx_Text.ForeColor = &H0&
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToHolograph()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Holograph"
        .ctrl_SkinableForm.BackColor = &H3A5959
        .ctrl_SkinableForm.CaptionTop = 285
        .ctrl_SkinableForm.CaptionColor = &HFFFFFF
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Holograph"
        .ctrl_btn_Previous.ForeColor = &HFFFFFF
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Holograph"
        .ctrl_btn_Next.ForeColor = &HFFFFFF
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Holograph"
        .ctrl_btn_Exit.ForeColor = &HFFFFFF
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\Holograph"
        .ctrl_ListObject.MouseMoveColor = &HFFFFFF
        .ctrl_ListObject.MouseDownColor = &HFFFFFF
        .ctrl_ListObject.ForeColor = &HFFFFFF
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\Holograph\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 90
        .ctrl_Toolbar.IconTop = 90
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Holograph"
        .ctrl_Toolbar.BackColor = &H3A5959
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\Holograph"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &H3A5959
        .ctrl_PullDownMenu.ForeColor = &HFFFFFF
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Holograph"
        .ctrl_ChannelBar.SubItemTop = 395
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &HFFFFFF
        .lbl_Statusbar.ForeColor = &HFFFFFF
        
        .pic_Viewport.BackColor = &H263C3C
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H263C3C
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToTreasureChest()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\TreasureChest"
        .ctrl_SkinableForm.BackColor = &H0&
        .ctrl_SkinableForm.CaptionTop = 240
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\TreasureChest"
        .ctrl_btn_Previous.ForeColor = &HFFFFFF
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\TreasureChest"
        .ctrl_btn_Next.ForeColor = &HFFFFFF
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\TreasureChest"
        .ctrl_btn_Exit.ForeColor = &HFFFFFF
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\TreasureChest"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &H0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\TreasureChest\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.BackColor = &H0&
        .ctrl_Toolbar.IconLeft = 90
        .ctrl_Toolbar.IconTop = 90
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\TreasureChest"
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\TreasureChest"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &H0&
        .ctrl_PullDownMenu.ForeColor = &HFFFFFF
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\TreasureChest"
        .ctrl_ChannelBar.SubItemTop = 395
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &HFFFFFF
        .lbl_Statusbar.ForeColor = &HFFFFFF
        
        .pic_Viewport.BackColor = &H304B95
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H304B95
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToALPI()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\ALPI"
        .ctrl_SkinableForm.BackColor = &H2E2E32
        .ctrl_SkinableForm.CaptionTop = 135
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\ALPI"
        .ctrl_btn_Previous.ForeColor = &HFFFFFF
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\ALPI"
        .ctrl_btn_Next.ForeColor = &HFFFFFF
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\ALPI"
        .ctrl_btn_Exit.ForeColor = &HFFFFFF
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\ALPI"
        .ctrl_ListObject.ForeColor = &HFFFFFF
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &H0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\ALPI\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 90
        .ctrl_Toolbar.IconTop = 90
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\ALPI"
        .ctrl_Toolbar.BackColor = &H2E2E32
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\ALPI"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &H2E2E32
        .ctrl_PullDownMenu.ForeColor = &HFFFFFF
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\ALPI"
        .ctrl_ChannelBar.SubItemTop = 395
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &HFFFFFF
        .lbl_Statusbar.ForeColor = &HFFFFFF
        
        .pic_Viewport.BackColor = &H0&
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H0&
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToDoesnt_Suck()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Doesnt_Suck"
        .ctrl_SkinableForm.BackColor = &HCECECE
        .ctrl_SkinableForm.CaptionTop = 270
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Doesnt_Suck"
        .ctrl_btn_Previous.ForeColor = &H0&
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Doesnt_Suck"
        .ctrl_btn_Next.ForeColor = &H0&
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Doesnt_Suck"
        .ctrl_btn_Exit.ForeColor = &H0&
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\Doesnt_Suck"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &HC0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\Doesnt_Suck\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\Doesnt_Suck\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\Doesnt_Suck\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\Doesnt_Suck\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\Doesnt_Suck\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\Doesnt_Suck\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\Doesnt_Suck\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\Doesnt_Suck\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\Doesnt_Suck\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 60
        .ctrl_Toolbar.IconTop = 60
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Pulldown Menu")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Toolbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Statusbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Doesnt_Suck"
        .ctrl_Toolbar.BackColor = &HCECECE
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\Doesnt_Suck"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &HCECECE
        .ctrl_PullDownMenu.ForeColor = &H0&
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Doesnt_Suck"
        .ctrl_ChannelBar.SubItemTop = 370
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &H0&
        .ctrl_ChannelBar.SubMouseMoveColor = &H0&
        .ctrl_ChannelBar.SubMouseDownColor = &H0&
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &H0&
        .lbl_Statusbar.ForeColor = &H0&
        
        .pic_Viewport.BackColor = &H0&
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H0&
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToSteelBlade()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\SteelBlade"
        .ctrl_SkinableForm.BackColor = &H0&
        .ctrl_SkinableForm.CaptionTop = 405
        .ctrl_SkinableForm.CaptionColor = &HFFFFFF
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\SteelBlade"
        .ctrl_btn_Previous.ForeColor = &HFFFFFF
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\SteelBlade"
        .ctrl_btn_Next.ForeColor = &HFFFFFF
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\SteelBlade"
        .ctrl_btn_Exit.ForeColor = &HFFFFFF
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\SteelBlade"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &H0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\SteelBlade\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\SteelBlade\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\SteelBlade\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\SteelBlade\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\SteelBlade\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\SteelBlade\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\SteelBlade\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\SteelBlade\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\SteelBlade\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 90
        .ctrl_Toolbar.IconTop = 90
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\SteelBlade"
        .ctrl_Toolbar.BackColor = &H0&
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\SteelBlade"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &H0&
        .ctrl_PullDownMenu.ForeColor = &HFFFFFF
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\SteelBlade"
        .ctrl_ChannelBar.SubItemTop = 395
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &HFFFFFF
        .lbl_Statusbar.ForeColor = &HFFFFFF
        
        .pic_Viewport.BackColor = &H0&
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H0&
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "\Skins\SteelBlade", True)
    End With
End Sub

Public Sub ChangeSkinToWazoo()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Wazoo"
        .ctrl_SkinableForm.BackColor = &HE0E0E0
        .ctrl_SkinableForm.CaptionTop = 375
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Wazoo"
        .ctrl_btn_Previous.ForeColor = &H0&
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Wazoo"
        .ctrl_btn_Next.ForeColor = &H0&
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Wazoo"
        .ctrl_btn_Exit.ForeColor = &H0&
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\Wazoo"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &HC0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\Wazoo\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\Wazoo\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\Wazoo\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\Wazoo\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\Wazoo\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\Wazoo\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\Wazoo\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\Wazoo\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\Wazoo\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 60
        .ctrl_Toolbar.IconTop = 60
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Pulldown Menu")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Toolbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Statusbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Wazoo"
        .ctrl_Toolbar.BackColor = &HE0E0E0
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\Wazoo"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &HE0E0E0
        .ctrl_PullDownMenu.ForeColor = &H0&
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Wazoo"
        .ctrl_ChannelBar.SubItemTop = 370
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &H0&
        .ctrl_ChannelBar.SubMouseMoveColor = &H0&
        .ctrl_ChannelBar.SubMouseDownColor = &H0&
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &H0&
        .lbl_Statusbar.ForeColor = &H0&
        
        .pic_Viewport.BackColor = &HC6B3B3
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &HC6B3B3
        .tbx_Text.ForeColor = &H0&
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToSteelRain()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\SteelRain"
        .ctrl_SkinableForm.BackColor = &H40241B
        .ctrl_SkinableForm.CaptionTop = 250
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\SteelRain"
        .ctrl_btn_Previous.ForeColor = &HFFFFFF
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\SteelRain"
        .ctrl_btn_Next.ForeColor = &HFFFFFF
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\SteelRain"
        .ctrl_btn_Exit.ForeColor = &HFFFFFF
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\SteelRain"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &HFFFFFF
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\SteelRain\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\SteelRain\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\SteelRain\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\SteelRain\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\SteelRain\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\SteelRain\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\SteelRain\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\SteelRain\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\SteelRain\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 90
        .ctrl_Toolbar.IconTop = 90
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\SteelRain"
        .ctrl_Toolbar.BackColor = &H40241B
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\SteelRain"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &H40241B
        .ctrl_PullDownMenu.ForeColor = &HFFFFFF
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\SteelRain"
        .ctrl_ChannelBar.SubItemTop = 395
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &HFFFFFF
        .lbl_Statusbar.ForeColor = &HFFFFFF
        
        .pic_Viewport.BackColor = &H0&
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H0&
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToCoupe()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Coupe"
        .ctrl_SkinableForm.BackColor = &HABFEFF
        .ctrl_SkinableForm.CaptionTop = 180
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Coupe"
        .ctrl_btn_Previous.ForeColor = &H0&
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Coupe"
        .ctrl_btn_Next.ForeColor = &H0&
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Coupe"
        .ctrl_btn_Exit.ForeColor = &H0&
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\Coupe"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &HC0&
        .ctrl_ListObject.MouseDownColor = &HFFFFFF
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\Coupe\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\Coupe\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\Coupe\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\Coupe\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\Coupe\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\Coupe\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\Coupe\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\Coupe\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\Coupe\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 60
        .ctrl_Toolbar.IconTop = 60
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Pulldown Menu")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Toolbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Statusbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Coupe"
        .ctrl_Toolbar.BackColor = &HABFEFF
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\Coupe"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &HABFEFF
        .ctrl_PullDownMenu.ForeColor = &H0&
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Coupe"
        .ctrl_ChannelBar.SubItemTop = 370
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &H0&
        .ctrl_ChannelBar.SubMouseMoveColor = &H0&
        .ctrl_ChannelBar.SubMouseDownColor = &H0&
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &H0&
        .lbl_Statusbar.ForeColor = &H0&
        
        .pic_Viewport.BackColor = &HCCFF&
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &HCCFF&
        .tbx_Text.ForeColor = &H0&
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToBoilerRoom()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\BoilerRoom"
        .ctrl_SkinableForm.BackColor = &H111327
        .ctrl_SkinableForm.CaptionTop = 255
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\BoilerRoom"
        .ctrl_btn_Previous.ForeColor = &HFFFFFF
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\BoilerRoom"
        .ctrl_btn_Next.ForeColor = &HFFFFFF
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\BoilerRoom"
        .ctrl_btn_Exit.ForeColor = &HFFFFFF
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\BoilerRoom"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &H0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\BoilerRoom\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\BoilerRoom\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\BoilerRoom\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\BoilerRoom\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\BoilerRoom\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\BoilerRoom\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\BoilerRoom\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\BoilerRoom\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\BoilerRoom\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 90
        .ctrl_Toolbar.IconTop = 90
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Open")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Document")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Search")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\BoilerRoom"
        .ctrl_Toolbar.BackColor = &H111327
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\BoilerRoom"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &H111327
        .ctrl_PullDownMenu.ForeColor = &HFFFFFF
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\BoilerRoom"
        .ctrl_ChannelBar.SubItemTop = 395
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseMoveColor = &HFFFFFF
        .ctrl_ChannelBar.SubMouseDownColor = &HFFFFFF
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &HFFFFFF
        .lbl_Statusbar.ForeColor = &HFFFFFF
        
        .pic_Viewport.BackColor = &H0&
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H0&
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToExecutive()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Executive"
        .ctrl_SkinableForm.BackColor = &H6E8385
        .ctrl_SkinableForm.CaptionTop = 370
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Executive"
        .ctrl_btn_Previous.ForeColor = &H0&
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Executive"
        .ctrl_btn_Next.ForeColor = &H0&
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Executive"
        .ctrl_btn_Exit.ForeColor = &H0&
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\Executive"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &HC0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\Executive\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\Executive\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\Executive\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\Executive\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\Executive\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\Executive\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\Executive\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\Executive\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\Executive\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 60
        .ctrl_Toolbar.IconTop = 60
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Pulldown Menu")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Toolbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Statusbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Executive"
        .ctrl_Toolbar.BackColor = &H6E8385
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\Executive"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &H6E8385
        .ctrl_PullDownMenu.ForeColor = &H0&
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Executive"
        .ctrl_ChannelBar.SubItemTop = 370
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &H0&
        .ctrl_ChannelBar.SubMouseMoveColor = &H0&
        .ctrl_ChannelBar.SubMouseDownColor = &H0&
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &H0&
        .lbl_Statusbar.ForeColor = &H0&
        
        .pic_Viewport.BackColor = &H0&
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H0&
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToWeaponx()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Weaponx"
        .ctrl_SkinableForm.BackColor = &HCECECE
        .ctrl_SkinableForm.CaptionTop = 135
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\Weaponx"
        .ctrl_btn_Previous.ForeColor = &H0&
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\Weaponx"
        .ctrl_btn_Next.ForeColor = &H0&
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\Weaponx"
        .ctrl_btn_Exit.ForeColor = &H0&
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\Weaponx"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &HC0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\Weaponx\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\Weaponx\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\Weaponx\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\Weaponx\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\Weaponx\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\Weaponx\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\Weaponx\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\Weaponx\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\Weaponx\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 60
        .ctrl_Toolbar.IconTop = 60
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Pulldown Menu")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Toolbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Statusbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\Weaponx"
        .ctrl_Toolbar.BackColor = &HCECECE
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\Weaponx"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &HCECECE
        .ctrl_PullDownMenu.ForeColor = &H0&
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\Weaponx"
        .ctrl_ChannelBar.SubItemTop = 370
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &H0&
        .ctrl_ChannelBar.SubMouseMoveColor = &H0&
        .ctrl_ChannelBar.SubMouseDownColor = &H0&
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &H0&
        .lbl_Statusbar.ForeColor = &H0&
        
        .pic_Viewport.BackColor = &H0&
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H0&
        .tbx_Text.ForeColor = &HFFFFFF
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "", False)
    End With
End Sub

Public Sub ChangeSkinToWinXP()
    With frm_Main
        .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\WinXP"
        .ctrl_SkinableForm.BackColor = &H8000000F
        .ctrl_SkinableForm.CaptionTop = 120
        .ctrl_SkinableForm.CaptionColor = &HFFFFFF
        Call frm_Main.ctrl_SkinableForm.LoadSkin(frm_Main)
        
        .ctrl_btn_Previous.SkinPath = App.Path & "\Skins\WinXP"
        .ctrl_btn_Previous.ForeColor = &H0&
        .ctrl_btn_Previous.LoadSkin
        .ctrl_btn_Previous.Refresh
        .ctrl_btn_Next.SkinPath = App.Path & "\Skins\WinXP"
        .ctrl_btn_Next.ForeColor = &H0&
        .ctrl_btn_Next.LoadSkin
        .ctrl_btn_Next.Refresh
        .ctrl_btn_Exit.SkinPath = App.Path & "\Skins\WinXP"
        .ctrl_btn_Exit.ForeColor = &H0&
        .ctrl_btn_Exit.LoadSkin
        .ctrl_btn_Exit.Refresh
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\WinXP"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &HC0&
        .ctrl_ListObject.MouseDownColor = &H0&
        .iml_Toolbar.ListImages.Clear
        .iml_Toolbar.ListImages.Add 1, , LoadPicture(App.Path & "\Skins\WinXP\Toolbar Icons\icn_Back.gif")
        .iml_Toolbar.ListImages.Add 2, , LoadPicture(App.Path & "\Skins\WinXP\Toolbar Icons\icn_Forward.gif")
        .iml_Toolbar.ListImages.Add 3, , LoadPicture(App.Path & "\Skins\WinXP\Toolbar Icons\icn_Home.gif")
        .iml_Toolbar.ListImages.Add 4, , LoadPicture(App.Path & "\Skins\WinXP\Toolbar Icons\icn_Refresh.gif")
        .iml_Toolbar.ListImages.Add 5, , LoadPicture(App.Path & "\Skins\WinXP\Toolbar Icons\icn_Open.gif")
        .iml_Toolbar.ListImages.Add 6, , LoadPicture(App.Path & "\Skins\WinXP\Toolbar Icons\icn_Document.gif")
        .iml_Toolbar.ListImages.Add 7, , LoadPicture(App.Path & "\Skins\WinXP\Toolbar Icons\icn_Search.gif")
        .iml_Toolbar.ListImages.Add 8, , LoadPicture(App.Path & "\Skins\WinXP\Toolbar Icons\icn_Help.gif")
        .iml_Toolbar.ListImages.Add 9, , LoadPicture(App.Path & "\Skins\WinXP\Toolbar Icons\icn_Stop.gif")
        .ctrl_Toolbar.UnloadButtons
        .ctrl_Toolbar.IconLeft = 60
        .ctrl_Toolbar.IconTop = 60
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(1).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(0, "Back")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(2).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(1, "Forward")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(3).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(2, "Home")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(4).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(3, "Refresh")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(5).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(4, "Pulldown Menu")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(6).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(5, "Toolbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(7).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(6, "Statusbar")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(8).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(7, "Help")
        Call frm_Main.ctrl_Toolbar.AddButton(frm_Main.iml_Toolbar.ListImages(9).Picture)
        Call frm_Main.ctrl_Toolbar.AddTooltipText(8, "Exit")
        .ctrl_ListObject.DrawMenu
        
        .ctrl_Toolbar.SkinPath = App.Path & "\Skins\WinXP"
        .ctrl_Toolbar.BackColor = &H8000000F
        .ctrl_Toolbar.DrawToolbar
        .ctrl_Toolbar.Refresh
        
        .ctrl_Panel.SkinPath = App.Path & "\Skins\WinXP"
        .ctrl_Panel.DrawPanel
        
        .ctrl_PullDownMenu.BackColor = &H8000000F
        .ctrl_PullDownMenu.ForeColor = &H0&
        .ctrl_PullDownMenu.Refresh
        
        .ctrl_ChannelBar.SkinPath = App.Path & "\Skins\WinXP"
        .ctrl_ChannelBar.SubItemTop = 430
        .ctrl_ChannelBar.MouseMoveColor = &H0&
        .ctrl_ChannelBar.MouseDownColor = &H0&
        .ctrl_ChannelBar.SubMouseMoveColor = &H0&
        .ctrl_ChannelBar.SubMouseDownColor = &H0&
        .ctrl_ChannelBar.DrawMenu
        
        .Line1.BorderColor = &H0&
        .lbl_Statusbar.ForeColor = &H0&
        
        .pic_Viewport.BackColor = &H8000000F
        .pic_Viewport.Refresh
        .tbx_Text.BackColor = &H8000000F
        .tbx_Text.ForeColor = &H0&
        
        Call .ctrl_TransparetForm.ShapeForm(frm_Main, App.Path & "\Skins\WinXP", True)
    End With
End Sub

