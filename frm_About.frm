VERSION 5.00
Begin VB.Form frm_About 
   BorderStyle     =   0  'None
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   900
      ScaleHeight     =   1155
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   960
      Width           =   2355
      Begin VB.Label lbl_About 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Skinable Form version 1.1 copyright(c) 2002 by Arbie Sarkissian"
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
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin SkinableForm.ctrl_SkinableButton ctrl_btn_OK 
      Height          =   375
      Left            =   1500
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "OK"
   End
   Begin SkinableForm.ctrl_Panel ctrl_Panel 
      Height          =   1380
      Left            =   780
      TabIndex        =   1
      Top             =   840
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2434
   End
   Begin SkinableForm.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      MaximizeBtn     =   0   'False
      MinimizeBtn     =   0   'False
      Caption         =   "About..."
      BackColor       =   4934218
   End
End
Attribute VB_Name = "frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctrl_btn_OK_Click()
    Unload frm_About
End Sub

Private Sub Form_Load()
    Call frm_About.ctrl_SkinableForm.LoadSkin(frm_About)
    Call frm_About.ctrl_Panel.DrawPanel
    Call frm_About.ctrl_btn_OK.LoadSkin
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    frm_About.ctrl_btn_OK.Refresh
End Sub
