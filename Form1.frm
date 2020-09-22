VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Add thumbnail buttons"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set tooltip"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set clip RECT"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set overlay icon"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0013
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.HScrollBar hs_prg 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Image img2 
      Height          =   240
      Left            =   3480
      Picture         =   "Form1.frx":0062
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   2
      Left            =   4200
      Picture         =   "Form1.frx":03EC
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   3840
      Picture         =   "Form1.frx":04BE
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   0
      Left            =   3480
      Picture         =   "Form1.frx":0590
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress value:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress state:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents TaskBarList As ITaskBarList3
Private Const TBPF_NOPROGRESS = 0
Private Const TBPF_INDETERMINATE = 1
Private Const TBPF_NORMAL = 2
Private Const TBPF_ERROR = 4
Private Const TBPF_PAUSED = 8
Dim OverlayIcon As Long

Private Sub Command1_Click()
    TaskBarList.SetOverlayIcon hwnd, img2.Picture.Handle
End Sub

Private Sub Command2_Click()
    TaskBarList.SetThumbnailClip hwnd, 20, 20, 150, 150
End Sub

Private Sub Command3_Click()
    TaskBarList.SetThumbnailTooltip hwnd, InputBox("Input tooltip text:", "Thumbnail tooltip", "New tooltip, " & Now)
End Sub

Private Sub Command4_Click()
    Dim icons(6) As Long
    icons(0) = img(0).Picture.Handle
    icons(1) = img(1).Picture.Handle
    icons(2) = img(2).Picture.Handle
    TaskBarList.ThumbBarAddButtons hwnd, Val(InputBox("Input how many buttons to add:", "Thumbnail buttons", 3)), icons
End Sub

Private Sub Form_Load()
    Set TaskBarList = New ITaskBarList3
End Sub

Private Sub hs_prg_Change()
    TaskBarList.SetProgressValue hwnd, hs_prg.Value, hs_prg.Max
End Sub

Private Sub hs_prg_Scroll()
    Call hs_prg_Change
End Sub

Private Sub List1_Click()
    If List1.ListIndex > 0 Then TaskBarList.SetProgressState hwnd, 2 ^ (List1.ListIndex - 1) _
    Else TaskBarList.SetProgressState hwnd, 0
End Sub

Private Sub TaskBarList_ButtonPressed(ByVal Index As Integer)
    Caption = "Button " & Index & " is pressed"
End Sub
