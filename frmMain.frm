VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Controls Example"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Click Quit"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Click Dealer"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Name"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Caption"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"frmMain.frx":0000
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim Wnd As Long

Wnd& = cGetWindow("AfxFrameOrView")

cSetCaption (Wnd), ("Caption Changed!")

End Sub

Private Sub Command2_Click()

Dim Wnd As Long, Child As Long, Txt As Long

Wnd& = cGetWindow("AfxFrameOrView")
Child& = cGetWindow("#32770")
Txt& = FindWindowEx(Child&, 0, "Edit", vbNullString)

cSetText (Txt), ("TextChanged")

End Sub


Private Sub Command3_Click()

Dim Child As Long, Btn As Long, Btn1 As Long, Btn2 As Long

Child& = cGetWindow("#32770")
Btn& = FindWindowEx(Child&, 0, "Button", vbNullString)
Btn1& = FindWindowEx(Child&, Btn&, "Button", vbNullString)
Btn2& = FindWindowEx(Child&, Btn1&, "Button", vbNullString)

cClickButton (Btn2&)

End Sub


Private Sub Command4_Click()

Dim Child As Long, Btn As Long, Btn1 As Long, Btn2 As Long
Dim Btn3 As Long, Btn4 As Long

Child& = cGetWindow("#32770")
Btn& = FindWindowEx(Child&, 0, "Button", vbNullString)
Btn1& = FindWindowEx(Child&, Btn&, "Button", vbNullString)
Btn2& = FindWindowEx(Child&, Btn1&, "Button", vbNullString)
Btn3& = FindWindowEx(Child&, Btn2&, "Button", vbNullString)
Btn4& = FindWindowEx(Child&, Btn3&, "Button", vbNullString)

cClickButton (Btn4&)

End Sub


Private Sub lblInstructions_Click()

Shell "c:\windows\mshearts.exe", 1

End Sub


