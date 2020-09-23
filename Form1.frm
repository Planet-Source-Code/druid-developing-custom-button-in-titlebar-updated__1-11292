VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Custom button in Titlebar!"
   ClientHeight    =   1770
   ClientLeft      =   4680
   ClientTop       =   3675
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   4185
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   3120
      X2              =   3240
      Y1              =   820
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   2280
      X2              =   3240
      Y1              =   820
      Y2              =   0
   End
   Begin VB.Label lblDescription 
      Appearance      =   0  '2D
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "This is the custom button in the Titlebar. Click it and look what happens."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the procedure that is called when you click the button
Public Sub cmdInTitlebar_Click()
    MsgBox "Example created by Druid Developing", vbInformation, "About this program"
End Sub

Private Sub Form_Load()
    Call Init
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Terminate
End Sub
