VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   Begin Project1.iToolbar iToolbar1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _extentx        =   8070
      _extenty        =   873
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    iToolbar1.AuthorizationKey 0, 2, 2, 2, 2, 2, 2, 2, 2, 2
End Sub

