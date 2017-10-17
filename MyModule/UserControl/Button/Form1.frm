VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin UsrButton.iButton iButton1 
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Caption         =   "is"
      ClickedColor    =   -2147483648
      Level           =   2
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   14737632
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin UsrButton.iButton iButton2 
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Caption         =   "is"
      ClickedColor    =   8388608
      Level           =   2
      BackColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632319
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserControl11_Click()
    MsgBox "Gua di Click"
End Sub
