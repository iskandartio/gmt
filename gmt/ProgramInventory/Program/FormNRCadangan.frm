VERSION 5.00
Object = "{8409C9A2-E167-42E0-9B20-105E4009A852}#1.0#0"; "USRLINE.OCX"
Object = "{5B6E0E90-AB64-4D5D-AC5E-5DC35FA1D835}#1.0#0"; "USRTEXT.OCX"
Object = "{3EF9EBB3-E2A3-407F-997F-07A6E8023C5D}#1.0#0"; "USRLABEL.OCX"
Object = "{929714A7-8741-466B-B8ED-064307B9D0CA}#1.0#0"; "USRFRAME.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{4081DB6F-5C42-11DA-99A5-000B6A30ACAC}#1.0#0"; "USRTDBGRID.OCX"
Begin VB.Form FormNR 
   Caption         =   "RETUR"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin UsrFrame.IFrame FramePreview 
      Height          =   15510
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   27358
      BackColor       =   16777215
      Begin UsrLabel.ILabel ILabel20 
         Height          =   255
         Left            =   840
         TabIndex        =   75
         Top             =   14040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         AutoWidth       =   0   'False
         Caption         =   "BINGRIANTO"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoRedraw      =   0   'False
      End
      Begin UsrLabel.ILabel ILabel19 
         Height          =   255
         Left            =   720
         TabIndex        =   74
         Top             =   13080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         AutoWidth       =   0   'False
         Caption         =   "TANGERANG"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoRedraw      =   0   'False
      End
      Begin UsrLine.ILine ILine5 
         Height          =   30
         Left            =   720
         TabIndex        =   73
         Top             =   13920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   53
         BackColor       =   16777215
      End
      Begin UsrLabel.ILabel ILabel18 
         Height          =   255
         Left            =   2880
         TabIndex        =   72
         Top             =   13080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         AutoWidth       =   0   'False
         Caption         =   "DITERIMA OLEH"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoRedraw      =   0   'False
      End
      Begin UsrLine.ILine ILine6 
         Height          =   30
         Left            =   2760
         TabIndex        =   71
         Top             =   13920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   53
         BackColor       =   16777215
      End
      Begin UsrLabel.ILabel PTOTALGR 
         Height          =   210
         Left            =   8400
         TabIndex        =   65
         Top             =   3960
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   370
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "99.999,00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PKGDISCGR 
         Height          =   210
         Left            =   6600
         TabIndex        =   64
         Top             =   3960
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   370
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "5000"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PKGRETURGR 
         Height          =   210
         Left            =   5760
         TabIndex        =   63
         Tag             =   "3960"
         Top             =   3960
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   370
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "5000"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PLGRAND 
         Height          =   210
         Left            =   4560
         TabIndex        =   62
         Top             =   3960
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   370
         Caption         =   "GRAND TOTAL"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PTOTALSUB 
         Height          =   210
         Index           =   0
         Left            =   9600
         TabIndex        =   61
         Top             =   3600
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   370
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "99.999,00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PKGDISCSUB 
         Height          =   210
         Index           =   0
         Left            =   7440
         TabIndex        =   60
         Top             =   3600
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   370
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "5000"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PKGRETURSUB 
         Height          =   210
         Index           =   0
         Left            =   5760
         TabIndex        =   59
         Top             =   3600
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   370
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "5000"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PLTOTAL 
         Height          =   210
         Index           =   0
         Left            =   5160
         TabIndex        =   58
         Top             =   3600
         Visible         =   0   'False
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   370
         Caption         =   "TOTAL"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PCURR 
         Height          =   210
         Left            =   9360
         TabIndex        =   57
         Top             =   1920
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   370
         Caption         =   "@CURR"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PTANGGALKW 
         Height          =   210
         Left            =   9360
         TabIndex        =   56
         Top             =   1680
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   370
         Caption         =   "@TANGGALKW"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PNOKW 
         Height          =   210
         Left            =   9360
         TabIndex        =   55
         Top             =   1440
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   370
         Caption         =   "@NOKW"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PTANGGAL 
         Height          =   210
         Left            =   9360
         TabIndex        =   54
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Caption         =   "@TANGGAL"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PNO 
         Height          =   210
         Left            =   9360
         TabIndex        =   53
         Top             =   960
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   370
         AutoWidth       =   0   'False
         Caption         =   "@NO"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PCUSTOMER 
         Height          =   210
         Left            =   480
         TabIndex        =   52
         Top             =   1200
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         Caption         =   "@CUSTOMER"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PTOTAL 
         Height          =   210
         Index           =   0
         Left            =   9600
         TabIndex        =   51
         Top             =   3360
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   370
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "999.999,00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PDISCOUNT 
         Height          =   210
         Index           =   0
         Left            =   8400
         TabIndex        =   50
         Top             =   3360
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   370
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "88.888.00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PKGDISC 
         Height          =   210
         Index           =   0
         Left            =   7440
         TabIndex        =   49
         Top             =   3360
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   370
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "5000,00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PHARGARETUR 
         Height          =   210
         Index           =   0
         Left            =   6480
         TabIndex        =   48
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "880.000,00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PKGRETUR 
         Height          =   210
         Index           =   0
         Left            =   5760
         TabIndex        =   47
         Top             =   3360
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   370
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "5000,00"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PTANGGALSJ 
         Height          =   210
         Index           =   0
         Left            =   5280
         TabIndex        =   46
         Top             =   3120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   370
         Caption         =   "@TANGGAL SJ"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PJENIS 
         Height          =   210
         Index           =   0
         Left            =   1200
         TabIndex        =   45
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
         Caption         =   "@JENIS"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PNOSJ 
         Height          =   210
         Index           =   0
         Left            =   1080
         TabIndex        =   44
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
         Caption         =   "@NOSJ"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PNOURUT 
         Height          =   210
         Index           =   0
         Left            =   480
         TabIndex        =   43
         Top             =   3360
         Visible         =   0   'False
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   370
         Alignment       =   1
         AutoWidth       =   0   'False
         Caption         =   "@NO"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PLTANGGALSJ 
         Height          =   210
         Index           =   0
         Left            =   4080
         TabIndex        =   42
         Top             =   3120
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   370
         Caption         =   "TANGGAL SJ"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PLNOSJ 
         Height          =   210
         Index           =   0
         Left            =   480
         TabIndex        =   41
         Top             =   3120
         Visible         =   0   'False
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   370
         Alignment       =   1
         AutoWidth       =   0   'False
         Caption         =   "NO SJ"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel17 
         Height          =   210
         Left            =   10440
         TabIndex        =   40
         Top             =   2640
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   370
         Caption         =   "TOTAL"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel16 
         Height          =   210
         Left            =   8400
         TabIndex        =   39
         Top             =   2760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   "DISCOUNT"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel15 
         Height          =   210
         Left            =   7800
         TabIndex        =   38
         Top             =   2760
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   370
         Caption         =   "KG"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel14 
         Height          =   210
         Left            =   6720
         TabIndex        =   37
         Top             =   2760
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
         Caption         =   "HARGA"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel13 
         Height          =   210
         Left            =   6120
         TabIndex        =   36
         Top             =   2760
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   370
         Caption         =   "KG"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel12 
         Height          =   210
         Left            =   8160
         TabIndex        =   35
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   "DISCOUNT"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel11 
         Height          =   210
         Left            =   6360
         TabIndex        =   34
         Top             =   2520
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   370
         Caption         =   "RETUR"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel10 
         Height          =   210
         Left            =   2640
         TabIndex        =   33
         Top             =   2640
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   370
         Caption         =   "JENIS BARANG"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel9 
         Height          =   210
         Left            =   480
         TabIndex        =   32
         Top             =   2640
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   370
         Alignment       =   1
         AutoWidth       =   0   'False
         Caption         =   "NO"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel8 
         Height          =   210
         Left            =   7560
         TabIndex        =   31
         Top             =   1920
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   370
         Caption         =   "MATA UANG"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel7 
         Height          =   210
         Left            =   7560
         TabIndex        =   30
         Top             =   1680
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   370
         Caption         =   "TANGGAL"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel6 
         Height          =   210
         Left            =   7560
         TabIndex        =   29
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Caption         =   "NO KWITANSI"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel5 
         Height          =   210
         Left            =   7560
         TabIndex        =   28
         Top             =   1200
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   370
         Caption         =   "TANGGAL"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel4 
         Height          =   210
         Left            =   7560
         TabIndex        =   27
         Top             =   960
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Caption         =   "NO RETUR"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel3 
         Height          =   210
         Left            =   480
         TabIndex        =   26
         Top             =   960
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   370
         Caption         =   "NAMA DAN ALAMAT CUSTOMER"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel2 
         Height          =   210
         Left            =   9840
         TabIndex        =   25
         Top             =   480
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   370
         Caption         =   "NOTA RETUR"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel ILabel1 
         Height          =   330
         Left            =   480
         TabIndex        =   24
         Top             =   480
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   582
         Caption         =   "PT GEMILANG MAJU TEXINDOTAMA"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PALAMAT 
         Height          =   210
         Left            =   480
         TabIndex        =   23
         Top             =   1440
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   370
         AutoHeight      =   -1  'True
         AutoWidth       =   0   'False
         Caption         =   "@ALAMAT"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLabel.ILabel PTERBILANG 
         Height          =   240
         Left            =   480
         TabIndex        =   22
         Tag             =   "4200"
         Top             =   4200
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   423
         Alignment       =   2
         AutoWidth       =   0   'False
         Caption         =   "@TERBILANG"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UsrLine.ILine ILine1 
         Height          =   150
         Left            =   480
         TabIndex        =   21
         Top             =   2400
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   265
         BackColor       =   16777215
      End
      Begin UsrLine.ILine ILine2 
         Height          =   150
         Left            =   480
         TabIndex        =   20
         Top             =   3000
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   265
         BackColor       =   16777215
      End
      Begin UsrLine.ILine PLINE 
         Height          =   150
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   3840
         Visible         =   0   'False
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   265
         BackColor       =   16777215
      End
   End
   Begin VB.CommandButton fList 
      Caption         =   "&LIST"
      Height          =   375
      Left            =   10560
      TabIndex        =   70
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT2"
      Height          =   375
      Left            =   9660
      TabIndex        =   69
      Top             =   240
      Width           =   795
   End
   Begin VB.TextBox fMataUang 
      Height          =   285
      Left            =   6240
      TabIndex        =   68
      Top             =   360
      Width           =   855
   End
   Begin VB.CheckBox fPreview 
      Caption         =   "Pre&view"
      Height          =   375
      Left            =   8640
      TabIndex        =   66
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton fNew 
      Caption         =   "&NEW"
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton fPrint 
      Caption         =   "&PRINT"
      Height          =   375
      Left            =   7440
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox fQuick 
      Height          =   315
      Left            =   3600
      TabIndex        =   15
      Top             =   960
      Width           =   1815
   End
   Begin TrueOleDBGrid70.TDBDropDown TDBDropDown1 
      Height          =   2655
      Left            =   1080
      TabIndex        =   13
      Top             =   2520
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4683
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO SJ"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "JENIS BARANG"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "HARGA"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "JUMLAH"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "nosc"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "idsc"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "idstock"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(0)._MinWidth=12632256"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=8361"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=8281"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1799"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1720"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(2)._MinWidth=221"
      Splits(0)._ColumnProps(21)=   "Column(3).Width=1535"
      Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=1455"
      Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(27)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(31)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(33)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(35)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(38)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(41)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(43)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(46)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(48)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(49)=   "Column(6).Order=7"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   2
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   ""
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.alignment=2"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=31"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=125,.parent=2,.namedParent=127"
      _StyleDefs(17)  =   "FilterBarStyle:id=128,.parent=1,.namedParent=130"
      _StyleDefs(18)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=126,.parent=125"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=129,.parent=128"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=76,.parent=11"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=73,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=74,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=75,.parent=15"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=24,.parent=11"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=21,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=22,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=23,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=72,.parent=11"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=69,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=70,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=71,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=92,.parent=11"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=89,.parent=12"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=90,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=91,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=108,.parent=11"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=105,.parent=12"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=106,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=107,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=116,.parent=11"
      _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=113,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=114,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=115,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=124,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=121,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=122,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=123,.parent=15"
      _StyleDefs(58)  =   "Named:id=29:Normal"
      _StyleDefs(59)  =   ":id=29,.parent=0"
      _StyleDefs(60)  =   "Named:id=30:Heading"
      _StyleDefs(61)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   ":id=30,.wraptext=-1"
      _StyleDefs(63)  =   "Named:id=31:Footing"
      _StyleDefs(64)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   "Named:id=32:Selected"
      _StyleDefs(66)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=33:Caption"
      _StyleDefs(68)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(69)  =   "Named:id=34:HighlightRow"
      _StyleDefs(70)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(71)  =   "Named:id=35:EvenRow"
      _StyleDefs(72)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(73)  =   "Named:id=36:OddRow"
      _StyleDefs(74)  =   ":id=36,.parent=29"
      _StyleDefs(75)  =   "Named:id=127:RecordSelector"
      _StyleDefs(76)  =   ":id=127,.parent=30"
      _StyleDefs(77)  =   "Named:id=130:FilterBar"
      _StyleDefs(78)  =   ":id=130,.parent=29"
   End
   Begin UsrTDBGrid.ITDBGrid Grid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4683
      Row             =   -1
      HeadLines       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton fDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton fSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin UsrText.IText fTanggal 
      Height          =   270
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   476
      Text            =   "__/__/__"
      DataType        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
   End
   Begin UsrText.IText fNo 
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText fNoKW 
      Height          =   270
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText fCustomer 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UsrText.IText fTanggalKW 
      Height          =   270
      Left            =   4920
      TabIndex        =   12
      Top             =   360
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   476
      Text            =   "__/__/__"
      DataType        =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
   End
   Begin VB.Label Label7 
      Caption         =   "MATA UANG"
      Height          =   255
      Left            =   6240
      TabIndex        =   67
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "TANGGAL KW"
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "LIHAT RETUR LAIN"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "TANGGAL NR"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "NO RETUR"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "NO KW"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "NAMA CUSTOMER"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FormNR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim z As New XArrayDB
Dim fAlamat As String
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub GoEvent(ByVal tEvent As String)
    If tEvent = "ADD" Then
        m_mode = "ADD"
        tProp = 2
    ElseIf tEvent = "SEE" Then
        m_mode = "SEE"
        tProp = 1
    End If
    v = IIf((tProp And 1) = 0, False, True)
        fDelete.Enabled = v
        fPrint.Enabled = v
    v = IIf((tProp And 2) = 0, False, True)
        fNo.Enabled = v
        fTanggal.Enabled = v
        fSave.Enabled = v
        Grid1.AllowUpdate = v
        fMataUang.Enabled = v
End Sub

Private Sub Command1_Click()
    FormReport.LoadMe "NR" & pTipe & ".rpt", fNo
End Sub

Private Sub fList_Click()
    FormList.LoadMe "NOTA RETUR", _
"select distinct NoNR, TanggalNR, t_NR" & pTipe & ".NoKW, TanggalKW, m_customer.Nama, t_NR" & pTipe & ".Total from (t_NR" & pTipe & " left join t_SPP" & pTipe & " on t_NR" & pTipe & ".NoKW=t_SPP" & pTipe & ".NoKW) left join m_customer on m_customer.Kode=t_SPP" & pTipe & ".Kode where 1=1", _
"Nama Customer@No NR@Tanggal NR@No KW@Tanggal KW", _
"m_Customer.Nama@NoNR@TanggalNR@t_NR" & pTipe & ".NoKW@TanggalKW", _
"2500@1000@1000@1000@1000", _
"String@String@Date@String@Date", _
"NO NR@TANGGAL NR@NO KW@TANGGAL KW@NAMA CUSTOMER@TOTAL", _
"1700@1000@1700@1000@2500@1500", "String@Date@String@Date@String@Decimal", Me, _
" order by NoNR"
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Grid1.Height = ScaleHeight - Grid1.Top - 200
    Grid1.Width = ScaleWidth - 2 * Grid1.Left
End Sub

Private Sub fPreview_Click()
    If fPreview Then
        FramePreview.ZOrder
        FramePreview.Visible = True
        ProsesPreview
    Else
        FramePreview.Visible = False
    End If
End Sub

Private Sub ProsesPreview()
On Error Resume Next
    Grid1.Update
    Dim GrandTotal As Double
    Dim GrandKgRetur As Double
    Dim GrandKgDisc As Double
    PNO = fNo
    PNO.Paint
    PTANGGAL = fTanggal
    PNOKW = fNoKW
    PNOKW.Paint
    PTANGGALKW = fTanggalKW
    PCURR = fMataUang
    PCUSTOMER = fCustomer
    PALAMAT = fAlamat
    PTANGGAL = fTanggal
    t = ILine2.Top
    c1 = 0
    c2 = 0
    GrandTotal = 0
    GrandKgRetur = 0
    GrandKgDisc = 0
    For i = 1 To PLINE.Count - 1
        Unload PLINE(i)
    Next
    Dim c() As Boolean
    ReDim c(Grid1.RowCount - 1)
    '0NO SJ.1JENIS BARANG.2(BOX).3(KG).4HARGA.5DISC (KG).6DISCOUNT.*NoSC.*IdSC.*IdStock
    For i = 0 To Grid1.RowCount - 1
        If Not c(i) Then
            c(i) = True
            t = t + PNOSJ(0).Top - ILine2.Top
            c1 = c1 + 1
            LoadCaption PNOSJ(), Grid1(i, "NO SJ"), c1, t
            LoadCaption PTANGGALSJ(), cTanggal(Grid1(i, "TANGGAL SJ")), c1, t
            t = t + PNOURUT(0).Top - PNOSJ(0).Top
            Dim SubJumlahRetur As Double
            Dim SubJumlahDiscount As Double
            Dim TotalHarga As Double
            SubJumlahRetur = 0
            SubJumlahDiscount = 0
            SubTotalHarga = 0
            k = 0
            For j = i To Grid1.RowCount - 1
                If Grid1(j, 0) = Grid1(i, 0) Then
                    c(j) = True
                    c2 = c2 + 1
                    k = k + 1
                    LoadCaption PNOURUT(), k, c2, t
                    LoadCaption PJENIS(), Grid1(j, "JENIS BARANG"), c2, t
                    kgretur = Grid1(j, "(KG)")
                    LoadCaption PKGRETUR(), cDecimal(kgretur), c2, t
                    hargaretur = Grid1(j, "HARGA")
                    LoadCaption PHARGARETUR(), cDecimal(hargaretur), c2, t
                    kgdisc = Grid1(j, "DISC (KG)")
                    LoadCaption PKGDISC(), cDecimal(kgdisc), c2, t
                    disc = Grid1(j, "DISCOUNT")
                    LoadCaption PDISCOUNT(), cDecimal(disc), c2, t
                    total = kgretur * hargaretur + kgdisc * disc
                    LoadCaption PTOTAL(), cDecimal(total), c2, t
                    SubJumlahRetur = SubJumlahRetur + kgretur
                    SubJumlahDiscount = SubJumlahDiscount + kgdisc
                    SubTotalHarga = SubTotalHarga + total
                    GrandKgRetur = GrandKgRetur + kgretur
                    GrandKgDisc = GrandKgDisc + kgdisc
                    GrandTotal = GrandTotal + total
                    t = t + PNOURUT(0).Height
                End If
            Next
            t = t + PKGRETURSUB(0).Top - PNOURUT(0).Top - PNOURUT(0).Height
            LoadCaption PKGRETURSUB(), cDecimal(SubJumlahRetur), c1, t
            LoadCaption PKGDISCSUB(), cDecimal(SubJumlahDiscount), c1, t
            LoadCaption PTOTALSUB(), cDecimal(SubTotalHarga), c1, t
            LoadCaption PLTOTAL(), "TOTAL", c1, t
            t = t + PLINE(0).Top - PKGRETURSUB(0).Top
            Load PLINE(c1)
            PLINE(c1).Visible = True
            PLINE(c1).Top = t
        End If
    Next
    t = t + PKGRETURGR.Tag - PLINE(0).Top
    PKGRETURGR = cDecimal(GrandKgRetur)
    PKGDISCGR = cDecimal(GrandKgDisc)
    PTOTALGR = cDecimal(GrandTotal)
    PTERBILANG.Top = t + PTERBILANG.Tag - PKGRETURGR.Tag
    PLGRAND.Top = t
    PKGRETURGR.Top = t
    PKGDISCGR.Top = t
    PTOTALGR.Top = t
    PTERBILANG = Terbilang(Round(GrandTotal, 2), fMataUang)
    For i = c1 + 1 To PNOSJ.Count - 1
        PNOSJ(i).Visible = False
        PTANGGALSJ(i).Visible = False
        PKGRETURSUB(i).Visible = False
        PKGDISCSUB(i).Visible = False
        PLTOTAL(i).Visible = False
    Next
    For i = c2 + 1 To PNOURUT.Count - 1
        PNOURUT(i).Visible = False
        PJENIS(i).Visible = False
        PKGRETUR(i).Visible = False
        PHARGARETUR(i).Visible = False
        PKGDISC(i).Visible = False
        PDISCOUNT(i).Visible = False
        PTOTAL(i).Visible = False
    Next
End Sub

Private Sub UpdatefQuick()
    temp(0) = fQuick
    a = "select NoNR from t_NR" & pTipe & " where NoKW='" & fNoKW & "'"
    query a
    fQuick.Clear
    If RS.RecordCount > 0 Then
        For i = 0 To RS.RecordCount - 1
            fQuick.List(i) = RS.Fields(0).Value
            RS.MoveNext
        Next
    End If
    fQuick = temp(0)
End Sub

Sub GetResult(ByVal tNo As String)
    If tNo = "" Then Exit Sub
    fQuick = Left(tNo, 5) & "/" & Right(tNo, 2)
    a = "select Nama, t_NR" & pTipe & ".NoKW, TanggalKW, NoNR, TanggalNR, MataUang, StatusNR " & _
        "from (t_NR" & pTipe & " left join t_SPP" & pTipe & " on t_NR" & pTipe & ".NoKW=t_SPP" & pTipe & ".NoKW) left join m_customer on t_SPP" & pTipe & ".Kode=m_customer.Kode where ShortNR='" & fQuick & "'"
    query a
    GoEvent "SEE"
    If RS.RecordCount = 0 Then Exit Sub
    If RS.Fields("StatusNR").Value = 1 Then fDelete.Enabled = False
    fCustomer = RS.Fields("Nama").Value
    fNo = RS.Fields("NoNR").Value
    fTanggal = cTanggal(RS.Fields("TanggalNR").Value)
    fNoKW = RS.Fields("NoKW").Value
    fTanggalKW = cTanggal(RS.Fields("TanggalKW").Value)
    fMataUang = RS.Fields("MataUang").Value
    UpdatefQuick
    a = "select t_NRDetail" & pTipe & ".NoSJ,Jenis&' '&KodeBarang&' '&Warna&' '&NoWarna&' '&Tube&' '&Grade,ReturBox,ReturKg," & _
        "t_NRDetail" & pTipe & ".Harga,DiscKg,Discount,NoSC,IdSC,t_NRDetail" & pTipe & ".IdStock from " & _
        "((t_NRDetail" & pTipe & " left join t_SPPDetail" & pTipe & " on t_NRDetail" & pTipe & ".NoSJ=t_SPPDetail" & pTipe & ".NoSJ and t_NRDetail" & pTipe & ".IdStock=t_SPPDetail" & pTipe & ".IdStock) " & _
        "left join t_SPP" & pTipe & " on t_SPP" & pTipe & ".NoSPP=t_SPPDetail" & pTipe & ".NoSPP) left join m_stock" & pTipe & " on m_stock" & pTipe & ".IdStock=t_NRDetail" & pTipe & ".IdStock where NoNR='" & fNo & "'"
    query a
    Dim rs1() As Variant
    rs1 = RS.GetRows
    Grid1.SetDB rs1
    fPreview.Value = 0
End Sub

Sub LoadMe(ByVal tNo As String)
    ClearScreen
    Set TDBDropDown1.Array = z
    fCustomer.Enabled = False
    fNoKW.Enabled = False
    fTanggalKW.Enabled = False
    Grid1.SetHeader "NO SJ.JENIS BARANG.(BOX).(KG).HARGA.DISC (KG).DISCOUNT.*NoSC.*IdSC.*IdStock"
    Grid1.SetWidth "1700.3000.500.700.1000.700.1000"
    Grid1.SetProp "JENIS BARANG.HARGA", "Locked", True
    Grid1.SetType "String.String.Integer.Decimal.Decimal.Decimal.Decimal"
    Grid1.Columns("NO SJ").AutoDropDown = True
    Grid1.Columns("NO SJ").DropDown = TDBDropDown1
    a = "select top 1 Nama, Alamat,NoKW, TanggalKW, MataUang from t_SPP" & pTipe & " left join m_customer on t_SPP" & pTipe & ".Kode=m_customer.Kode where NoKW='" & tNo & "' order by Nama"
    query a
    If RS.RecordCount = 0 Then Exit Sub
    fCustomer = RS.Fields(0).Value
    fAlamat = RS.Fields(1).Value
    fNoKW = RS.Fields(2).Value
    fTanggalKW = cTanggal(RS.Fields(3).Value)
    fMataUang = RS.Fields("MataUang").Value
    fNo_LostFocus
    a = "select t_SPPDetail" & pTipe & ".NoSJ,Jenis+' '+KodeBarang+' '+Warna+' '+NoWarna&' '+Tube+' '+Grade,Harga,t_SPPDetail" & pTipe & ".JumlahKg,NoSC,IdSC,m_stock" & pTipe & ".IdStock from t_SPPDetail" & pTipe & " left join m_stock" & pTipe & " on t_SPPDetail" & pTipe & ".IdStock=m_stock" & pTipe & ".IdStock where NoKWDetail='" & tNo & "'"
    query a
    z.ReDim 0, 0, 0, TDBDropDown1.Columns.Count - 1
    z.DeleteRows 0
    If RS.RecordCount > 0 Then z.LoadRows RS.GetRows
    TDBDropDown1.Rebind
    UpdatefQuick
    GoEvent "ADD"
    fTanggal = pServerDate
End Sub

Private Sub fDelete_Click()
On Error GoTo err
    CN.BeginTrans
    a = "delete from t_NR" & pTipe & " where NoNR='" & fNo & "'"
    ExecMe a
    a = "delete from t_NRDetail" & pTipe & " where NoNR='" & fNo & "'"
    ExecMe a
    For i = 0 To Grid1.RowCount - 1
        a = "update t_SCDetail" & pTipe & " set Sisa=Sisa-" & cNum(Grid1(i, "(KG)")) & " where NoSC='" & Grid1(i, "NoSC") & "' and IdSC=" & Grid1(i, "IdSC")
        ExecMe a
        a = "update m_stock" & pTipe & " set JumlahBox=JumlahBox-" & Grid1(i, "(BOX)") & _
            ",JumlahKG=JumlahKG-" & cNum(Grid1(i, "(KG)")) & " where IdStock=" & Grid1(i, "IdStock")
        ExecMe a
    Next
    CN.CommitTrans
    MsgBox "SUKSES"
    UpdatefQuick
    GoEvent "ADD"
    SendData "1HAPUS NR NO: " & fNo & Chr(8)
    DoEvents
    Exit Sub
err:
    CN.RollbackTrans
    MsgBox "GAGAL"
End Sub

Private Sub ClearScreen()
    Grid1.Clear
    fNo = ""
    fTanggal = pServerDate
    fNo_LostFocus
    Grid1.Clear
    GoEvent "ADD"
End Sub

Private Sub fNew_Click()
    ClearScreen
End Sub

Private Sub fNo_LostFocus()
    BuatNomor fNo, fTanggal, pNomorNR, fQuick, "select max(NoNR) from t_NR" & pTipe & " where TanggalNR>" & pAddNoLong
End Sub

Private Sub fPrint_Click()
    FramePreview.PrintControl
End Sub

Private Sub fQuick_KeyDown(KeyCode As Integer, Shift As Integer)
    fQuick_Validate False
End Sub

Private Sub fQuick_Validate(Cancel As Boolean)
On Error Resume Next
    GetResult fQuick
End Sub

Private Sub FramePreview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragControl FramePreview
End Sub

Private Sub fSave_Click()
On Error GoTo err
    CN.BeginTrans
    fNo.Validate
    fTanggal.Validate
    fNoKW.Validate
    Grid1.Update
Dim tot As Double
    tot = 0
    For i = 0 To Grid1.RowCount - 1
        tot = tot + Grid1(i, "(KG)") * Grid1(i, "HARGA") + Grid1(i, "DISC (KG)") * Grid1(i, "DISCOUNT")
    Next
    a = "select Kode from m_customer where Nama='" & fCustomer & "'"
    query a
    Kode = RS.Fields(0).Value
    a = "insert into t_NR" & pTipe & "(NoNR,TanggalNR,NoKW,KodeCustomer,Total, ShortNR) values('" & _
        fNo & _
        "'," & cD(fTanggal) & _
        ",'" & fNoKW & _
        "'," & Kode & _
        "," & cNum(tot) & ",'" & Left(fNo, 5) & "/" & Right(fNo, 2) & "')"
    If ExecMe(a) = 0 Then GoTo err
    For i = 0 To Grid1.RowCount - 1
        a = "insert into t_NRDetail" & pTipe & "(NoNR,IdNR,IdStock,ReturBox,ReturKg,Harga,DiscKg,Discount,NoSJ) values('" & _
            fNo & _
            "'," & i & _
            "," & Grid1(i, "IdStock") & _
            "," & cNum(Grid1(i, "(BOX)")) & _
            "," & cNum(Grid1(i, "(KG)")) & _
            "," & cNum(Grid1(i, "HARGA")) & _
            "," & cNum(Grid1(i, "DISC (KG)")) & _
            "," & cNum(Grid1(i, "DISCOUNT")) & _
            ",'" & Grid1(i, "NO SJ") & "')"
        If ExecMe(a) = 0 Then GoTo err
        sisa = Grid1(i, "(KG)")
        If sisa = "" Then sisa = 0
        a = "update t_SCDetail" & pTipe & " set Sisa=Sisa+" & cNum(sisa) & " where NoSC='" & Grid1(i, "NoSC") & "' and IdSC=" & Grid1(i, "IdSC")
        ExecMe a
        a = "update m_stock" & pTipe & " set JumlahBox=JumlahBox+" & cNum(Grid1(i, "(BOX)")) & _
            ",JumlahKG=JumlahKG+" & cNum(Grid1(i, "(KG)")) & " where IdStock=" & Grid1(i, "IdStock")
        ExecMe a
    Next
    CN.CommitTrans
    MsgBox "SUKSES"
    UpdatefQuick
    SendData "1BUAT NR NO: " & fNo & Chr(8)
    DoEvents
    GetResult fNo
    Exit Sub
err:
    CN.RollbackTrans
    MsgBox "GAGAL"
End Sub

Private Sub fTanggal_LostFocus()
    fNo_LostFocus
End Sub

Private Sub Grid1_AfterColEdit(ByVal ColIndex As Integer)
    If ColIndex = 0 Then TDBDropDown1_DropDownClose
End Sub

Private Sub TDBDropDown1_DropDownClose()
    Grid1.Columns(0).Value = TDBDropDown1.Columns(0).Value
    Grid1.Columns(1).Value = TDBDropDown1.Columns(1).Value
    Grid1.Columns(4).Value = TDBDropDown1.Columns(2).Value
    Grid1.Columns(7).Value = TDBDropDown1.Columns(4).Value
    Grid1.Columns(8).Value = TDBDropDown1.Columns(5).Value
    Grid1.Columns(9).Value = TDBDropDown1.Columns(6).Value
End Sub

