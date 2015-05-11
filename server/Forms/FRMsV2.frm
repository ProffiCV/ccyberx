VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmsv2 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "sServerXV3"
   ClientHeight    =   8385
   ClientLeft      =   255
   ClientTop       =   1515
   ClientWidth     =   14895
   Icon            =   "FRMsV2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MouseIcon       =   "FRMsV2.frx":030A
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   Picture         =   "FRMsV2.frx":0614
   ScaleHeight     =   8385
   ScaleWidth      =   14895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmUpdInfo8 
      Interval        =   800
      Left            =   11460
      Top             =   6780
   End
   Begin VB.Timer tmOut8 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   11430
      Top             =   6360
   End
   Begin VB.Timer tmCl8 
      Interval        =   600
      Left            =   11430
      Top             =   5820
   End
   Begin VB.Timer tmUpdInfo7 
      Interval        =   800
      Left            =   7740
      Top             =   6780
   End
   Begin VB.Timer tmOut7 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   7710
      Top             =   6300
   End
   Begin VB.Timer tmCL7 
      Interval        =   600
      Left            =   7710
      Top             =   5820
   End
   Begin VB.Timer tmHide 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   9210
      Top             =   300
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   7410
      PasswordChar    =   "*"
      TabIndex        =   56
      Top             =   -330
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Timer tmupicon 
      Interval        =   20
      Left            =   5520
      Top             =   7740
   End
   Begin VB.Timer tmTime 
      Interval        =   2000
      Left            =   6780
      Top             =   7800
   End
   Begin VB.Timer tmListenning 
      Interval        =   3000
      Left            =   3030
      Top             =   7800
   End
   Begin VB.Timer tmShowINfo 
      Interval        =   80
      Left            =   2520
      Top             =   7800
   End
   Begin VB.Timer tmResetCon 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1950
      Top             =   7800
   End
   Begin VB.Timer tmCL6 
      Interval        =   600
      Left            =   4050
      Top             =   5790
   End
   Begin VB.Timer tmOut6 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   4050
      Top             =   6210
   End
   Begin VB.Timer tmUpdInfo6 
      Interval        =   800
      Left            =   4050
      Top             =   6630
   End
   Begin VB.Timer tmCl5 
      Interval        =   600
      Left            =   300
      Top             =   5790
   End
   Begin VB.Timer tmOut5 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   300
      Top             =   6210
   End
   Begin VB.Timer tmUpdInfo5 
      Interval        =   800
      Left            =   300
      Top             =   6630
   End
   Begin VB.Timer tmCl4 
      Interval        =   600
      Left            =   11430
      Top             =   1950
   End
   Begin VB.Timer tmOut4 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   11430
      Top             =   2370
   End
   Begin VB.Timer tmUpdInfo4 
      Interval        =   800
      Left            =   11430
      Top             =   2790
   End
   Begin VB.Timer tmCl3 
      Interval        =   600
      Left            =   7650
      Top             =   2100
   End
   Begin VB.Timer tmOut3 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   7650
      Top             =   2580
   End
   Begin VB.Timer tmUpdInfo3 
      Interval        =   800
      Left            =   7680
      Top             =   3060
   End
   Begin VB.Timer tmCl2 
      Interval        =   600
      Left            =   3960
      Top             =   2160
   End
   Begin VB.Timer tmOut2 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3960
      Top             =   2640
   End
   Begin VB.Timer tmUpdInfo2 
      Interval        =   800
      Left            =   3960
      Top             =   3090
   End
   Begin VB.Timer tmCl1 
      Interval        =   600
      Left            =   270
      Top             =   2040
   End
   Begin VB.Timer tmOut1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   270
      Top             =   2460
   End
   Begin VB.Timer tmUpdInfo1 
      Interval        =   800
      Left            =   270
      Top             =   2880
   End
   Begin MSWinsockLib.Winsock wsCl1 
      Left            =   270
      Top             =   1620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5001
   End
   Begin MSWinsockLib.Winsock wsCl2 
      Left            =   3900
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   500
   End
   Begin MSWinsockLib.Winsock wsCl3 
      Left            =   7680
      Top             =   1620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2003
   End
   Begin MSWinsockLib.Winsock wsCl4 
      Left            =   11430
      Top             =   1530
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2004
   End
   Begin MSWinsockLib.Winsock wsCl5 
      Left            =   300
      Top             =   5370
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2005
   End
   Begin MSWinsockLib.Winsock wsCl6 
      Left            =   4050
      Top             =   5370
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2006
   End
   Begin MSWinsockLib.Winsock wsCl7 
      Left            =   7740
      Top             =   5340
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2007
   End
   Begin MSWinsockLib.Winsock wsCl8 
      Left            =   11460
      Top             =   5340
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2008
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "400$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   11
      Left            =   8580
      TabIndex        =   131
      Top             =   7230
      Width           =   540
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   8520
      TabIndex        =   130
      Top             =   5640
      Width           =   630
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   6
      Left            =   10455
      TabIndex        =   129
      Top             =   5910
      Width           =   630
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   7
      Left            =   10455
      TabIndex        =   128
      Top             =   6150
      Width           =   630
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   8
      Left            =   10455
      TabIndex        =   127
      Top             =   6390
      Width           =   630
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   10
      Left            =   10590
      TabIndex        =   126
      Top             =   6930
      Width           =   495
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "210$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   12
      Left            =   10590
      TabIndex        =   125
      Top             =   7230
      Width           =   495
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "240 MB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   13
      Left            =   10590
      TabIndex        =   124
      Top             =   7500
      Width           =   495
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "George Tavares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   10020
      TabIndex        =   123
      Top             =   4650
      Width           =   1065
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234-1234-1234-1234"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   9
      Left            =   9645
      TabIndex        =   122
      Top             =   6660
      Width           =   1440
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   5
      Left            =   10590
      TabIndex        =   121
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   10455
      TabIndex        =   120
      Top             =   5400
      Width           =   630
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   9885
      TabIndex        =   119
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Label lbi7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   9885
      TabIndex        =   118
      Top             =   5160
      Width           =   1200
   End
   Begin VB.Label fmePc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "cCyberXV2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   7
      Left            =   11730
      TabIndex        =   117
      Top             =   4170
      Width           =   1035
   End
   Begin VB.Label fmePc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "cCyberXV2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   7950
      TabIndex        =   116
      Top             =   4170
      Width           =   1035
   End
   Begin VB.Label fmePc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "cCyberXV2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   11700
      TabIndex        =   115
      Top             =   450
      Width           =   1035
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "400$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   12120
      TabIndex        =   114
      Top             =   7260
      Width           =   720
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   12030
      TabIndex        =   113
      Top             =   5670
      Width           =   840
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   13965
      TabIndex        =   112
      Top             =   5940
      Width           =   840
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   13965
      TabIndex        =   111
      Top             =   6180
      Width           =   840
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   13965
      TabIndex        =   110
      Top             =   6420
      Width           =   840
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   14085
      TabIndex        =   109
      Top             =   6960
      Width           =   720
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "210$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   14085
      TabIndex        =   108
      Top             =   7260
      Width           =   720
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "240 MB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   14070
      TabIndex        =   107
      Top             =   7530
      Width           =   735
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "George Tavares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   13230
      TabIndex        =   106
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234-1234-1234-1234"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   12660
      TabIndex        =   105
      Top             =   6690
      Width           =   2145
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   14085
      TabIndex        =   104
      Top             =   5670
      Width           =   720
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   13965
      TabIndex        =   103
      Top             =   5430
      Width           =   840
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   13035
      TabIndex        =   102
      Top             =   4950
      Width           =   1770
   End
   Begin VB.Label lbi8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   13035
      TabIndex        =   101
      Top             =   5190
      Width           =   1770
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   13590
      TabIndex        =   100
      Top             =   1410
      Width           =   1200
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   13590
      TabIndex        =   99
      Top             =   1170
      Width           =   1200
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   14160
      TabIndex        =   98
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   5
      Left            =   14295
      TabIndex        =   97
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234-1234-1234-1234"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   9
      Left            =   13350
      TabIndex        =   96
      Top             =   2910
      Width           =   1440
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "George Tavares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   13725
      TabIndex        =   95
      Top             =   900
      Width           =   1065
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "240 MB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   13
      Left            =   14295
      TabIndex        =   94
      Top             =   3780
      Width           =   495
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "210$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   12
      Left            =   14295
      TabIndex        =   93
      Top             =   3510
      Width           =   495
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "400$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   11
      Left            =   12300
      TabIndex        =   92
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   10
      Left            =   14295
      TabIndex        =   91
      Top             =   3180
      Width           =   495
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   8
      Left            =   14160
      TabIndex        =   90
      Top             =   2640
      Width           =   630
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   7
      Left            =   14160
      TabIndex        =   89
      Top             =   2400
      Width           =   630
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   6
      Left            =   14160
      TabIndex        =   88
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label lbi4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   12030
      TabIndex        =   87
      Top             =   1890
      Width           =   840
   End
   Begin VB.Image getMenu 
      Height          =   390
      Index           =   7
      Left            =   11250
      Picture         =   "FRMsV2.frx":12FEAE
      Top             =   4050
      Width           =   405
   End
   Begin VB.Image getMenu 
      Height          =   390
      Index           =   3
      Left            =   11220
      Picture         =   "FRMsV2.frx":130778
      Top             =   330
      Width           =   405
   End
   Begin VB.Label sleep 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   14295
      TabIndex        =   86
      Top             =   7920
      Width           =   540
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   9285
      TabIndex        =   85
      Top             =   1410
      Width           =   1770
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   9285
      TabIndex        =   84
      Top             =   1170
      Width           =   1770
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   10215
      TabIndex        =   83
      Top             =   1650
      Width           =   840
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   10335
      TabIndex        =   82
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234-1234-1234-1234"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   8910
      TabIndex        =   81
      Top             =   2880
      Width           =   2145
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "George Tavares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   9480
      TabIndex        =   80
      Top             =   900
      Width           =   1575
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "240 MB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   10320
      TabIndex        =   79
      Top             =   3780
      Width           =   735
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "210$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   10335
      TabIndex        =   78
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   10335
      TabIndex        =   77
      Top             =   3180
      Width           =   720
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   10215
      TabIndex        =   76
      Top             =   2610
      Width           =   840
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   10215
      TabIndex        =   75
      Top             =   2370
      Width           =   840
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   10215
      TabIndex        =   74
      Top             =   2130
      Width           =   840
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   6195
      TabIndex        =   73
      Top             =   5160
      Width           =   1200
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   6195
      TabIndex        =   72
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   6765
      TabIndex        =   71
      Top             =   5400
      Width           =   630
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   5
      Left            =   6900
      TabIndex        =   70
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234-1234-1234-1234"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   9
      Left            =   5955
      TabIndex        =   69
      Top             =   6660
      Width           =   1440
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "George Tavares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   6330
      TabIndex        =   68
      Top             =   4650
      Width           =   1065
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "240 MB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   13
      Left            =   6900
      TabIndex        =   67
      Top             =   7500
      Width           =   495
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "210$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   12
      Left            =   6900
      TabIndex        =   66
      Top             =   7230
      Width           =   495
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   10
      Left            =   6900
      TabIndex        =   65
      Top             =   6930
      Width           =   495
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   8
      Left            =   6765
      TabIndex        =   64
      Top             =   6390
      Width           =   630
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   7
      Left            =   6765
      TabIndex        =   63
      Top             =   6150
      Width           =   630
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   6
      Left            =   6765
      TabIndex        =   62
      Top             =   5910
      Width           =   630
   End
   Begin VB.Label ver 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "rv150808"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   14235
      TabIndex        =   61
      Top             =   8160
      Width           =   585
   End
   Begin VB.Label cmdBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   14610
      MouseIcon       =   "FRMsV2.frx":131042
      MousePointer    =   99  'Custom
      TabIndex        =   60
      Top             =   30
      Width           =   255
   End
   Begin VB.Label cmdBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   14340
      MouseIcon       =   "FRMsV2.frx":13134C
      MousePointer    =   99  'Custom
      TabIndex        =   59
      Top             =   30
      Width           =   255
   End
   Begin VB.Label lbt 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   13470
      TabIndex        =   58
      Top             =   30
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   8355
      Left            =   9870
      Picture         =   "FRMsV2.frx":131656
      Top             =   0
      Width           =   5040
   End
   Begin VB.Label mnu 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   3135
      MouseIcon       =   "FRMsV2.frx":1BA7C8
      MousePointer    =   99  'Custom
      TabIndex        =   57
      Top             =   30
      Width           =   435
   End
   Begin VB.Image pBusy 
      Height          =   390
      Left            =   15600
      Picture         =   "FRMsV2.frx":1BAAD2
      Top             =   240
      Width           =   390
   End
   Begin VB.Image pFree 
      Height          =   375
      Left            =   15660
      Picture         =   "FRMsV2.frx":1BB334
      Top             =   1620
      Width           =   390
   End
   Begin VB.Image getMenu 
      Height          =   390
      Index           =   6
      Left            =   7470
      Picture         =   "FRMsV2.frx":1BBB46
      Top             =   4050
      Width           =   405
   End
   Begin VB.Image getMenu 
      Height          =   390
      Index           =   5
      Left            =   3810
      Picture         =   "FRMsV2.frx":1BC410
      Top             =   4050
      Width           =   405
   End
   Begin VB.Image getMenu 
      Height          =   390
      Index           =   4
      Left            =   90
      Picture         =   "FRMsV2.frx":1BCCDA
      Top             =   4050
      Width           =   405
   End
   Begin VB.Image getMenu 
      Height          =   390
      Index           =   2
      Left            =   7470
      Picture         =   "FRMsV2.frx":1BD5A4
      Top             =   330
      Width           =   405
   End
   Begin VB.Image getMenu 
      Height          =   390
      Index           =   1
      Left            =   3810
      Picture         =   "FRMsV2.frx":1BDE6E
      Top             =   330
      Width           =   405
   End
   Begin VB.Image getMenu 
      Height          =   390
      Index           =   0
      Left            =   120
      Picture         =   "FRMsV2.frx":1BE738
      Top             =   330
      Width           =   405
   End
   Begin VB.Image block 
      Height          =   390
      Left            =   15600
      Picture         =   "FRMsV2.frx":1BF002
      Top             =   960
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Move sServerXV3 Window From Here!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6630
      MousePointer    =   15  'Size All
      TabIndex        =   55
      ToolTipText     =   "Click and Move the Mouse..."
      Top             =   30
      Width           =   2805
   End
   Begin VB.Label fmePc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "cCyberXV2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   4290
      TabIndex        =   54
      Top             =   4170
      Width           =   1035
   End
   Begin VB.Label fmePc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "cCyberXV2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   570
      TabIndex        =   53
      Top             =   4170
      Width           =   1035
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   4830
      TabIndex        =   52
      Top             =   5640
      Width           =   630
   End
   Begin VB.Label lbi6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "400$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   11
      Left            =   4890
      TabIndex        =   51
      Top             =   7230
      Width           =   540
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   1170
      TabIndex        =   50
      Top             =   5640
      Width           =   630
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   6
      Left            =   3030
      TabIndex        =   49
      Top             =   5910
      Width           =   630
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   7
      Left            =   3030
      TabIndex        =   48
      Top             =   6150
      Width           =   630
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   8
      Left            =   3030
      TabIndex        =   47
      Top             =   6390
      Width           =   630
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   10
      Left            =   3165
      TabIndex        =   46
      Top             =   6930
      Width           =   495
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "400$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   11
      Left            =   1230
      TabIndex        =   45
      Top             =   7230
      Width           =   540
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "210$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   12
      Left            =   3165
      TabIndex        =   44
      Top             =   7260
      Width           =   495
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "240 MB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   13
      Left            =   3165
      TabIndex        =   43
      Top             =   7530
      Width           =   495
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "George Tavares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   2595
      TabIndex        =   42
      Top             =   4650
      Width           =   1065
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234-1234-1234-1234"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   9
      Left            =   2220
      TabIndex        =   41
      Top             =   6660
      Width           =   1440
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   5
      Left            =   3165
      TabIndex        =   40
      Top             =   5670
      Width           =   495
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   3030
      TabIndex        =   39
      Top             =   5430
      Width           =   630
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   2460
      TabIndex        =   38
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Label lbi5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   2460
      TabIndex        =   37
      Top             =   5160
      Width           =   1200
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   8460
      TabIndex        =   36
      Top             =   1920
      Width           =   630
   End
   Begin VB.Label lbi3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "400$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   11
      Left            =   8520
      TabIndex        =   35
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   6090
      TabIndex        =   34
      Top             =   1410
      Width           =   1200
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   6090
      TabIndex        =   33
      Top             =   1170
      Width           =   1200
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   6660
      TabIndex        =   32
      Top             =   1650
      Width           =   630
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   5
      Left            =   6795
      TabIndex        =   31
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234-1234-1234-1234"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   9
      Left            =   5850
      TabIndex        =   30
      Top             =   2910
      Width           =   1440
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "George Tavares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5715
      TabIndex        =   29
      Top             =   900
      Width           =   1575
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "240 MB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   13
      Left            =   6795
      TabIndex        =   28
      Top             =   3810
      Width           =   495
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "210$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   12
      Left            =   6795
      TabIndex        =   27
      Top             =   3510
      Width           =   495
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "400$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   11
      Left            =   4830
      TabIndex        =   26
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   10
      Left            =   6795
      TabIndex        =   25
      Top             =   3210
      Width           =   495
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   8
      Left            =   6660
      TabIndex        =   24
      Top             =   2640
      Width           =   630
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   7
      Left            =   6660
      TabIndex        =   23
      Top             =   2400
      Width           =   630
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   6
      Left            =   6660
      TabIndex        =   22
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label lbi2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   4800
      TabIndex        =   21
      Top             =   1890
      Width           =   630
   End
   Begin VB.Label mnu 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   2370
      MouseIcon       =   "FRMsV2.frx":1BF8CC
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   30
      Width           =   480
   End
   Begin VB.Label mnu 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   1380
      MouseIcon       =   "FRMsV2.frx":1BFBD6
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   30
      Width           =   720
   End
   Begin VB.Label mnu 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sServerXV3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   60
      MouseIcon       =   "FRMsV2.frx":1BFEE0
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   30
      Width           =   1095
   End
   Begin VB.Label msg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
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
      Left            =   60
      TabIndex        =   17
      Top             =   7890
      Width           =   660
   End
   Begin VB.Label fmePc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "cCyberXV2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   7950
      TabIndex        =   16
      Top             =   450
      Width           =   1035
   End
   Begin VB.Label fmePc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "cCyberXV2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   4320
      TabIndex        =   15
      Top             =   450
      Width           =   1035
   End
   Begin VB.Label fmePc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "cCyberXV2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   630
      TabIndex        =   14
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   2430
      TabIndex        =   13
      Top             =   1410
      Width           =   1200
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-12-06 18:50:30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   2430
      TabIndex        =   12
      Top             =   1140
      Width           =   1200
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   3000
      TabIndex        =   11
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   5
      Left            =   3135
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234-1234-1234-1234"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   9
      Left            =   2190
      TabIndex        =   9
      Top             =   2910
      Width           =   1440
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "George Tavares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2055
      TabIndex        =   8
      Top             =   870
      Width           =   1575
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "240 MB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   13
      Left            =   3135
      TabIndex        =   7
      Top             =   3780
      Width           =   495
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "210$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   12
      Left            =   3135
      TabIndex        =   6
      Top             =   3510
      Width           =   495
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "400$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   11
      Left            =   1170
      TabIndex        =   5
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100$00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   10
      Left            =   3135
      TabIndex        =   4
      Top             =   3180
      Width           =   495
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   8
      Left            =   3000
      TabIndex        =   3
      Top             =   2640
      Width           =   630
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   7
      Left            =   3000
      TabIndex        =   2
      Top             =   2370
      Width           =   630
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   6
      Left            =   3000
      TabIndex        =   1
      Top             =   2130
      Width           =   630
   End
   Begin VB.Label lbi1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   1140
      TabIndex        =   0
      Top             =   1890
      Width           =   630
   End
   Begin VB.Menu menu 
      Caption         =   "sServerXV2"
      Visible         =   0   'False
      Begin VB.Menu cmdShowNet 
         Caption         =   "&Show Network"
         Enabled         =   0   'False
      End
      Begin VB.Menu cmdConfig 
         Caption         =   "&Configure"
      End
      Begin VB.Menu cmdResetPass 
         Caption         =   "&Reset Password"
      End
      Begin VB.Menu sp0 
         Caption         =   "-"
      End
      Begin VB.Menu cmdExit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu cmdAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu cmdregist 
         Caption         =   "&Enter Registration Code"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Visible         =   0   'False
      Begin VB.Menu cmdred 
         Caption         =   "&Reduced"
      End
      Begin VB.Menu cmdOnlineUsers 
         Caption         =   "Online Users"
         Enabled         =   0   'False
      End
      Begin VB.Menu cmdOfflines 
         Caption         =   "Offline Users"
         Enabled         =   0   'False
      End
      Begin VB.Menu sp 
         Caption         =   "-"
      End
      Begin VB.Menu cmdUseDetails 
         Caption         =   "Session History"
      End
   End
   Begin VB.Menu mnuLock 
      Caption         =   "Lock"
      Visible         =   0   'False
      Begin VB.Menu cmdLock 
         Caption         =   "Lo&ck"
         Index           =   0
      End
      Begin VB.Menu cmdLock 
         Caption         =   "&Unlock"
         Index           =   1
      End
      Begin VB.Menu cmdLock 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu cmdLock 
         Caption         =   "&Cancel"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuClt 
         Caption         =   "Options"
         Begin VB.Menu cmdLocks 
            Caption         =   "Logoff"
            Index           =   3
         End
         Begin VB.Menu cmdLocks 
            Caption         =   "&Restart"
            Index           =   4
         End
         Begin VB.Menu cmdLocks 
            Caption         =   "&Shutdown"
            Index           =   5
         End
         Begin VB.Menu cmdLocks 
            Caption         =   "&Free Code"
            Index           =   6
         End
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Visible         =   0   'False
      Begin VB.Menu cmdDontSleep 
         Caption         =   "&Disable Auto-Sleep"
      End
      Begin VB.Menu cmdAlls 
         Caption         =   "&Logoff All"
         Index           =   0
      End
      Begin VB.Menu cmdAlls 
         Caption         =   "&Restart All"
         Index           =   1
      End
      Begin VB.Menu cmdAlls 
         Caption         =   "&Shutdown All"
         Index           =   2
      End
      Begin VB.Menu cmdAlls 
         Caption         =   "&Reset Connection"
         Index           =   3
      End
      Begin VB.Menu cmdAlls 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu cmdAlls 
         Caption         =   """Flight Mode"""
         Index           =   5
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu cmdReg 
         Caption         =   "&Registration"
      End
   End
End
Attribute VB_Name = "frmsv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim oldx&, oldy&
Dim wasMinimized As Boolean

Private dblOK As Boolean
'aguardar resposta do cliente
'fechar no fim
Private flgWaitWs1 As Boolean
Private flgWaitWs2 As Boolean
Private flgWaitWs3 As Boolean
Private flgWaitWs4 As Boolean
Private flgWaitWs5 As Boolean
Private flgWaitWs6 As Boolean
Private flgWaitWs7 As Boolean
Private flgWaitWs8 As Boolean



Private Sub cmdAbout_Click()
MsgBox Version
End Sub


Private Sub cmdAlls_Click(Index As Integer)

Select Case Index
Case 0, 1, 2
For targetP = 0 To 7
If (targetP + 1) <= getLicense(4863) Then
    cmdLock_Click Index + 4
End If

Next
Case 3 'reset connection
If MsgBox("Are you sure you want to reset all free clients Connection?", vbQuestion + vbYesNo, App.EXEName) = vbYes Then

    For Index = 1 To 8
    DoEvents
    If Index <= getLicense(4863) Then
            If pci(Index).state <> "BUSY" Then
                tellWS (Index - 1), "Restarting..." & wsCl1.LocalIP
                Pause 0.04
                RestartWS Index
            End If
            
    End If
    Next


End If

Case 4 'space
Case 5 'enter fly mode
Me.cmdAlls(5).Checked = Not Me.cmdAlls(5).Checked
tell "Flight mode turned " & IIf(Me.cmdAlls(5).Checked = True, " ON", "OFF") & "..."

 For Index = 1 To 8
    DoEvents
    If Index <= getLicense(4863) Then
        If Me.cmdAlls(5).Checked = True Then
                If pci(Index).state = "ON" Then
                    tellWS (Index - 1), "Fligh Mode: " & IIf(Me.cmdAlls(5).Checked = True, " ON", "OFF")
                    Pause 0.04
                    RestartWS Index, False
                End If
         Else
                If pci(Index).state = "OFF" Then
                 tellWS (Index - 1), "Fligh Mode: " & IIf(Me.cmdAlls(5).Checked = True, " ON", "OFF")
                 Pause 0.04
                   RestartWS Index
                End If
                
         End If
            
    End If
 Next


End Select
End Sub

Private Sub cmdBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdBox(Index).Appearance = 1

End Sub

Private Sub cmdBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdBox(Index).Appearance = 0
If Button = 1 And X > 0 Then
''Debug.print X, Y, Me.Left, Me.Top
Me.cmdBox(Index).Appearance = 0
Select Case Index
    Case 0
    AnimateWin Me.hwnd, 800, AW_BLEND Or AW_HIDE
    
        
        
        Me.WindowState = vbMinimized
        wasMinimized = True
        Me.Visible = True
        Me.cmdBox(Index).Appearance = 0
    Case 1
 Call letsgodown
End Select
End If

End Sub
Private Sub cmdConfig_Click()
'setit alwais ontop
frmLogin.getUserPass
End Sub

Private Sub cmdDontSleep_Click()
Me.cmdDontSleep.Checked = Not Me.cmdDontSleep.Checked
myInt = 0
Me.tmResetCon.Enabled = Not Me.cmdDontSleep.Checked
Me.sleep = "ASleep::False"
End Sub

Public Sub cmdExit_Click()
If MsgBox("Do you really want to exit?", vbQuestion + vbYesNo, App.EXEName) = vbYes Then
tmTime.Enabled = False

Me.tmCl1.Enabled = False
Me.tmCl2.Enabled = False
Me.tmCl3.Enabled = False
Me.tmCl4.Enabled = False
Me.tmCl5.Enabled = False
Me.tmCL6.Enabled = False
Me.tmCL7.Enabled = False
Me.tmCl8.Enabled = False

Me.wsCl1.Close
Me.wsCl2.Close
Me.wsCl3.Close
Me.wsCl4.Close
Me.wsCl5.Close
Me.wsCl6.Close
Me.wsCl7.Close
Me.wsCl8.Close

frmTr.sup.Close
Unload frmTr
closeDb


Load frmStart
End If

End Sub

Public Sub cmdLock_Click(Index As Integer)
Dim cme As String

Select Case Index
    Case 0
        cme = "LOK0"
        pci(targetP + 1).state = "LOCKON"
    Case 1
        cme = "LOK1"
        pci(targetP + 1).state = "LOCKOFF"
    Case 2
    Case 3
        Exit Sub
        
        'extern
        Case 4
              cme = "SHUT" & EWX_LOGOFF
        Case 5
              cme = "SHUT" & EWX_REBOOT
        Case 6
              cme = "SHUT" & EWX_POWEROFF
        Case 7
              cme = "CODE" & 0 'free code
End Select


Select Case targetP
    Case 0
            If Me.wsCl1.state = 7 Then
                wsCl1.SendData cme
            End If
    Case 1
            If Me.wsCl2.state = 7 Then
                wsCl2.SendData cme
            End If
    Case 2
            If Me.wsCl3.state = 7 Then
                wsCl3.SendData cme
            End If
    Case 3
            If Me.wsCl4.state = 7 Then
                wsCl4.SendData cme
            End If
    Case 4
            If Me.wsCl5.state = 7 Then
                wsCl5.SendData cme
            End If
    Case 5
            If Me.wsCl6.state = 7 Then
                wsCl6.SendData cme
            End If
    Case 6
            If Me.wsCl7.state = 7 Then
                wsCl7.SendData cme
            End If
    Case 7
            If Me.wsCl8.state = 7 Then
                wsCl8.SendData cme
            End If
End Select

Call putIcon(targetP, Me.block)
End Sub

Private Sub cmdLocks_Click(Index As Integer)

If Index = 6 Then
    Select Case pci(targetP + 1).state
        Case Is = "BUSY"
            cmdLock_Click 1 + Index
        Case Is = "", "ON", "OFF"
        Case Else
            If Me.txtPass.Visible = False And Me.txtPass.Text = "" Then
                Me.txtPass.Visible = True
                Me.tmHide.Enabled = True
                Else
                If Me.txtPass.Text = "maintenancestaff3209" Then
                Me.txtPass.Text = ""
                cmdLock_Click 1 + Index
                End If
                
                Exit Sub
                
            End If
            
    End Select
    
Else
    cmdLock_Click 1 + Index
End If

End Sub

Private Sub cmdred_Click()
tell "Not Implemented yet..."
Exit Sub
Me.cmdred.Checked = Not Me.cmdred.Checked
If Me.cmdred.Checked = True Then
Me.WindowState = vbMinimized
Form1.Show
Else
Unload Form1
End If

End Sub

Private Sub cmdreg_Click()
frmReg.Show vbModal
End Sub

Private Sub cmdregist_Click()
Dim regcode As String

If demo = True Then
regcode = InputBox("Enter the registration code", App.EXEName & " Registration", "")

If regcode = "" Then Exit Sub

'register here
SaveSetting App.EXEName, "key", "demo", regcode
buildLicense
Else
tell "sServerXV3 is already Activated"
End If


End Sub

Private Sub cmdResetPass_Click()
Dim tmpi$
tmpi$ = InputBox("Which is the reason to clean the password?", App.EXEName, "")
If tmpi$ = "theownedson1997198120052324.bodix." Then
DeleteSetting App.EXEName, "Data", "pass"
MsgBox "Your Password was cleaned." & NL & _
"You should create another password.", vbExclamation, App.EXEName
Else
    If tmpi$ <> "" Then tell "Wrong reason..."
End If

End Sub

Private Sub cmdUseDetails_Click()
frmHist.Show
End Sub


Public Sub myLoad()


Dim itr&


itr = 4001
Me.wsCl1.LocalPort = itr
Me.wsCl2.LocalPort = itr + 1
Me.wsCl3.LocalPort = itr + 2
Me.wsCl4.LocalPort = itr + 3
Me.wsCl5.LocalPort = itr + 4
Me.wsCl6.LocalPort = itr + 5
Me.wsCl7.LocalPort = itr + 6
Me.wsCl8.LocalPort = itr + 7

frmTr.sup.LocalPort = 4000
Me.Enabled = False

Me.Caption = Me.Caption & "."
    tell "Ready! " '& Version
    Call SetupFrames
tmOutInterval = 120

'''''''''''''''''''''''''''''''''''''''''Clear all
Dim obj
For Each obj In Me
If TypeName(obj) = "Label" Then
    If Left(obj.Name, 3) = "lbi" Then
    obj.Caption = ""
    obj.ForeColor = RGB(220, 220, 220)
    End If
    
End If
Next
''''''''''''''''''''''''''''''''''''''''''''''''''''
setPrices 0, 0, 0, 0

'preparar o buffer detalhes das contas...
creatDB False
freeCards

Me.Enabled = True
Load frmTr
frmTr.tmn.Enabled = True


HideCaption Me.hwnd
Do
DoEvents
Me.Visible = True
Loop Until Me.Visible = True



Me.Height = 8420
'Call AnimateWin(frmsv2.hwnd, 10, AW_HIDE)

'For itr = 0 To 5
'Call putIcon(itr, Nothing)
'Next

Me.Visible = False
Call AnimateWin(frmsv2.hwnd, 200, AW_CENTER)
Me.Refresh

''''''''''''''''

End Sub

Private Sub letsgodown()
cmdExit_Click
End Sub


Private Sub Form_Load()
DetectComponents

If demo = False Then
Me.mnu(3).Left = -12500
End If



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
Cancel = True

frmStart.Show

End If

End Sub

Private Function putIcon(Index, icons)
Me.getMenu(Index).Picture = icons
End Function

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbNormal Then
    frmsv2.cmdred.Checked = False
    Unload Form1
        If wasMinimized = True Then
        wasMinimized = False
        Me.Visible = False
         AnimateWin Me.hwnd, 300, AW_BLEND Or AW_ACTIVATE
         Me.Refresh
        End If
     
    ElseIf Me.WindowState = vbMinimized Then
     wasMinimized = True

End If

End Sub

Private Sub getmenu_DblClick(Index As Integer)
If dblOK = True Then
    targetP = Index
    PopupMenu mnuLock, 4
End If

End Sub

Private Sub getmenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
dblOK = Button = 1
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
oldx = X
oldy = Y
Me.Label1.Caption = "Move the mouse..."

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tell "Click here and move the mouse to move the window."
If Button = 1 Then

Me.Left = Me.Left + (X - oldx)
Me.Top = Me.Top + (Y - oldy)
Me.Label1.Caption = "Relative Screen Position (" & Me.Left \ Screen.TwipsPerPixelX & "," & Me.Top \ Screen.TwipsPerPixelY & ")"

End If

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Label1 = "Move sServerXV3 Window From Here!"
End Sub

Private Sub mnu_Click(Index As Integer)
Select Case Index
Case 0
    PopupMenu menu
Case 1
    PopupMenu mnuOpt
Case 2
    PopupMenu mnuView
Case 3
    PopupMenu mnuHelp
End Select

End Sub

Private Sub tmCl1_Timer()
Me.tmCl1.Enabled = False
With Me.fmePc(0)
''Debug.print "MY STATE 1" & Me.wsCl1.state
        Select Case Me.wsCl1.state
            Case sckClosed
                Me.fmePc(0).ForeColor = RGB(80, 90, 90)
                tellWS 0, """Protected!"""
               pci(1).state = "OFF"
            Case 1
            Case sckListening
                pci(1).state = "ON"
                tellWS 0, "Ready!"
                Me.fmePc(0).ForeColor = vbWhite
            Case 3
            Case 4
            Case 5
            Case 6
            Case sckConnected
            Me.fmePc(0).ForeColor = RGB(0, 200, 0)
            tellWS 0, pci(1).pcName & " Online..."
            
            Case 8
                tellWS 0, "Closing..." & wsCl1.RemoteHost
                RestartWS 1
            Case sckError
                tellWS 0, "Closing..." & wsCl1.RemoteHost
                RestartWS 1
                
        End Select
End With
Me.tmCl1.Enabled = True

End Sub



'TIMER 2
Private Sub tmCl2_Timer()
Me.tmCl2.Enabled = False
With Me.fmePc(1)
''Debug.print "MY STATE " & Me.wsCl2.state
        Select Case Me.wsCl2.state
            Case sckClosed
                pci(2).state = "OFF"
                tellWS 1, """Protected!"""
                .ForeColor = RGB(80, 90, 90)
            Case 1
            Case sckListening
                tellWS 1, "Ready!"
                .ForeColor = vbWhite
                pci(2).state = "ON"
                
            Case 3
            Case 4
            Case 5
            Case 6
            Case sckConnected
            .ForeColor = RGB(0, 200, 0)
            tellWS 1, pci(2).pcName & " Online..."
'                    If pci(2).state = "BUSY" Then
'                            Else
'                     pci(2).state = "FREE"
'                    End If
'
            Case 8
                tellWS 1, "Closing..." & wsCl2.RemoteHost
                RestartWS 2
                ''Debug.print Err.Number
                ''Debug.print "ok"
                
            Case sckError
                tellWS 1, "Closing..." & wsCl2.RemoteHost
                RestartWS 2
        End Select
End With
Me.tmCl2.Enabled = True

End Sub

'TIMER 3
Private Sub tmCl3_Timer()
Me.tmCl3.Enabled = False
With Me.fmePc(2)
''Debug.print "MY STATE " & Me.wsCl3.state
        Select Case Me.wsCl3.state
            Case sckClosed
            pci(3).state = "OFF"
                tellWS 2, """Protected!"""
                .ForeColor = RGB(80, 90, 90)
            Case 1
            Case sckListening
            pci(3).state = "ON"
                tellWS 2, "Ready!"
                .ForeColor = vbWhite
            Case 3
            Case 4
            Case 5
            Case 6
            Case sckConnected
'                     If pci(3).state = "BUSY" Then
'                            Else
'                     pci(3).state = "FREE"
'                    End If
            .ForeColor = RGB(0, 200, 0)
             tellWS 2, pci(3).pcName & " Online..."
            Case 8
                tellWS 2, "Closing..." & wsCl3.RemoteHost
                RestartWS 3
            Case sckError
                tellWS 2, "Closing..." & wsCl3.RemoteHost
                RestartWS 3
        End Select
End With
Me.tmCl3.Enabled = True

End Sub

'TIMER 4
Private Sub tmCl4_Timer()
Me.tmCl4.Enabled = False
With Me.fmePc(3)
''Debug.print "MY STATE " & Me.wsCl4.state
        Select Case Me.wsCl4.state
            Case sckClosed
                tellWS 3, """Protected!"""
                .ForeColor = RGB(80, 90, 90)
                pci(4).state = "OFF"
            Case 1
            Case sckListening
                tellWS 3, "Ready!"
                .ForeColor = vbWhite
                pci(4).state = "ON"
            Case 3
            Case 4
            Case 5
            Case 6
            Case sckConnected
'                  If pci(4).state = "BUSY" Then
'                            Else
'                     pci(4).state = "FREE"
'                    End If
            .ForeColor = RGB(0, 200, 0)
             tellWS 3, pci(4).pcName & " Online..."
            Case 8
                tellWS 3, "Closing..." & wsCl4.RemoteHost
                RestartWS 4
            Case sckError
                tellWS 3, "Closing..." & wsCl4.RemoteHost
                RestartWS 4
        End Select
End With
Me.tmCl4.Enabled = True

End Sub

'TIMER 5
Private Sub tmCl5_Timer()
Me.tmCl5.Enabled = False
With Me.fmePc(4)
''Debug.print "MY STATE " & Me.wsCl5.state
        Select Case Me.wsCl5.state
            Case sckClosed
                tellWS 4, """Protected!"""
                .ForeColor = RGB(80, 90, 90)
                pci(5).state = "OFF"
            Case 1
            Case sckListening
            pci(5).state = "ON"
                tellWS 4, "Ready!"
                .ForeColor = vbWhite
            Case 3
            Case 4
            Case 5
            Case 6
            Case sckConnected
'                  If pci(5).state = "BUSY" Then
'                            Else
'                     pci(5).state = "FREE"
'                    End If
            .ForeColor = RGB(0, 200, 0)
             tellWS 4, pci(5).pcName & " Online..."
            Case 8
                tellWS 4, "Closing..." & wsCl1.RemoteHost
                RestartWS 5
            Case sckError
                tellWS 4, "Closing..." & wsCl1.RemoteHost
                RestartWS 5
        End Select
End With
Me.tmCl5.Enabled = True

End Sub

'TIMER 6
Private Sub tmCl6_Timer()
Me.tmCL6.Enabled = False
With Me.fmePc(5)
''Debug.print "MY STATE " & Me.wsCl6.state
        Select Case Me.wsCl6.state
            Case sckClosed
                tellWS 5, """Protected!"""
                .ForeColor = RGB(80, 90, 90)
            pci(6).state = "OFF"
            Case 1
            Case sckListening
            pci(6).state = "ON"
                tellWS 5, "Ready!"
                .ForeColor = vbWhite
            Case 3
            Case 4
            Case 5
            Case 6
            Case sckConnected
'                     If pci(6).state = "BUSY" Then
'                            Else
'                     pci(6).state = "FREE"
'                    End If
            .ForeColor = RGB(0, 200, 0)
             tellWS 5, pci(6).pcName & " Online..."
            Case 8
                tellWS 5, "Closing..." & wsCl6.RemoteHost
                RestartWS 6
            Case sckError
                tellWS 5, "Closing..." & wsCl6.RemoteHost
                RestartWS 6
        End Select
End With
Me.tmCL6.Enabled = True

End Sub

Private Sub tmCl7_Timer()
Me.tmCL7.Enabled = False
With Me.fmePc(6)
''Debug.print "MY STATE " & Me.wsCl6.state
        Select Case Me.wsCl7.state
            Case sckClosed
                tellWS 6, """Protected!"""
                .ForeColor = RGB(80, 90, 90)
            pci(7).state = "OFF"
            Case 1
            Case sckListening
            pci(7).state = "ON"
                tellWS 6, "Ready!"
                .ForeColor = vbWhite
            Case 3
            Case 4
            Case 5
            Case 6
            Case sckConnected
'                     If pci(6).state = "BUSY" Then
'                            Else
'                     pci(6).state = "FREE"
'                    End If
            .ForeColor = RGB(0, 200, 0)
             tellWS 6, pci(7).pcName & " Online..."
            Case 8
                tellWS 6, "Closing..." & wsCl7.RemoteHost
                RestartWS 7
            Case sckError
                tellWS 6, "Closing..." & wsCl7.RemoteHost
                RestartWS 7
        End Select
End With
Me.tmCL7.Enabled = True

End Sub

Private Sub tmCl8_Timer()
Me.tmCl8.Enabled = False
With Me.fmePc(7)
''Debug.print "MY STATE " & Me.wsCl6.state
        Select Case Me.wsCl8.state
            Case sckClosed
                tellWS 7, """Protected!"""
                .ForeColor = RGB(80, 90, 90)
            pci(8).state = "OFF"
            Case 1
            Case sckListening
            pci(8).state = "ON"
                tellWS 7, "Ready!"
                .ForeColor = vbWhite
            Case 3
            Case 4
            Case 5
            Case 6
            Case sckConnected
'                     If pci(6).state = "BUSY" Then
'                            Else
'                     pci(6).state = "FREE"
'                    End If
            .ForeColor = RGB(0, 200, 0)
             tellWS 7, pci(8).pcName & " Online..."
            Case 8
                tellWS 7, "Closing..." & wsCl8.RemoteHost
                RestartWS 8
            Case sckError
                tellWS 7, "Closing..." & wsCl8.RemoteHost
                RestartWS 8
        End Select
End With
Me.tmCl8.Enabled = True

End Sub


Private Sub tmHide_Timer()
Me.tmHide.Enabled = False
Me.txtPass.Visible = False
If Me.txtPass.Text <> "maintenancestaff3209" And Me.txtPass.Text <> "" Then
Me.txtPass.Text = ""
tell "Wrong Pass. Operation Canceled!"
ver.ForeColor = vbWhite
Else
If Me.txtPass.Text <> "" Then tell "...Again to continue.."
End If


End Sub

Private Sub tmListenning_Timer()
Me.tmListenning.Enabled = False
    RestartWS 1
    RestartWS 2
    RestartWS 3
    RestartWS 4
    RestartWS 5
    RestartWS 6
    RestartWS 7
    RestartWS 8
    
    If creatDB = True Then
    OpenDB
        checkCardsAutoRemove
    End If
    
    
    
End Sub

Private Sub tmOut1_Timer()
Me.tmOut1.Enabled = False
If flgWaitWs1 = False Then
If Me.wsCl1.state = 7 Then
flgWaitWs1 = True
Me.wsCl1.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
"Good Bye!" & vbCrLf
Do
DoEvents
Loop Until flgWaitWs1 = False

    RestartWS 1
End If
Else
flgWaitWs1 = False 'permitir quebra do loop senddata....
End If


End Sub

Private Sub tmOut2_Timer()
Me.tmOut2.Enabled = False
If flgWaitWs2 = False Then
If Me.wsCl2.state = 7 Then
flgWaitWs2 = True
Me.wsCl2.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
"Good Bye!" & vbCrLf
Do
DoEvents
Loop Until flgWaitWs2 = False

    RestartWS 2
End If
Else
flgWaitWs2 = False 'permitir quebra do loop senddata....
End If

End Sub

Private Sub tmOut3_Timer()
Me.tmOut3.Enabled = False
If flgWaitWs3 = False Then
If Me.wsCl3.state = 7 Then
flgWaitWs3 = True
Me.wsCl3.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
"Good Bye!" & vbCrLf
Do
DoEvents
Loop Until flgWaitWs3 = False

    RestartWS 3
End If
Else
flgWaitWs3 = False 'permitir quebra do loop senddata....
End If

End Sub

Private Sub tmOut4_Timer()
Me.tmOut4.Enabled = False
If flgWaitWs4 = False Then
If Me.wsCl4.state = 7 Then
flgWaitWs1 = True
Me.wsCl4.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
"Good Bye!" & vbCrLf
Do
DoEvents
Loop Until flgWaitWs4 = False

    RestartWS 4
End If
Else
flgWaitWs4 = False 'permitir quebra do loop senddata....
End If

End Sub

Private Sub tmOut5_Timer()
Me.tmOut5.Enabled = False
If flgWaitWs5 = False Then
If Me.wsCl5.state = 7 Then
flgWaitWs5 = True
Me.wsCl5.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
"Good Bye!" & vbCrLf
Do
DoEvents
Loop Until flgWaitWs5 = False

    RestartWS 5
End If
Else
flgWaitWs5 = False 'permitir quebra do loop senddata....
End If

End Sub

Private Sub tmOut6_Timer()
Me.tmOut6.Enabled = False
If flgWaitWs6 = False Then
If Me.wsCl6.state = 7 Then
flgWaitWs6 = True
Me.wsCl6.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
"Good Bye!" & vbCrLf
Do
DoEvents
Loop Until flgWaitWs6 = False

    RestartWS 6
End If
Else
flgWaitWs6 = False 'permitir quebra do loop senddata....
End If

End Sub

Private Sub tmOut7_Timer()
Me.tmOut7.Enabled = False
If flgWaitWs7 = False Then
If Me.wsCl7.state = 7 Then
flgWaitWs7 = True
Me.wsCl7.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
"Good Bye!" & vbCrLf
Do
DoEvents
Loop Until flgWaitWs7 = False

    RestartWS 7
End If
Else
flgWaitWs7 = False 'permitir quebra do loop senddata....
End If

End Sub

Private Sub tmOut8_Timer()
Me.tmOut8.Enabled = False
If flgWaitWs8 = False Then
If Me.wsCl8.state = 7 Then
flgWaitWs5 = True
Me.wsCl8.SendData App.EXEName & " Was not writed to be your Server." & vbCrLf & _
"Good Bye!" & vbCrLf
Do
DoEvents
Loop Until flgWaitWs8 = False

    RestartWS 8
End If
Else
flgWaitWs8 = False 'permitir quebra do loop senddata....
End If

End Sub
Private Sub tmResetCon_Timer()

DoEvents
Me.sleep = "ASleep :" & Format(myInt, "000\/") & Format(tmOutInterval, "000")

With Me.tmResetCon
    .Enabled = False
    myInt = myInt + 1
    If myInt >= tmOutInterval Then
    myInt = 0
    tell "Reseting connections"
       If pci(1).state <> "BUSY" Then RestartWS 1
       If pci(2).state <> "BUSY" Then RestartWS 2
       If pci(3).state <> "BUSY" Then RestartWS 3
       If pci(4).state <> "BUSY" Then RestartWS 4
       If pci(5).state <> "BUSY" Then RestartWS 5
       If pci(6).state <> "BUSY" Then RestartWS 6
       If pci(7).state <> "BUSY" Then RestartWS 7
       If pci(8).state <> "BUSY" Then RestartWS 8
       
    End If

    .Enabled = True
    
    
End With

End Sub

Private Sub tmShowINfo_Timer()



Me.tmShowINfo.Enabled = False
Dim tot&, usd&, pcu&, netu&
Dim ol As Label
Dim ind As Integer

For ind = 1 To 8
If ind > getLicense(4863) Then Exit For
DoEvents
    With pci(ind)
        If .dispInfo <> 0 Then 'mostrar dado
        Select Case ind
        Case 1
            Me.lbi1(0) = .clientID
            Me.lbi1(1) = .login 'login time
            Me.lbi1(2) = .logoff
            Me.lbi1(4) = .pcuTime
            Me.lbi1(5) = .pcuPrice
            Me.lbi1(7) = .netNow  'download net
            Me.lbi1(8) = .netPrice 'preço net
            Me.lbi1(9) = .code 'used code
            Me.lbi1(10) = .balTotal
            Me.lbi1(11) = .balUsed
            Me.lbi1(12) = .balRemain 'remain balance
            Me.lbi1(13) = .netTotal 'total 'bytes downb
        Case 2
            Me.lbi2(0) = .clientID
            Me.lbi2(1) = .login 'login time
            Me.lbi2(2) = .logoff
            Me.lbi2(4) = .pcuTime
            Me.lbi2(5) = .pcuPrice
            Me.lbi2(7) = .netNow  'download net
            Me.lbi2(8) = .netPrice 'preço net
            Me.lbi2(9) = .code 'used code
            Me.lbi2(10) = .balTotal
            Me.lbi2(11) = .balUsed
            Me.lbi2(12) = .balRemain 'remain balance
            Me.lbi2(13) = .netTotal 'total 'bytes downb
        Case 3
            Me.lbi3(0) = .clientID
            Me.lbi3(1) = .login 'login time
            Me.lbi3(2) = .logoff
            Me.lbi3(4) = .pcuTime
            Me.lbi3(5) = .pcuPrice
            Me.lbi3(7) = .netNow  'download net
            Me.lbi3(8) = .netPrice 'preço net
            Me.lbi3(9) = .code 'used code
            Me.lbi3(10) = .balTotal
            Me.lbi3(11) = .balUsed
            Me.lbi3(12) = .balRemain 'remain balance
            Me.lbi3(13) = .netTotal 'total 'bytes downb
        Case 4
            Me.lbi4(0) = .clientID
            Me.lbi4(1) = .login 'login time
            Me.lbi4(2) = .logoff
            Me.lbi4(4) = .pcuTime
            Me.lbi4(5) = .pcuPrice
            Me.lbi4(7) = .netNow  'download net
            Me.lbi4(8) = .netPrice 'preço net
            Me.lbi4(9) = .code 'used code
            Me.lbi4(10) = .balTotal
            Me.lbi4(11) = .balUsed
            Me.lbi4(12) = .balRemain 'remain balance
            Me.lbi4(13) = .netTotal
        Case 5
            Me.lbi5(0) = .clientID
            Me.lbi5(1) = .login 'login time
            Me.lbi5(2) = .logoff
            Me.lbi5(4) = .pcuTime
            Me.lbi5(5) = .pcuPrice
            Me.lbi5(7) = .netNow  'download net
            Me.lbi5(8) = .netPrice 'preço net
            Me.lbi5(9) = .code 'used code
            Me.lbi5(10) = .balTotal
            Me.lbi5(11) = .balUsed
            Me.lbi5(12) = .balRemain 'remain balance
            Me.lbi5(13) = .netTotal
        Case 6
            Me.lbi6(0) = .clientID
            Me.lbi6(1) = .login 'login time
            Me.lbi6(2) = .logoff
            Me.lbi6(4) = .pcuTime
            Me.lbi6(5) = .pcuPrice
            Me.lbi6(7) = .netNow  'download net
            Me.lbi6(8) = .netPrice 'preço net
            Me.lbi6(9) = .code 'used code
            Me.lbi6(10) = .balTotal
            Me.lbi6(11) = .balUsed
            Me.lbi6(12) = .balRemain 'remain balance
            Me.lbi6(13) = .netTotal
        Case 7
            Me.lbi7(0) = .clientID
            Me.lbi7(1) = .login 'login time
            Me.lbi7(2) = .logoff
            Me.lbi7(4) = .pcuTime
            Me.lbi7(5) = .pcuPrice
            Me.lbi7(7) = .netNow  'download net
            Me.lbi7(8) = .netPrice 'preço net
            Me.lbi7(9) = .code 'used code
            Me.lbi7(10) = .balTotal
            Me.lbi7(11) = .balUsed
            Me.lbi7(12) = .balRemain 'remain balance
            Me.lbi7(13) = .netTotal
        Case 8
            Me.lbi8(0) = .clientID
            Me.lbi8(1) = .login 'login time
            Me.lbi8(2) = .logoff
            Me.lbi8(4) = .pcuTime
            Me.lbi8(5) = .pcuPrice
            Me.lbi8(7) = .netNow  'download net
            Me.lbi8(8) = .netPrice 'preço net
            Me.lbi8(9) = .code 'used code
            Me.lbi8(10) = .balTotal
            Me.lbi8(11) = .balUsed
            Me.lbi8(12) = .balRemain 'remain balance
            Me.lbi8(13) = .netTotal
        End Select
        End If
    End With
    
    
Next

Me.tmShowINfo.Enabled = True

End Sub

Private Sub tmTime_Timer()
DoEvents
Me.lbt = Format(Now, "dd/mmm/yyyy hh:mm ")
tell getDataBaseDetails(3245) & " Prices: Net - " & _
Replace(Format(getPrices.Pnet, "#,##0.00"), ",", "$") & " PC per Hour - " & _
Replace(Format(getPrices.Pwindows, "#,##0.00"), ",", "$") & " Copy " & IIf(demo = True, " Demo", ", Licensed ")
End Sub


Private Sub SetupFrames()
'organiza os frames
Dim it%
For it = 0 To Me.fmePc.Count - 1
Debug.Print getLicense(4863)
    If it + 1 > getLicense(4863) Then
        Me.fmePc(it).Enabled = False
        
        Select Case it
            Case 0
                Me.tmCl1.Enabled = False
                Me.wsCl1.Close
                
            Case 1
                Me.tmCl2.Enabled = False
                Me.wsCl2.Close
            Case 2
                Me.tmCl3.Enabled = False
                Me.wsCl3.Close
            Case 3
                Me.tmCl4.Enabled = False
                Me.wsCl4.Close
            Case 4
                Me.tmCl5.Enabled = False
                Me.wsCl5.Close
            Case 5
                Me.tmCL6.Enabled = False
                Me.wsCl6.Close
            Case 6
                Me.tmCL7.Enabled = False
                Me.wsCl7.Close
            Case 7
                Me.tmCl8.Enabled = False
                Me.wsCl8.Close
        
        End Select
        
        tellWS it, "No License Found."
        Me.getMenu(it).Enabled = False
    End If
    
Next
End Sub

Private Sub tmUpdInfo1_Timer()
With Me
    
        Select Case pci(1).state
            Case Is = "ON"
            Case Is = "OFF"
                'ClearList Me.lstRemData(0)
            Case Is = "SAVE"
                'User is going out...
            
        End Select
  
    
End With

End Sub

Private Sub tmupicon_Timer()
DoEvents
If Me.Visible = False Then Exit Sub
Me.tmupicon.Enabled = False
Static Index As Integer


Select Case pci(Index + 1).state
    Case "BUSY"
       putIcon Index, Me.pBusy
    Case "LOCKON"
       putIcon Index, Me.block
    Case "LOCKOFF"
       putIcon Index, Me.pFree
    Case "OFF"
      putIcon Index, Nothing
    Case "ON"
      putIcon Index, Nothing
    Case "FREE"
      putIcon Index, Me.pFree
    
End Select

'Debug.Print pci(index + 1).state
Index = Index + 1
If Index = 8 Then Index = 0
Me.tmupicon.Enabled = True
End Sub

Private Sub txtPass_Change()
ver.ForeColor = vbRed

Me.tmHide.Enabled = False
Me.tmHide.Enabled = True
End Sub

'CONNECTIONS REQUESTED
Private Sub wsCl1_ConnectionRequest(ByVal requestID As Long)
If getLicense(4863) < 1 Then Exit Sub

tell "1#, Connection Requested From " & wsCl1.RemoteHostIP
tellWS 0, "Waiting Permission [" & wsCl1.RemoteHostIP & "]"
    With wsCl1
        If .state <> 7 Then
            .Close
            .Accept requestID
            If Me.wsCl1.state = 7 Then
                Me.wsCl1.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
                "Welcome. The Permission Timer Was Started..." & vbCrLf
                'request permission
                Me.tmOut1.Enabled = False
                Me.tmOut1.Enabled = True
            End If
            
        End If
    End With
    
    
End Sub

'DATA ARRIVAL
Private Sub wsCl1_DataArrival(ByVal bytesTotal As Long)
Dim pdta As String
With wsCl1
    If .state = sckConnected Then
        .GetData pdta
        InterpretData1 pdta
    End If
End With

End Sub
Private Sub wsCl2_DataArrival(ByVal bytesTotal As Long)
Dim pdta As String
With wsCl2
    If .state = sckConnected Then
        .GetData pdta
        InterpretData2 pdta
    End If
End With

End Sub
Private Sub wsCl3_DataArrival(ByVal bytesTotal As Long)
Dim pdta As String
With wsCl3
    If .state = sckConnected Then
        .GetData pdta
        InterpretData3 pdta
    End If
End With

End Sub
Private Sub wsCl4_DataArrival(ByVal bytesTotal As Long)
Dim pdta As String
With wsCl4
    If .state = sckConnected Then
        .GetData pdta
        InterpretData4 pdta
    End If
End With

End Sub
Private Sub wsCl5_DataArrival(ByVal bytesTotal As Long)
Dim pdta As String
With wsCl5
    If .state = sckConnected Then
        .GetData pdta
        InterpretData5 pdta
    End If
End With

End Sub
Private Sub wsCl6_DataArrival(ByVal bytesTotal As Long)
Dim pdta As String
With wsCl6
    If .state = sckConnected Then
        .GetData pdta
        InterpretData6 pdta
    End If
End With

End Sub

Private Sub wsCl7_ConnectionRequest(ByVal requestID As Long)
If getLicense(4863) < 1 Then Exit Sub

tell "7#, Connection Requested From " & wsCl7.RemoteHostIP
tellWS 6, "Waiting Permission [" & wsCl7.RemoteHostIP & "]"
    With wsCl7
        If .state <> 7 Then
            .Close
            .Accept requestID
            If Me.wsCl7.state = 7 Then
                Me.wsCl7.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
                "Welcome. The Permission Timer Was Started..." & vbCrLf
                'request permission
                Me.tmOut7.Enabled = False
                Me.tmOut7.Enabled = True
            End If
            
        End If
    End With
    
    
End Sub

'DATA ARRIVAL
Private Sub wsCl7_DataArrival(ByVal bytesTotal As Long)
Dim pdta As String
With wsCl7
    If .state = sckConnected Then
        .GetData pdta
        InterpretData7 pdta
        
    End If
End With

End Sub

Private Sub wsCl8_ConnectionRequest(ByVal requestID As Long)
If getLicense(4863) < 8 Then Exit Sub

tell "8#, Connection Requested From " & wsCl8.RemoteHostIP
tellWS 7, "Waiting Permission [" & wsCl8.RemoteHostIP & "]"
    With wsCl8
        If .state <> 7 Then
            .Close
            .Accept requestID
            If Me.wsCl8.state = 7 Then
                Me.wsCl8.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
                "Welcome. The Permission Timer Was Started..." & vbCrLf
                'request permission
                Me.tmOut8.Enabled = False
                Me.tmOut8.Enabled = True
            End If
            
        End If
    End With
    
    
End Sub

'DATA ARRIVAL
Private Sub wsCl8_DataArrival(ByVal bytesTotal As Long)
Dim pdta As String
With wsCl8
    If .state = sckConnected Then
        .GetData pdta
        InterpretData8 pdta
    End If
End With

End Sub

Private Sub wsCl2_ConnectionRequest(ByVal requestID As Long)
If getLicense(4863) < 2 Then Exit Sub

tell "2#, Connection Requested From " & wsCl2.RemoteHostIP
tellWS 1, "Waiting Permission [" & wsCl2.RemoteHostIP & "]"
    With wsCl2
        If .state <> 7 Then
            .Close
            .Accept requestID
            If Me.wsCl2.state = 7 Then
                Me.wsCl2.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
                "Welcome. The Permission Timer Was Started..." & vbCrLf
                Me.tmOut2.Enabled = False
                Me.tmOut2.Enabled = True
            End If
            
        End If
    End With
    
End Sub

Private Sub wscl3_ConnectionRequest(ByVal requestID As Long)
If getLicense(4863) < 3 Then Exit Sub

tell "3#, Connection Requested From " & wsCl3.RemoteHostIP
tellWS 2, "Waiting Permission [" & wsCl3.RemoteHostIP & "]"
    With wsCl3
        If .state <> 7 Then
            .Close
            .Accept requestID
            If Me.wsCl3.state = 7 Then
                Me.wsCl3.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
                "Welcome. The Permission Timer Was Started..." & vbCrLf
                Me.tmOut3.Enabled = False
                Me.tmOut3.Enabled = True
            End If
            
        End If
    End With
    
End Sub

Private Sub wscl4_ConnectionRequest(ByVal requestID As Long)
If getLicense(4863) < 4 Then Exit Sub

tell "4#, Connection Requested From " & wsCl4.RemoteHostIP
tellWS 3, "Waiting Permission [" & wsCl4.RemoteHostIP & "]"
    With wsCl4
        If .state <> 7 Then
            .Close
            .Accept requestID
            If Me.wsCl4.state = 7 Then
                Me.wsCl4.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
                "Welcome. The Permission Timer Was Started..." & vbCrLf
                Me.tmOut4.Enabled = False
                Me.tmOut4.Enabled = True
            End If
            
        End If
    End With
    
End Sub

Private Sub wscl5_ConnectionRequest(ByVal requestID As Long)
If getLicense(4863) < 5 Then Exit Sub

tell "5#, Connection Requested From " & wsCl5.RemoteHostIP
tellWS 4, "Waiting Permission [" & wsCl5.RemoteHostIP & "]"
    With wsCl5
        If .state <> 7 Then
            .Close
            .Accept requestID
            If Me.wsCl5.state = 7 Then
                Me.wsCl5.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
                "Welcome. The Permission Timer Was Started..." & vbCrLf
                Me.tmOut5.Enabled = False
                Me.tmOut5.Enabled = True
            End If
            
        End If
    End With
    
End Sub

Private Sub wscl6_ConnectionRequest(ByVal requestID As Long)
If getLicense(4863) < 6 Then Exit Sub

tell "6#, Connection Requested From " & wsCl6.RemoteHostIP
tellWS 5, "Waiting Permission [" & wsCl6.RemoteHostIP & "]"
    With wsCl6
        If .state <> 7 Then
            .Close
            .Accept requestID
            If Me.wsCl6.state = 7 Then
                Me.wsCl6.SendData "GIVE" & App.EXEName & " Copyright(c) 2005-2006 Edson Martins " & vbCrLf & _
                "Welcome. The Permission Timer Was Started..." & vbCrLf
                Me.tmOut6.Enabled = False
                Me.tmOut6.Enabled = True
            End If
            
        End If
    End With
    
End Sub

'PERMISSSION TIMER....
Private Sub wsCl1_SendComplete()
flgWaitWs1 = False
End Sub
Private Sub wsCl2_SendComplete()
flgWaitWs2 = False
End Sub
Private Sub wsCl3_SendComplete()
flgWaitWs3 = False
End Sub
Private Sub wsCl4_SendComplete()
flgWaitWs4 = False
End Sub
Private Sub wsCl5_SendComplete()
flgWaitWs5 = False
End Sub
Private Sub wsCl6_SendComplete()
flgWaitWs6 = False
End Sub
Private Sub wsCl7_SendComplete()
flgWaitWs7 = False
End Sub
Private Sub wsCl8_SendComplete()
flgWaitWs8 = False
End Sub

'SENDPROGRESS
'Private Sub wsCl1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'tellWS 0, bytesSent & " Bytes Sent to " & wsCl1.RemoteHost
'End Sub
'Private Sub wsCl2_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'tellWS 1, bytesSent & " Bytes Sent to " & wsCl2.RemoteHost
'End Sub
'Private Sub wsCl3_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'tellWS 2, bytesSent & " Bytes Sent to " & wsCl3.RemoteHost
'End Sub
'Private Sub wsCl4_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'tellWS 3, bytesSent & " Bytes Sent to " & wsCl4.RemoteHost
'End Sub
'Private Sub wsCl5_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'tellWS 4, bytesSent & " Bytes Sent to " & wsCl5.RemoteHost
'End Sub
'Private Sub wsCl6_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'tellWS 5, bytesSent & " Bytes Sent to " & wsCl6.RemoteHost
'End Sub


'SUBS AND FUNCTIONS FOR ALL
'mostra titulo do frame...
Private Sub tellWS(fme As Integer, arg As String)
Me.fmePc(fme).Caption = arg$
End Sub

Private Sub RestartWS(Index As Integer, Optional restart As Boolean = True)

Dim ws As Object
For Each ws In frmsv2
DoEvents
''Debug.print ws.Name
If ws.Name = "wsCl" & Index Then
    ws.Close
    pci(Index).state = "OFF"
    If restart = True Then
    ws.Listen
    End If
    
Exit For
End If

Next

End Sub
'////////////////////////////////////////////////////ONE BY ONE
'INTERPRET DATA
Private Sub InterpretData1(dta As String)
Dim td$
With pci(1)
dta$ = Replace(dta, vbCrLf, "")
''Debug.print "Recv " & dta
    Select Case UCase(Left(dta, 4))
        Case Is = "NEWM"
            Me.tmOut1.Enabled = False
            flgWaitWs1 = False
            dta = Replace(dta, vbCrLf, "")
                .pcName = Right(dta, Len(dta) - 4)
            tellWS 0, .pcName & " Online..."
            .state = "LOCKOFF"
            'send configurations here
            '-----------------
            If wsCl1.state = 7 Then
                wsCl1.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
            '---------------------------------------------------------
            
          Case Is = "KILL"
            .state = "OFF"
            If wsCl1.state = 7 Then
            wsCl1.SendData "OKOUT" & vbCrLf
            RestartWS 1
            End If
            'from here
          Case "CFG?"
            If wsCl1.state = 7 Then
                wsCl1.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case "CODE", "CHEK"
          'ver se está a ser usado por algum cliente...
          Dim tc$
          tc$ = Right(dta, Len(dta) - 4)
          tellWS 0, "Validating " & tc$
            If isCodeBUSY(tc$, 1) = True Then
                If wsCl1.state = 7 Then
                    wsCl1.SendData "BUSY"
                End If
            
            Else 'validar o codigo
               If wsCl1.state = 7 Then
                   'wsCl1.SendData
                   wsCl1.SendData valCod(tc$, 1)
               End If
            End If
          Case "DTLS"
          
          td$ = Right(dta$, Len(dta$) - 4)
          'Exit Sub
          Debug.Print td$
          If Len(td$) > 88 Then td = Left(td, 88)
            If Len(td$) = 88 Then
                CopyMemory tmde(1), ByVal td$, 88
    
                    .pcuTime = Left(tmde(1).huso, 8)
                    .pcuPrice = Format(CDbl(Right(tmde(1).huso, 4) * 100), "0$00")
                    .netTotal = trasnBytes(CLng(Val(Trim$(tmde(1).netc))))
                    .netNow = trasnBytes(CLng(Trim$(tmde(1).netn)))
                    .netPrice = IIf(tmde(1).netp <> "Cybero", Format(CDbl("0" & Val(tmde(1).netp)) * 100, "0$00"), "Cyber Offset")
                    
                    '///////////////////////////////
                    .balTotal = Format(CLng(tmde(1).tbal) * 100, "0$00")
                    .balUsed = Format((tmde(1).totu) * 100, "0$00")
                    .balRemain = Format((CLng(CLng(tmde(1).tbal) - tmde(1).totu)) * 100, "0$00")
                    '///////////////////////////////
                    
            'Debug.Print
            Debug.Print Mid(td$, 25, 51)
            addUpdateCard mCards(1).code, Mid(td$, 25, 51)
            If CLng(tmde(1).TLogoff) = 1 Then
            'logging off
            
                  Dim dtl As DETAILS
                  .logoff = Format(Now, "dd/mm/yy hh:mm:ss")
                  .state = "FREE"
                  dtl.pc = pci(1).pcName
                  dtl.din = pci(1).login
                  dtl.dout = pci(1).logoff
                  dtl.tmv = pci(1).pcuTime

                  With mCards(1)
                      dtl.data = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
                  End With

                  addDetails dtl

            End If
           End If
           Case "USER"
           .login = Format(Now, "dd/mm/yy hh:mm:ss")
           .state = "BUSY"
           
           Case "TYPE"
            myInt = 0
    End Select
End With

End Sub

'DOIS CLIENTE ...

'INTERPRET DATA
Private Sub InterpretData2(dta As String)
Dim td$
With pci(2)
dta$ = Replace(dta, vbCrLf, "")
''Debug.print "Recv " & dta
    Select Case UCase(Left(dta, 4))
        Case Is = "NEWM"
            Me.tmOut2.Enabled = False
            flgWaitWs2 = False
            dta = Replace(dta, vbCrLf, "")
                .pcName = Right(dta, Len(dta) - 4)
            tellWS 1, .pcName & " Online..."
           .state = "LOCKOFF"
            'send configurations here
            If wsCl2.state = 7 Then
                wsCl2.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case Is = "KILL"
            .state = "OFF"
            If wsCl2.state = 7 Then
            wsCl2.SendData "OKOUT" & vbCrLf
            RestartWS 2
            End If
            'from here
          Case "CFG?"
            If wsCl2.state = 7 Then
                wsCl2.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case "CODE", "CHEK"
          'ver se está a ser usado por algum cliente...
          Dim tc$
          tc$ = Right(dta, Len(dta) - 4)
          tellWS 1, "Validating " & tc$
            If isCodeBUSY(tc$, 2) = True Then
                If wsCl2.state = 7 Then
                    wsCl2.SendData "BUSY"
                End If
            
            Else
                    'validar o codigo
               If wsCl2.state = 7 Then
                   'wsCl1.SendData
                   wsCl2.SendData valCod(tc$, 2)
               End If
               
                    
            End If
          
          Case "DTLS"
          td$ = Right(dta$, Len(dta$) - 4)
          'Exit Sub
          Debug.Print td$
          If Len(td$) > 88 Then td = Left(td, 88)
            If Len(td$) = 88 Then
                CopyMemory tmde(2), ByVal td$, 88
    
                    .pcuTime = Left(tmde(2).huso, 8)
                    .pcuPrice = Format(CDbl(Right(tmde(2).huso, 4) * 100), "0$00")
                    .netTotal = trasnBytes(CLng(Val(Trim$(tmde(2).netc))))
                    .netNow = trasnBytes(CLng(Trim$(tmde(2).netn)))
                    .netPrice = IIf(tmde(2).netp <> "Cybero", Format(CDbl("0" & Val(tmde(2).netp)) * 100, "0$00"), "Cyber Offset")
                    
                    '///////////////////////////////
                    .balTotal = Format(CLng(tmde(2).tbal) * 100, "0$00")
                    .balUsed = Format((tmde(2).totu) * 100, "0$00")
                    .balRemain = Format((CLng(CLng(tmde(2).tbal) - tmde(2).totu)) * 100, "0$00")
                    '///////////////////////////////
                    
            'Debug.Print
            Debug.Print Mid(td$, 25, 51)
            addUpdateCard mCards(2).code, Mid(td$, 25, 51)
            If CLng(tmde(2).TLogoff) = 1 Then
            'logging off
            
                  Dim dtl As DETAILS
                  .logoff = Format(Now, "dd/mm/yy hh:mm:ss")
                  .state = "FREE"
                  dtl.pc = pci(2).pcName
                  dtl.din = pci(2).login
                  dtl.dout = pci(2).logoff
                  dtl.tmv = pci(2).pcuTime

                  With mCards(2)
                      dtl.data = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
                  End With

                  addDetails dtl

            End If
           End If
           Case "USER"
           '.clientID = Trim$(Right(dta$, Len(dta$) - 4))
           .login = Format(Now, "dd/mm/yy hh:mm:ss")
           .state = "BUSY"
           Case "TYPE"
           myInt = 0
           
    End Select
End With

End Sub

'TERCEIRO CLIENTE
Private Sub InterpretData3(dta As String)
Dim td$
With pci(3)
dta$ = Replace(dta, vbCrLf, "")
''Debug.print "Recv " & dta
    Select Case UCase(Left(dta, 4))
        Case Is = "NEWM"
            Me.tmOut3.Enabled = False
            flgWaitWs3 = False
            dta = Replace(dta, vbCrLf, "")
                .pcName = Right(dta, Len(dta) - 4)
            tellWS 2, .pcName & " Online..."
             .state = "LOCKOFF"
            'send configurations here
            If wsCl3.state = 7 Then
                wsCl3.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case Is = "KILL"
            .state = "OFF"
            If wsCl3.state = 7 Then
            wsCl3.SendData "OKOUT" & vbCrLf
            RestartWS 3
            End If
            'from here
          Case "CFG?"
            If wsCl3.state = 7 Then
                wsCl3.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case "CODE"
          'ver se está a ser usado por algum cliente...
          Dim tc$
          tc$ = Right(dta, Len(dta) - 4)
          tellWS 2, "Validating " & tc$
            If isCodeBUSY(tc$, 3) = True Then
                If wsCl3.state = 7 Then
                    wsCl3.SendData "BUSY"
                End If
            
            Else
                    'validar o codigo
               If wsCl3.state = 7 Then
                   'wsCl1.SendData
                   wsCl3.SendData valCod(tc$, 3)
               End If
               
                    
            End If
          
          Case "DTLS"
          td$ = Right(dta$, Len(dta$) - 4)
          'Exit Sub
          Debug.Print td$
          If Len(td$) > 88 Then td = Left(td, 88)
            If Len(td$) = 88 Then
                CopyMemory tmde(3), ByVal td$, 88
    
                    .pcuTime = Left(tmde(3).huso, 8)
                    .pcuPrice = Format(CDbl(Right(tmde(3).huso, 4) * 100), "0$00")
                    .netTotal = trasnBytes(CLng(Val(Trim$(tmde(3).netc))))
                    .netNow = trasnBytes(CLng(Trim$(tmde(3).netn)))
                    .netPrice = IIf(tmde(3).netp <> "Cybero", Format(CDbl("0" & Val(tmde(3).netp)) * 100, "0$00"), "Cyber Offset")
                    
                    '///////////////////////////////
                    .balTotal = Format(CLng(tmde(3).tbal) * 100, "0$00")
                    .balUsed = Format((tmde(3).totu) * 100, "0$00")
                    .balRemain = Format((CLng(CLng(tmde(3).tbal) - tmde(3).totu)) * 100, "0$00")
                    '///////////////////////////////
                    
            'Debug.Print
            Debug.Print Mid(td$, 25, 51)
            addUpdateCard mCards(3).code, Mid(td$, 25, 51)
            If CLng(tmde(3).TLogoff) = 1 Then
            'logging off
            
                  Dim dtl As DETAILS
                  .logoff = Format(Now, "dd/mm/yy hh:mm:ss")
                  .state = "FREE"
                  dtl.pc = pci(3).pcName
                  dtl.din = pci(3).login
                  dtl.dout = pci(3).logoff
                  dtl.tmv = pci(3).pcuTime

                  With mCards(3)
                      dtl.data = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
                  End With

                  addDetails dtl

            End If
           End If
           Case "USER"
           'clientID = Trim$(Right(dta$, Len(dta$) - 4))
           .login = Format(Now, "dd/mm/yy hh:mm:ss")
           .logoff = ""
           .state = "BUSY"
           Case "TYPE"
            myInt = 0
    End Select
End With

End Sub


'QUARTO CLIENTE
Private Sub InterpretData4(dta As String)
Dim td$
With pci(4)
dta$ = Replace(dta, vbCrLf, "")
''Debug.print "Recv " & dta
    Select Case UCase(Left(dta, 4))
        Case Is = "NEWM"
            Me.tmOut4.Enabled = False
            flgWaitWs4 = False
            dta = Replace(dta, vbCrLf, "")
                .pcName = Right(dta, Len(dta) - 4)
            tellWS 3, .pcName & " Online..."
             .state = "LOCKOFF"
            'send configurations here
            If wsCl4.state = 7 Then
                wsCl4.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case Is = "KILL"
            .state = "OFF"
            If wsCl4.state = 7 Then
            wsCl4.SendData "OKOUT" & vbCrLf
            RestartWS 4
            End If
            'from here
          Case "CFG?"
            If wsCl4.state = 7 Then
                wsCl4.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case "CODE"
          'ver se está a ser usado por algum cliente...
          Dim tc$
          tc$ = Right(dta, Len(dta) - 4)
          tellWS 3, "Validating " & tc$
            If isCodeBUSY(tc$, 4) = True Then
                If wsCl4.state = 7 Then
                    wsCl4.SendData "BUSY"
                End If
            
            Else
                    'validar o codigo
               If wsCl4.state = 7 Then
                   'wsCl1.SendData
                   wsCl4.SendData valCod(tc$, 4)
               End If
               
                    
            End If
        
          Case "DTLS"
          td$ = Right(dta$, Len(dta$) - 4)
          'Exit Sub
          Debug.Print td$
          If Len(td$) > 88 Then td = Left(td, 88)
            If Len(td$) = 88 Then
                CopyMemory tmde(4), ByVal td$, 88
    
                    .pcuTime = Left(tmde(4).huso, 8)
                    .pcuPrice = Format(CDbl(Right(tmde(4).huso, 4) * 100), "0$00")
                    .netTotal = trasnBytes(CLng(Val(Trim$(tmde(4).netc))))
                    .netNow = trasnBytes(CLng(Trim$(tmde(4).netn)))
                    .netPrice = IIf(tmde(4).netp <> "Cybero", Format(CDbl("0" & Val(tmde(4).netp)) * 100, "0$00"), "Cyber Offset")
                    
                    '///////////////////////////////
                    .balTotal = Format(CLng(tmde(4).tbal) * 100, "0$00")
                    .balUsed = Format((tmde(4).totu) * 100, "0$00")
                    .balRemain = Format((CLng(CLng(tmde(4).tbal) - tmde(4).totu)) * 100, "0$00")
                    '///////////////////////////////
                    
            'Debug.Print
            Debug.Print Mid(td$, 25, 51)
            addUpdateCard mCards(4).code, Mid(td$, 25, 51)
            If CLng(tmde(4).TLogoff) = 1 Then
            'logging off
            
                  Dim dtl As DETAILS
                  .logoff = Format(Now, "dd/mm/yy hh:mm:ss")
                  .state = "FREE"
                  dtl.pc = pci(4).pcName
                  dtl.din = pci(4).login
                  dtl.dout = pci(4).logoff
                  dtl.tmv = pci(4).pcuTime

                  With mCards(4)
                      dtl.data = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
                  End With

                  addDetails dtl

            End If
           End If
           Case "USER"
           '.clientID = Trim$(Right(dta$, Len(dta$) - 4))
           .login = Format(Now, "dd/mm/yy hh:mm:ss")
           .logoff = ""
           .state = "BUSY"
           Case "TYPE"
            myInt = 0
    End Select
End With

End Sub


'QUINTO CLIENTE
Private Sub InterpretData5(dta As String)
Dim td$
With pci(5)
dta$ = Replace(dta, vbCrLf, "")
''Debug.print "Recv " & dta
    Select Case UCase(Left(dta, 4))
        Case Is = "NEWM"
            Me.tmOut5.Enabled = False
            flgWaitWs5 = False
            dta = Replace(dta, vbCrLf, "")
                .pcName = Right(dta, Len(dta) - 4)
            tellWS 4, .pcName & " Online..."
             .state = "LOCKOFF"
            'send configurations here
            If wsCl5.state = 7 Then
                wsCl5.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case Is = "KILL"
            .state = "OFF"
            If wsCl5.state = 7 Then
            wsCl5.SendData "OKOUT" & vbCrLf
            RestartWS 5
            End If
            'from here
          Case "CFG?"
            If wsCl5.state = 7 Then
                wsCl5.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case "CODE"
          'ver se está a ser usado por algum cliente...
          Dim tc$
          tc$ = Right(dta, Len(dta) - 4)
          tellWS 4, "Validating " & tc$
            If isCodeBUSY(tc$, 5) = True Then
                If wsCl5.state = 7 Then
                    wsCl5.SendData "BUSY"
                End If
            
            Else
                    'validar o codigo
               If wsCl5.state = 7 Then
                   'wsCl1.SendData
                   wsCl5.SendData valCod(tc$, 5)
               End If
               
                    
            End If
       
          Case "DTLS"
          td$ = Right(dta$, Len(dta$) - 4)
          'Exit Sub
          Debug.Print td$
          If Len(td$) > 88 Then td = Left(td, 88)
            If Len(td$) = 88 Then
                CopyMemory tmde(5), ByVal td$, 88
    
                    .pcuTime = Left(tmde(5).huso, 8)
                    .pcuPrice = Format(CDbl(Right(tmde(5).huso, 4) * 100), "0$00")
                    .netTotal = trasnBytes(CLng(Val(Trim$(tmde(5).netc))))
                    .netNow = trasnBytes(CLng(Trim$(tmde(5).netn)))
                    .netPrice = IIf(tmde(5).netp <> "Cybero", Format(CDbl("0" & Val(tmde(5).netp)) * 100, "0$00"), "Cyber Offset")
                    
                    '///////////////////////////////
                    .balTotal = Format(CLng(tmde(5).tbal) * 100, "0$00")
                    .balUsed = Format((tmde(5).totu) * 100, "0$00")
                    .balRemain = Format((CLng(CLng(tmde(5).tbal) - tmde(5).totu)) * 100, "0$00")
                    '///////////////////////////////
                    
            'Debug.Print
            Debug.Print Mid(td$, 25, 51)
            addUpdateCard mCards(5).code, Mid(td$, 25, 51)
            If CLng(tmde(5).TLogoff) = 1 Then
            'logging off
            
                  Dim dtl As DETAILS
                  .logoff = Format(Now, "dd/mm/yy hh:mm:ss")
                  .state = "FREE"
                  dtl.pc = pci(5).pcName
                  dtl.din = pci(5).login
                  dtl.dout = pci(5).logoff
                  dtl.tmv = pci(5).pcuTime

                  With mCards(5)
                      dtl.data = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
                  End With

                  addDetails dtl

            End If
           End If
           Case "USER"
           '.clientID = Trim$(Right(dta$, Len(dta$) - 4))
           .login = Format(Now, "dd/mm/yy hh:mm:ss")
           .logoff = ""
           .state = "BUSY"
           Case "TYPE"
            myInt = 0
    End Select
End With

End Sub


'SEXTO CLIENTE
Private Sub InterpretData6(dta As String)
Dim td$
With pci(6)
dta$ = Replace(dta, vbCrLf, "")
''Debug.print "Recv " & dta
    Select Case UCase(Left(dta, 4))
        Case Is = "NEWM"
            Me.tmOut6.Enabled = False
            flgWaitWs6 = False
            dta = Replace(dta, vbCrLf, "")
                .pcName = Right(dta, Len(dta) - 4)
            tellWS 5, .pcName & " Online..."
            .state = "LOCKOFF"
            'send configurations here
            If wsCl6.state = 7 Then
                wsCl6.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case Is = "KILL"
            .state = "OFF"
            If wsCl6.state = 7 Then
            wsCl6.SendData "OKOUT" & vbCrLf
            RestartWS 6
            End If
            'from here
          Case "CFG?"
            If wsCl6.state = 7 Then
                wsCl6.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case "CODE"
          'ver se está a ser usado por algum cliente...
          Dim tc$
          tc$ = Right(dta, Len(dta) - 4)
          tellWS 5, "Validating " & tc$
            If isCodeBUSY(tc$, 6) = True Then
                If wsCl6.state = 7 Then
                    wsCl6.SendData "BUSY"
                End If
            
            Else
                    'validar o codigo
               If wsCl6.state = 7 Then
                   'wsCl1.SendData
                   wsCl6.SendData valCod(tc$, 6)
               End If
               
                    
            End If
  
          Case "DTLS"
          td$ = Right(dta$, Len(dta$) - 4)
          'Exit Sub
          Debug.Print td$
          If Len(td$) > 88 Then td = Left(td, 88)
            If Len(td$) = 88 Then
                CopyMemory tmde(6), ByVal td$, 88
    
                    .pcuTime = Left(tmde(6).huso, 8)
                    .pcuPrice = Format(CDbl(Right(tmde(6).huso, 4) * 100), "0$00")
                    .netTotal = trasnBytes(CLng(Val(Trim$(tmde(6).netc))))
                    .netNow = trasnBytes(CLng(Trim$(tmde(6).netn)))
                    .netPrice = IIf(tmde(6).netp <> "Cybero", Format(CDbl("0" & Val(tmde(6).netp)) * 100, "0$00"), "Cyber Offset")
                    
                    '///////////////////////////////
                    .balTotal = Format(CLng(tmde(6).tbal) * 100, "0$00")
                    .balUsed = Format((tmde(6).totu) * 100, "0$00")
                    .balRemain = Format((CLng(CLng(tmde(6).tbal) - tmde(6).totu)) * 100, "0$00")
                    '///////////////////////////////
                    
            'Debug.Print
            Debug.Print Mid(td$, 25, 51)
            addUpdateCard mCards(6).code, Mid(td$, 25, 51)
            If CLng(tmde(6).TLogoff) = 1 Then
            'logging off
            
                  Dim dtl As DETAILS
                  .logoff = Format(Now, "dd/mm/yy hh:mm:ss")
                  .state = "FREE"
                  dtl.pc = pci(6).pcName
                  dtl.din = pci(6).login
                  dtl.dout = pci(6).logoff
                  dtl.tmv = pci(6).pcuTime

                  With mCards(6)
                      dtl.data = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
                  End With

                  addDetails dtl

            End If
           End If
           Case "USER"
           '.clientID = Trim$(Right(dta$, Len(dta$) - 4))
           .login = Format(Now, "dd/mm/yy hh:mm:ss")
           .logoff = ""
           .state = "BUSY"
           Case "TYPE"
            myInt = 0
    End Select
End With

End Sub

'Setimo cliente
Private Sub InterpretData7(dta As String)
Dim td$
With pci(7)
dta$ = Replace(dta, vbCrLf, "")
''Debug.print "Recv " & dta
    Select Case UCase(Left(dta, 4))
        Case Is = "NEWM"
            Me.tmOut7.Enabled = False
            flgWaitWs7 = False
            dta = Replace(dta, vbCrLf, "")
                .pcName = Right(dta, Len(dta) - 4)
            tellWS 6, .pcName & " Online..."
            .state = "LOCKOFF"
            'send configurations here
            If wsCl7.state = 7 Then
                wsCl7.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case Is = "KILL"
            .state = "OFF"
            If wsCl7.state = 7 Then
            wsCl7.SendData "OKOUT" & vbCrLf
            RestartWS 7
            End If
            'from here
          Case "CFG?"
            If wsCl7.state = 7 Then
                wsCl7.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case "CODE"
          'ver se está a ser usado por algum cliente...
          Dim tc$
          tc$ = Right(dta, Len(dta) - 4)
          tellWS 6, "Validating " & tc$
            If isCodeBUSY(tc$, 7) = True Then
                If wsCl7.state = 7 Then
                    wsCl7.SendData "BUSY"
                End If
            
            Else
                    'validar o codigo
               If wsCl7.state = 7 Then
                   'wsCl1.SendData
                   wsCl7.SendData valCod(tc$, 7)
               End If
               
                    
            End If
  
          Case "DTLS"
          td$ = Right(dta$, Len(dta$) - 4)
          'Exit Sub
          Debug.Print td$
          If Len(td$) > 88 Then td = Left(td, 88)
            If Len(td$) = 88 Then
                CopyMemory tmde(7), ByVal td$, 88
    
                    .pcuTime = Left(tmde(7).huso, 8)
                    .pcuPrice = Format(CDbl(Right(tmde(7).huso, 4) * 100), "0$00")
                    .netTotal = trasnBytes(CLng(Val(Trim$(tmde(7).netc))))
                    .netNow = trasnBytes(CLng(Trim$(tmde(7).netn)))
                    .netPrice = IIf(tmde(7).netp <> "Cybero", Format(CDbl("0" & Val(tmde(7).netp)) * 100, "0$00"), "Cyber Offset")
                    
                    '///////////////////////////////
                    .balTotal = Format(CLng(tmde(7).tbal) * 100, "0$00")
                    .balUsed = Format((tmde(7).totu) * 100, "0$00")
                    .balRemain = Format((CLng(CLng(tmde(7).tbal) - tmde(7).totu)) * 100, "0$00")
                    '///////////////////////////////
                    
            'Debug.Print
            Debug.Print Mid(td$, 25, 51)
            addUpdateCard mCards(7).code, Mid(td$, 25, 51)
            If CLng(tmde(7).TLogoff) = 1 Then
            'logging off
            
                  Dim dtl As DETAILS
                  .logoff = Format(Now, "dd/mm/yy hh:mm:ss")
                  .state = "FREE"
                  dtl.pc = pci(7).pcName
                  dtl.din = pci(7).login
                  dtl.dout = pci(7).logoff
                  dtl.tmv = pci(7).pcuTime

                  With mCards(7)
                      dtl.data = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
                  End With

                  addDetails dtl

            End If
           End If
           Case "USER"
           '.clientID = Trim$(Right(dta$, Len(dta$) - 4))
           .login = Format(Now, "dd/mm/yy hh:mm:ss")
           .logoff = ""
           .state = "BUSY"
           Case "TYPE"
            myInt = 0
    End Select
End With

End Sub

'Oitavo cliente
'Setimo cliente
Private Sub InterpretData8(dta As String)
Dim td$
With pci(8)
dta$ = Replace(dta, vbCrLf, "")
''Debug.print "Recv " & dta
    Select Case UCase(Left(dta, 4))
        Case Is = "NEWM"
            Me.tmOut8.Enabled = False
            flgWaitWs8 = False
            dta = Replace(dta, vbCrLf, "")
                .pcName = Right(dta, Len(dta) - 4)
            tellWS 7, .pcName & " Online..."
            .state = "LOCKOFF"
            'send configurations here
            If wsCl8.state = 7 Then
                wsCl8.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case Is = "KILL"
            .state = "OFF"
            If wsCl8.state = 7 Then
            wsCl8.SendData "OKOUT" & vbCrLf
            RestartWS 8
            End If
            'from here
          Case "CFG?"
            If wsCl8.state = 7 Then
                wsCl8.SendData "HEIS" & Format(CLng(getPrices.Pnet), "0000") & _
                Format(CLng(getPrices.Pwindows), "0000") & _
                Format(CLng(getPrices.offSet), "0000") & Format(cCyberXV2FLG, "0000") & vbCrLf
            End If
          Case "CODE"
          'ver se está a ser usado por algum cliente...
          Dim tc$
          tc$ = Right(dta, Len(dta) - 4)
          tellWS 7, "Validating " & tc$
            If isCodeBUSY(tc$, 8) = True Then
                If wsCl8.state = 7 Then
                    wsCl8.SendData "BUSY"
                End If
            
            Else
                    'validar o codigo
               If wsCl8.state = 7 Then
                   'wsCl1.SendData
                   wsCl8.SendData valCod(tc$, 8)
               End If
               
                    
            End If
  
          Case "DTLS"
          td$ = Right(dta$, Len(dta$) - 4)
          'Exit Sub
          Debug.Print td$
          If Len(td$) > 88 Then td = Left(td, 88)
            If Len(td$) = 88 Then
                CopyMemory tmde(8), ByVal td$, 88
    
                    .pcuTime = Left(tmde(8).huso, 8)
                    .pcuPrice = Format(CDbl(Right(tmde(8).huso, 4) * 100), "0$00")
                    .netTotal = trasnBytes(CLng(Val(Trim$(tmde(8).netc))))
                    .netNow = trasnBytes(CLng(Trim$(tmde(8).netn)))
                    .netPrice = IIf(tmde(8).netp <> "Cybero", Format(CDbl("0" & Val(tmde(8).netp)) * 100, "0$00"), "Cyber Offset")
                    
                    '///////////////////////////////
                    .balTotal = Format(CLng(tmde(8).tbal) * 100, "0$00")
                    .balUsed = Format((tmde(8).totu) * 100, "0$00")
                    .balRemain = Format((CLng(CLng(tmde(8).tbal) - tmde(8).totu)) * 100, "0$00")
                    '///////////////////////////////
                    
            'Debug.Print
            Debug.Print Mid(td$, 25, 51)
            addUpdateCard mCards(8).code, Mid(td$, 25, 51)
            If CLng(tmde(8).TLogoff) = 1 Then
            'logging off
            
                  Dim dtl As DETAILS
                  .logoff = Format(Now, "dd/mm/yy hh:mm:ss")
                  .state = "FREE"
                  dtl.pc = pci(8).pcName
                  dtl.din = pci(8).login
                  dtl.dout = pci(8).logoff
                  dtl.tmv = pci(8).pcuTime

                  With mCards(8)
                      dtl.data = .id & .code & .date & .life & .flag & .tbal & .tusd & .bytes
                  End With

                  addDetails dtl

            End If
           End If
           Case "USER"
           '.clientID = Trim$(Right(dta$, Len(dta$) - 4))
           .login = Format(Now, "dd/mm/yy hh:mm:ss")
           .logoff = ""
           .state = "BUSY"
           Case "TYPE"
            myInt = 0
    End Select
End With

End Sub


'////////////////////////////////////////////////////ALL
'limpa os dados da lista
Private Sub ClearList(lst As ListView)
Dim it%
    For it% = 1 To lst.ListItems.Count
    DoEvents
        lst.ListItems(it%).SubItems(1) = ""
    Next
    
End Sub

'ver se cliente esta a usar o mesmo codigo
Private Function isCodeBUSY(ByVal cod$, Index&) As Boolean
Dim it&
For it& = 1 To 8
    If Index <> it Then
    Debug.Print mCards(it).code
        If mCards(it).code = cod$ And pci(it).state = "BUSY" Then
        Call alertSocket(it, frmsv2)
                isCodeBUSY = True
                Exit Function
        End If
    End If
    
Next

isCodeBUSY = False
End Function

Private Function alertSocket(ByVal si&, frm As Form)
Dim ws As Object

For Each ws In frm
    If ws.Name = "wsCl" & si& Then
    If ws.state = 7 Then
        ws.SendData "BCOD" & "Someone is trying to use your code. Alert sent " & Format(Now, "hh:mm:ss")
        Exit Function
    End If
    Exit For
    End If
Next

End Function



Public Sub displayPrices()
Dim lbObj, ig&, tmplb

For ig = 1 To 8
    For Each lbObj In frmsv2
    
    If TypeName(lbObj) = "Label" Then
      If Left(lbObj.Name, 4) = "lbi" & ig Then
      
      If lbObj.Index = 3 Then
        lbObj.Caption = Format(prcSetup.Pwindows * 100, "0$00") & " Per Hour"  'preço windows
      End If
        
      If lbObj.Index = 6 Then
        lbObj.Caption = Format(prcSetup.Pnet * 100, "0$00") & " Per MB [" & Format(prcSetup.offSet / 2, "-0.00") & "]"  'preço internet offset
      End If
      
      
      End If
     End If
    Next
   
    
Next
End Sub
