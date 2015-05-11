VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmSup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cards Manager PacketXV3"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   Icon            =   "frmSup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmSup.frx":030A
   ScaleHeight     =   7440
   ScaleWidth      =   10680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmTime 
      Interval        =   2800
      Left            =   3300
      Top             =   6990
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   4290
      Top             =   0
   End
   Begin VB.TextBox txtHost 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1260
      TabIndex        =   5
      ToolTipText     =   "Press Enter To Save Changes"
      Top             =   60
      Width           =   1875
   End
   Begin VB.CommandButton cmdConUn 
      Caption         =   "&Connect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   60
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   150
      TabIndex        =   0
      Top             =   570
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   11139
      _Version        =   393216
      TabOrientation  =   3
      TabHeight       =   520
      BackColor       =   16711680
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "1 Generation"
      TabPicture(0)   =   "frmSup.frx":FAA5C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbPrice"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Shape1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Shape1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbc"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdGen"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "List1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lstc"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdCard(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdCard(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdCard(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lhm(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lhm(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "2 Transport"
      TabPicture(1)   =   "frmSup.frx":FAA78
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(3)=   "Shape2"
      Tab(1).Control(4)=   "Shape3"
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(6)=   "Label11"
      Tab(1).Control(7)=   "Label12"
      Tab(1).Control(8)=   "Label13(0)"
      Tab(1).Control(9)=   "Label14(0)"
      Tab(1).Control(10)=   "Label13(1)"
      Tab(1).Control(11)=   "Label14(1)"
      Tab(1).Control(12)=   "lscard"
      Tab(1).Control(13)=   "lcData"
      Tab(1).Control(14)=   "lbDet"
      Tab(1).Control(15)=   "cmdTrans(3)"
      Tab(1).Control(16)=   "cmdTrans(0)"
      Tab(1).Control(17)=   "cmdTrans(1)"
      Tab(1).Control(18)=   "cmdTrans(2)"
      Tab(1).Control(19)=   "pgb"
      Tab(1).Control(20)=   "txtFind"
      Tab(1).Control(21)=   "Frame2"
      Tab(1).Control(22)=   "Frame3"
      Tab(1).Control(23)=   "cmdTrans(4)"
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "3 Manager"
      TabPicture(2)   =   "frmSup.frx":FAA94
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdTrans 
         Caption         =   "&Read V1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -70800
         TabIndex        =   58
         Top             =   4620
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2475
         Left            =   -69210
         TabIndex        =   47
         Top             =   3720
         Visible         =   0   'False
         Width           =   3765
         Begin VB.CommandButton cmdWhatDo 
            Caption         =   "&Save As"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2430
            TabIndex        =   55
            Top             =   420
            Width           =   1155
         End
         Begin VB.CommandButton cmdWhatDo 
            Caption         =   "&Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   54
            Top             =   420
            Width           =   1005
         End
         Begin VB.CommandButton cmdWhatDo 
            Caption         =   "&Remove"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   53
            Top             =   420
            Width           =   1275
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " What do you want  to do, with selected cards."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   150
            TabIndex        =   48
            Top             =   120
            Width           =   3405
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2565
         Left            =   -69210
         TabIndex        =   42
         Top             =   930
         Width           =   3765
         Begin VB.OptionButton opsearch 
            Caption         =   "All Cards"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   90
            TabIndex        =   64
            Top             =   300
            Width           =   1125
         End
         Begin VB.CheckBox chkBal 
            Appearance      =   0  'Flat
            Caption         =   "Use for Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   60
            Top             =   1140
            Width           =   1755
         End
         Begin VB.OptionButton opsearch 
            Caption         =   "Code and Data.Code don´t mach"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   90
            TabIndex        =   56
            Top             =   1710
            Width           =   3345
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "&Select"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   2070
            Width           =   1275
         End
         Begin VB.OptionButton opsearch 
            Caption         =   "New Cards?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   1410
            TabIndex        =   51
            Top             =   600
            Width           =   1515
         End
         Begin VB.OptionButton opsearch 
            Caption         =   "Used Cards?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   1410
            TabIndex        =   50
            Top             =   330
            Width           =   1515
         End
         Begin VB.OptionButton opsearch 
            Caption         =   "Price = ?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   46
            Top             =   1410
            Width           =   1545
         End
         Begin VB.OptionButton opsearch 
            Caption         =   "Price < ?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   45
            Top             =   1140
            Width           =   1575
         End
         Begin VB.OptionButton opsearch 
            Caption         =   "Price > ?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   44
            Top             =   840
            Width           =   1575
         End
         Begin VB.OptionButton opsearch 
            Caption         =   "By Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   43
            Top             =   570
            Width           =   1125
         End
         Begin VB.Line Line4 
            X1              =   1680
            X2              =   1830
            Y1              =   1230
            Y2              =   1230
         End
         Begin VB.Line Line3 
            X1              =   1830
            X2              =   1830
            Y1              =   930
            Y2              =   1560
         End
         Begin VB.Line Line2 
            X1              =   1650
            X2              =   1830
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line1 
            X1              =   1680
            X2              =   1830
            Y1              =   930
            Y2              =   930
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "How do you want Selection Result?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   60
            TabIndex        =   49
            Top             =   120
            Width           =   2595
         End
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -69210
         MaxLength       =   19
         TabIndex        =   41
         Top             =   570
         Width           =   2625
      End
      Begin VB.Frame Frame1 
         Caption         =   " Data Base "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1965
         Left            =   -67800
         TabIndex        =   36
         Top             =   4140
         Width           =   2385
         Begin VB.CommandButton cmdRebd 
            Caption         =   "&Server Cards"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   90
            TabIndex        =   57
            Top             =   1350
            Width           =   1755
         End
         Begin VB.CommandButton cmdRebd 
            Caption         =   "R&ebuild Server"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   90
            TabIndex        =   39
            Top             =   990
            Width           =   1755
         End
         Begin VB.CommandButton cmdRebd 
            Caption         =   "&Rebuild Local"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   90
            TabIndex        =   38
            Top             =   630
            Width           =   1755
         End
         Begin VB.CommandButton cmdRebd 
            Caption         =   "&Backup Local"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   90
            TabIndex        =   37
            Top             =   270
            Width           =   1755
         End
      End
      Begin MSComctlLib.ProgressBar pgb 
         Height          =   105
         Left            =   -72510
         TabIndex        =   31
         Top             =   5520
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.CommandButton cmdTrans 
         Caption         =   "&Export"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -70800
         TabIndex        =   30
         Top             =   4260
         Width           =   1455
      End
      Begin VB.CommandButton cmdTrans 
         Caption         =   "&Import"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -70800
         TabIndex        =   29
         Top             =   3900
         Width           =   1455
      End
      Begin VB.CommandButton cmdTrans 
         Caption         =   "&Load"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -70800
         TabIndex        =   28
         Top             =   3540
         Width           =   1455
      End
      Begin VB.CommandButton cmdTrans 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -70800
         TabIndex        =   27
         Top             =   3180
         Width           =   1455
      End
      Begin VB.TextBox lbDet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -72540
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   570
         Width           =   3225
      End
      Begin VB.ListBox lcData 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         ItemData        =   "frmSup.frx":FAAB0
         Left            =   -74910
         List            =   "frmSup.frx":FAAB2
         TabIndex        =   24
         Top             =   5910
         Width           =   5595
      End
      Begin VB.ListBox lscard 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5310
         ItemData        =   "frmSup.frx":FAAB4
         Left            =   -74910
         List            =   "frmSup.frx":FAAB6
         MultiSelect     =   2  'Extended
         TabIndex        =   21
         Top             =   570
         Width           =   2265
      End
      Begin VB.ListBox lhm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2910
         Index           =   1
         ItemData        =   "frmSup.frx":FAAB8
         Left            =   3600
         List            =   "frmSup.frx":FAABA
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   1140
         Width           =   1245
      End
      Begin VB.ListBox lhm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2910
         Index           =   0
         ItemData        =   "frmSup.frx":FAABC
         Left            =   2430
         List            =   "frmSup.frx":FAABE
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   1140
         Width           =   1125
      End
      Begin VB.CommandButton cmdCard 
         Caption         =   "&Activate"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4890
         TabIndex        =   16
         Top             =   1860
         Width           =   1455
      End
      Begin VB.CommandButton cmdCard 
         Caption         =   "&Accept"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4890
         TabIndex        =   15
         Top             =   1500
         Width           =   1455
      End
      Begin VB.CommandButton cmdCard 
         Caption         =   "&Reset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4890
         TabIndex        =   14
         Top             =   1140
         Width           =   1455
      End
      Begin MSComctlLib.ListView lstc 
         Height          =   2910
         Left            =   6390
         TabIndex        =   12
         Top             =   1140
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   5133
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16761024
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5070
         ItemData        =   "frmSup.frx":FAAC0
         Left            =   90
         List            =   "frmSup.frx":FAAC2
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   930
         Width           =   2265
      End
      Begin VB.CommandButton cmdGen 
         Caption         =   "&Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   9
         Top             =   540
         Width           =   1095
      End
      Begin VB.ComboBox cbc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   90
         Width           =   765
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Is a card code? ENTER to find!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Index           =   1
         Left            =   -66510
         TabIndex        =   63
         Top             =   510
         Width           =   1185
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Read Old DB Ver 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   1
         Left            =   -72480
         TabIndex        =   59
         Top             =   4740
         Width           =   1500
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selection Input"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   0
         Left            =   -69180
         TabIndex        =   40
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Send Data to Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   0
         Left            =   -72480
         TabIndex        =   35
         Top             =   4350
         Width           =   1470
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save in Local Disk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   -72480
         TabIndex        =   34
         Top             =   3240
         Width           =   1320
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get From Local Disk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   -72480
         TabIndex        =   33
         Top             =   3600
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get From Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   -72480
         TabIndex        =   32
         Top             =   3990
         Width           =   1200
      End
      Begin VB.Shape Shape3 
         Height          =   165
         Left            =   -72540
         Top             =   5490
         Width           =   3225
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   2385
         Left            =   -72540
         Top             =   3030
         Width           =   3225
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   -70290
         TabIndex        =   25
         Top             =   5700
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   -72510
         TabIndex        =   23
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cards"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   -74910
         TabIndex        =   22
         Top             =   330
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "How many         Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2460
         TabIndex        =   20
         Top             =   930
         Width           =   1545
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Specify Distribution, Ex 20 cards of 150$00 and so on... Select 20 then 150 then click Accept"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   2400
         TabIndex        =   17
         Top             =   660
         Width           =   6720
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   3195
         Index           =   2
         Left            =   2400
         Top             =   900
         Width           =   7245
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   1935
         Index           =   1
         Left            =   2400
         Top             =   4110
         Width           =   7245
      End
      Begin VB.Label lbPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price 0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   8910
         TabIndex        =   13
         Top             =   6060
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":) Cards are Sorted From 0 to 9!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   60
         TabIndex        =   11
         Top             =   6060
         Width           =   2280
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   5145
         Index           =   0
         Left            =   60
         Top             =   900
         Width           =   2325
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "How many Cards to Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   90
         TabIndex        =   7
         Top             =   150
         Width           =   2160
      End
   End
   Begin MSWinsockLib.Winsock tcp 
      Left            =   8220
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SuperVisorXV3"
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
      Index           =   2
      Left            =   9390
      TabIndex        =   62
      Top             =   7050
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright© 2007 Edson Martins"
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
      Index           =   1
      Left            =   8310
      TabIndex        =   61
      Top             =   7230
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   6915
      Index           =   0
      Left            =   10320
      Picture         =   "frmSup.frx":FAAC4
      Top             =   -3720
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   6915
      Index           =   1
      Left            =   10320
      Picture         =   "frmSup.frx":109186
      Top             =   540
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   180
      TabIndex        =   6
      Top             =   90
      Width           =   1065
   End
   Begin VB.Label lbs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Closed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3870
      TabIndex        =   3
      Top             =   120
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lbp 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   7140
      Width           =   7755
   End
End
Attribute VB_Name = "frmSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_SELECTSTRING = &H18C
Private ccard As Long
Private findFlag As Long
Private ccount As Long
Private rcard As Long

Private Sub chkBal_Click()
If Me.chkBal.Value = 0 Then
    Me.opsearch(1).Caption = "Price > ?"
    Me.opsearch(2).Caption = "Price < ?"
    Me.opsearch(3).Caption = "Price = ?"
Else
    Me.opsearch(1).Caption = "Remain > ?"
    Me.opsearch(2).Caption = "Remain < ?"
    Me.opsearch(3).Caption = "Remain = ?"
End If

End Sub

Public Sub cmdCard_Click(Index As Integer)
Dim gen&

Select Case Index
Case 0
ccard = 0
Me.lhm(0).Clear
Me.lhm(1).Clear
Me.lstc.ListItems.Clear
Me.lstc.Tag = ""
If Me.List1.ListCount <> 0 Then
setRedraw Me.lhm(0)
For Index = 1 To Me.List1.ListCount Step 1
    Me.lhm(0).AddItem Format(Index, "@@@")
    DoEvents
Next
setRedraw Me.lhm(0), 1

End If
setRedraw Me.lhm(1)
Me.lhm(1).AddItem Format(Format(75, "#,##0.00"), "@@@@@@@")
For Index = 100 To 5000 Step 50
DoEvents
Me.lhm(1).AddItem Format(Format(CDbl(Index), "#,##0.00"), "@@@@@@@")
Next
setRedraw Me.lhm(1), 1
Case 1
If Me.lhm(0).ListIndex <> -1 And Me.lhm(1).ListIndex <> -1 Then
    With Me.lstc
        .ListItems.Add , , Trim$(Me.lhm(0).Text)
        .ListItems(.ListItems.Count).SubItems(1) = Me.lhm(1).Text
        .ListItems(.ListItems.Count).SubItems(2) = _
        Format(Val(Replace(Me.lhm(1).Text, ".", "")) * Val(Replace(Me.lhm(0).Text, ".", "")), "#,##0.00")
        Me.lstc.Tag = Val("0" & Me.lstc.Tag) + _
        Val(Replace(Me.lhm(1).Text, ".", "")) * Val(Replace(Me.lhm(0).Text, ".", ""))
        ccard = ccard + Val("0" & Trim$(Me.lhm(0).Text))
        Me.lhm(1).RemoveItem Me.lhm(1).ListIndex
        Me.lhm(1).ListIndex = -1
        
        
        gen& = CLng(Me.lhm(0).Text)
        gen& = CLng(Me.lhm(0).ListCount - gen)
        Me.lhm(0).ListIndex = -1
        Me.lhm(0).Clear
        setRedraw Me.lhm(0)
        For Index = 1 To gen
            Me.lhm(0).AddItem Format(Index, "@@@")
        Next
        setRedraw Me.lhm(0), 1
        
    End With
    

End If

Case 2
If MsgBox("Do you really want to execute Activate command?", vbQuestion + vbYesNo) = vbYes Then
associateCards
Me.List1.Clear
Me.SSTab1.Tab = 1
End If

End Select

End Sub

Private Function associateCards()
Dim cc&, tc As Long, offset As Long, vr&, prc
Dim card$
offset = 0
Me.lscard.Clear
Me.lcData.Clear
For cc = 1 To Me.lstc.ListItems.Count
    tc = CLng(Val(Me.lstc.ListItems(cc).Text))
    prc = CLng(Val(Me.lstc.ListItems(cc).SubItems(1)))
    For vr = 1 To tc
        '-----------------------------------
        lscard.AddItem Me.List1.List(offset), offset
        card$ = Format(offset + 1, "000") 'ID GENERATION
        card$ = card$ & Me.List1.List(offset) 'CODE
        card$ = card$ & Format(Now, "ddmmyy")
        card$ = card$ & "00" 'LIFES
        card$ = card$ & "F" 'FLAG NOT USED
        card$ = card$ & Format(prc, "0000")
        card$ = card$ & "0000000000000000"
        '---------------------------------------
        Debug.Print card$, Len(card$)
        Me.lcData.AddItem card$, offset
        Me.List1.Selected(offset) = True
        offset = offset + 1
    Next
   

Next
'001 1111-2222-3333-4444 00 100806 T 0200 0014 000000000000

End Function


Private Sub cmdConUn_Click()
Select Case Left(Me.cmdConUn.Caption, 2)
    Case "&D"
        Me.cmdConUn.Caption = "&Connect"
        Me.tcp.Close
    Case "&C"
        Me.cmdConUn.Caption = "&Disconnect"
        Me.tcp.RemotePort = 4000
        Me.tcp.RemoteHost = Me.txtHost.Text
        Me.tcp.Connect
        
End Select

End Sub

Private Sub cmdGen_Click()
If Me.cbc.Text = "" Then
tell "PLease select how many cards to generate..."
Else

Me.List1.Clear
setRedraw Me.List1
generateCards
setRedraw Me.List1, 1
cmdCard_Click 0
End If

End Sub

Private Sub cmdRebd_Click(Index As Integer)
Dim ccoun As Long
busy
Me.cmdRebd(Index).Enabled = False
Select Case Index
Case 0
If dbOpenned = True Then
ccoun = getCardCount
closeDb
    FileCopy dbName, App.Path & "\BCK_" & Format(Now, "dd_mm_yy_hh_mm_ss\[") & ccoun & "].bck"
    tell "Backup created sucessfull Name BCK_" & Format(Now, "dd_mm_yy_hh_mm_ss\[") & ccoun & "].bck"
OpenDB
End If

Case 1
    If MsgBox("Do you really want to rebuild Local Data Base?", vbQuestion + vbYesNo) = vbYes Then
        closeDb
        creatDB True
        OpenDB
    End If
Case 2
    If MsgBox("Do you really want to rebuild Server Data Base?", vbQuestion + vbYesNo) = vbYes Then
        talk "DEL19812005"
    End If
Case 3
        talk "DTL"
End Select
Me.cmdRebd(Index).Enabled = True
free
End Sub

Private Sub talk(ByVal what$)
If Me.tcp.State = 7 Then
    Me.tcp.SendData what
End If

End Sub

Private Sub cmdSelect_Click()
Me.lbDet.Text = ""
Dim tcard$, mycard As CARDI
Timer1.Enabled = False
Me.Frame3.Visible = False
Me.tmTime.Enabled = False
busy
Me.cmdSelect.Enabled = False
Dim itr&

Pause 0.04
     
If Me.lscard.ListCount <> 0 Then

    Select Case findFlag
        Case -1
                For itr& = 0 To Me.lscard.ListCount - 1
                    DoEvents
                    Me.lscard.Selected(itr) = True
                    Me.lscard.ListIndex = itr
                    tell "Selecting. Please wait... " & Format((itr + 1) / Me.lscard.ListCount, "0%")
                Next
            
           
               tell Me.lscard.SelCount & "/" & Me.lscard.ListCount & " Selected"
           
            
        Case 6
            For itr& = 0 To Me.lscard.ListCount - 1
            DoEvents
            Me.lscard.Selected(itr) = False
            Me.lscard.ListIndex = itr
            tell "Searching. Please wait... " & Format((itr + 1) / Me.lscard.ListCount, "0%")
           
            tcard$ = Me.lcData.List(itr)
            tcard$ = Left(tcard$, 51)
            
            CopyMemory mycard, ByVal tcard, 51
            If mycard.code <> Me.lscard.List(itr) Then
                Me.lscard.Selected(itr) = True
            End If
            
            Next
            
            If Me.lscard.SelCount = 0 Then
                tell "There is no Invalid data. Every thing is OK"
            Else
                tell Me.lscard.SelCount & "/" & Me.lscard.ListCount & " card" & IIf(Me.lscard.SelCount > 1, "s", "") & " damaged"
            End If
        Case 0 'date
            Dim dta$
            dta$ = Me.txtFind.Text
            dta$ = Replace(dta$, "200", "0") 'se 2006 06
            dta$ = Replace(dta$, "\", "")
            dta$ = Replace(dta$, "/", "")
            dta$ = Replace(dta$, "-", "")
            Debug.Print dta$
            
            For itr& = 0 To Me.lscard.ListCount - 1
            'DoEvents
            Me.lscard.Selected(itr) = False
            Me.lscard.ListIndex = itr
            tell "Searching. Please wait... " & Format((itr + 1) / Me.lscard.ListCount, "0%")
           
            tcard$ = Me.lcData.List(itr)
            tcard$ = Left(tcard$, 51)
            
            CopyMemory mycard, ByVal tcard, 51
            If Right(mycard.date, Len(dta$)) = dta$ Then
                Me.lscard.Selected(itr) = True
            End If
            
            Next
            
            If Len(Me.txtFind.Text) = 6 Then Me.txtFind.Text = Format(Me.txtFind.Text, "00-00-00")
            If Me.lscard.SelCount = 0 Then
                tell "No cards created at " & Me.txtFind.Text & " are loaded..."
            Else
                tell Me.lscard.SelCount & "/" & Me.lscard.ListCount & " card" & IIf(Me.lscard.SelCount > 1, "s", "") & " created at " & Me.txtFind.Text
            End If
            
        Case 1, 2, 3 '> 'for all them
        Me.txtFind.Text = Format(Val(Me.txtFind.Text), "#,##0.00")
            For itr& = 0 To Me.lscard.ListCount - 1
            DoEvents
            Me.lscard.Selected(itr) = False
            Me.lscard.ListIndex = itr
            
            tell "Searching. Please wait... " & Format((itr + 1) / Me.lscard.ListCount, "0%")
            
            
            tcard$ = Me.lcData.List(itr)
            tcard$ = Left(tcard$, 51)
            
            CopyMemory mycard, ByVal tcard, 51
            If Me.chkBal.Value = 0 Then
                Select Case findFlag 'individual treatment
                    Case 1
                    If Val(Me.txtFind.Text) < Val(mycard.tbal) Then _
                    Me.lscard.Selected(itr) = True
                    Case 2
                    If Val(Me.txtFind.Text) > Val(mycard.tbal) Then _
                    Me.lscard.Selected(itr) = True
                    Case 3
                    If Val(Me.txtFind.Text) = Val(mycard.tbal) Then _
                    Me.lscard.Selected(itr) = True
                End Select
            Else
            Dim tvc As Double
            tvc = Val(Val(mycard.tbal) - Val(mycard.tusd))
                Select Case findFlag 'individual treatment
                    Case 1
                    If Val(Me.txtFind.Text) < tvc Then _
                    Me.lscard.Selected(itr) = True
                    Case 2
                    If Val(Me.txtFind.Text) > tvc Then _
                    Me.lscard.Selected(itr) = True
                    Case 3
                    If Val(Me.txtFind.Text) = tvc Then _
                    Me.lscard.Selected(itr) = True
                End Select
            End If
            
            Next
            
            If Me.lscard.SelCount = 0 Then
                tell "No cards satisfy your search input criterion."
            Else
                tell Me.lscard.SelCount & "/" & Me.lscard.ListCount & " card" & IIf(Me.lscard.SelCount > 1, "s", "") & " found based on your search input criterion.."
            End If
            
        Case 4 'used
            For itr& = 0 To Me.lscard.ListCount - 1
            DoEvents
            Me.lscard.Selected(itr) = False
            Me.lscard.ListIndex = itr
            
            tell "Searching. Please wait... " & Format((itr + 1) / Me.lscard.ListCount, "0%")
                
                If InStr(Me.lcData.List(itr), "F") = 0 Then
                Me.lscard.Selected(itr) = True
                tell "Selecting. Please wait... " & Format((itr + 1) / Me.lscard.ListCount, "0%")
                End If
                
            Next
            
            If Me.lscard.SelCount = 0 Then
                tell "There is no used cards in the list."
            Else
                tell Me.lscard.SelCount & " used card" & IIf(Me.lscard.SelCount > 1, "s", "") & " found."
            End If
        Case 5 'new
            For itr& = 0 To Me.lscard.ListCount - 1
            DoEvents
            Me.lscard.Selected(itr) = False
            Me.lscard.ListIndex = itr
            
            tell "Searching. Please wait... " & Format((itr + 1) / Me.lscard.ListCount, "0%")
               
                If InStr(Me.lcData.List(itr), "F") <> 0 Then
                Me.lscard.Selected(itr) = True
                End If
                
            Next
            
            If Me.lscard.SelCount = 0 Then
                tell "There is no new cards in the list."
            Else
                tell Me.lscard.SelCount & " new card" & IIf(Me.lscard.SelCount > 1, "s", "") & " found."
            End If
            
    End Select
    

End If

free
Me.cmdSelect.Enabled = True
Me.tmTime.Enabled = True
Timer1.Enabled = True
End Sub

Private Sub cmdTrans_Click(Index As Integer)
busy
Me.lbDet.Text = ""
Dim code$, data$, ok&, bad&, itr&
Me.cmdTrans(Index).Enabled = False

ok& = 0
bad& = 0
Select Case Index
Case 4
Dim fnameinport As String
fnameinport = InputBox("Type the complete Old Data Base Path to Import", App.EXEName & " Migration...", "CardsL.mdb")

If fnameinport <> "" Then
    If Dir(fnameinport) <> "" Then
    Me.lscard.Clear
    Me.lcData.Clear
     setRedraw Me.lscard
        importDataFromOldContentor fnameinport
 setRedraw Me.lscard, 1
Else
Me.lbp.Caption = "File not found..."
    End If
    
End If

Case 3

Me.pgb.Min = 0
Me.pgb.Max = Me.lscard.ListCount

For itr& = 1 To Me.lscard.ListCount
DoEvents
Me.pgb.Value = itr& - 1
Me.lscard.ListIndex = itr - 1
code$ = Me.lscard.List(itr& - 1)
data$ = Me.lcData.List(itr& - 1)
If InStr(1, data$, code$, vbTextCompare) <> 0 Then
ok& = ok& + 1
addUpdateCard Me.lscard.List(itr - 1), Me.lcData.List(itr - 1)
Else
bad& = bad& + 1
End If

Pause 0.001
tell "Saving data " & Format(itr& / Me.lscard.ListCount, "0%")
Next

Pause 2
tell "Total Cards " & Me.lscard.ListCount & ", Saved " & ok & " Error " & bad

Me.lbDet.Text = ""
If bad = 0 Then
Me.lscard.Clear
Me.lcData.Clear
End If

Case 0
If Me.lscard.ListCount <> 0 Then
    If MsgBox("This action will remove the current data of the list" & _
    vbCrLf & "Continue anyway?", vbQuestion + vbYesNo) = vbNo Then
        GoTo fim
    End If
    Me.lscard.Clear
    Me.lcData.Clear
End If
 setRedraw Me.lscard
 getCard Me.lscard, Me.lcData
 setRedraw Me.lscard, 1
 
Case 1
    If Me.lscard.ListCount <> 0 Then
        If MsgBox("This action will clear the List to Continue..." & _
        vbCrLf & "Continue anyway?", vbQuestion + vbYesNo) = vbNo Then
            GoTo fim
        End If
        Me.lscard.Clear
        Me.lcData.Clear
    End If
    talk "REP" 'ler o numero de contas...
    Pause 0.8
    If ccount <> 0 Then
    rcard = 0
        talk "GET"
    Else
    'tell "No Cards found at " & Me.txtHost.Text
    End If
    
    
Case 2
        If MsgBox("Do you really want to send selected cards" & _
        vbCrLf & "to Server: " & Me.txtHost.Text, vbQuestion + vbYesNo) = vbNo Then
            GoTo fim
        End If
        
        tell "Preparing to read Send Cards to " & Me.txtHost.Text
        Pause 0.9
        Me.pgb.Min = 0
        Me.pgb.Max = Me.lscard.ListCount
        For itr& = 0 To Me.lscard.ListCount - 1
        Me.lbp = "Sending Cards " & Format((itr + 1) / Me.lscard.ListCount, "0%")
           If Me.lscard.Selected(itr) = True Then
           gcmd = ""
            
            Me.pgb.Value = itr& + 1
            
            talk "SAV" & Me.lcData.List(itr)
            Do
                DoEvents
            Loop Until gcmd = "NEX"
           End If
           
        Next
        Pause 0.6
        tell "Selected card" & IIf(Me.lscard.SelCount > 1, "s", "") & " was sent to " & Me.txtHost.Text
    
End Select

fim:
Me.cmdTrans(Index).Enabled = True
free
End Sub

Private Sub cmdWhatDo_Click(Index As Integer)
Dim ok&, bad&, itr&, code$, data$
Select Case Index
    Case 0
    If MsgBox("Do you really want to remove selected Card" & IIf(Me.lscard.SelCount > 1, "s?", "?"), vbQuestion + vbYesNo, App.EXEName) = vbYes Then
        busy
        Me.pgb.Min = 0
        Me.pgb.Max = Me.lscard.SelCount
        For itr& = 1 To Me.lscard.ListCount
        DoEvents
       
        Me.lscard.ListIndex = itr - 1
        code$ = Me.lscard.List(itr - 1)
        data$ = Me.lcData.List(itr - 1)
        If InStr(1, data$, code$, vbTextCompare) <> 0 Then
      
        Pause 0.001
        If Me.lscard.Selected(itr - 1) = True Then
         Me.pgb.Value = Me.pgb.Value + 1
          tell "Removing card " & Me.lscard.List(itr - 1) & Format(itr& / Me.lscard.ListCount, ", 0% ") & " concluded!"
            If removecard(Me.lscard.List(itr - 1)) = True Then
                ok& = ok& + 1
            Else
                bad& = bad& + 1
            End If
        
        Else
       'skiped
        End If
        End If
        
        
        Next
        Me.lbDet.Text = ""
        If bad <> 0 Then
        tell "Errors occurred while removiing some cards (" & bad& & ")"
        
        Else
        tell "Selected Card" & IIf(Me.lscard.ListCount > 1, "s", "") & " was removed." & ok& & "/" & Me.lscard.SelCount & " To cancel click Save, Load to Confirm."
        
        End If
        
        free
    End If
    
    Case 1
    Load frmdet
    frmdet.Timer1.Enabled = True
    busy
    Dim mycard As CARDI
    
    For itr = 0 To Me.lscard.ListCount - 1
    DoEvents
    If Me.lscard.Selected(itr) = True Then
    data = Me.lcData.List(itr)
    data$ = Left(data$, 51)
    
    CopyMemory mycard, ByVal data$, 51
        With frmdet.lstdet
        
            .ListItems.Add , , mycard.id
            .ListItems(.ListItems.Count).Checked = True
            .ListItems(.ListItems.Count).SubItems(1) = mycard.code
            .ListItems(.ListItems.Count).SubItems(2) = Format(mycard.date, "00-00-00")
            .ListItems(.ListItems.Count).SubItems(3) = mycard.life
            .ListItems(.ListItems.Count).SubItems(4) = mycard.flag
            .ListItems(.ListItems.Count).SubItems(5) = Format(CLng(Trim$(mycard.tbal)), "#,##0.00")
            .ListItems(.ListItems.Count).SubItems(6) = Format(CLng(Trim$(mycard.tusd)), "#,##0.00")
            .ListItems(.ListItems.Count).SubItems(7) = Format(CLng(Trim$(mycard.tbal)) - CLng(Trim$(mycard.tusd)), "#,##0.00")
            .ListItems(.ListItems.Count).SubItems(8) = trasnBytes(CLng(mycard.bytes))
            
        End With
        Else
        tell "You must select cards to view details first."
    End If
    
    Next
    
    free
    frmdet.Show
            
    Case 2
    
End Select

End Sub

Private Sub Form_Load()
findFlag = -2
Me.lbp.Caption = "Terminal: \\" & Environ("COMPUTERNAME")

Me.txtHost.Text = GetSetting(App.EXEName, "host", "name", Environ("COMPUTERNAME"))

Dim itr

For itr = 1 To 250 Step 1
    Me.cbc.AddItem Format(itr, "000")
    DoEvents
Next

Me.lhm(1).AddItem Format(Format(75, "#,##0.00"), "@@@@@@@")
For itr = 100 To 5000 Step 50
DoEvents
Me.lhm(1).AddItem Format(Format(itr, "#,##0.00"), "@@@@@@@")
Next


With Me.lstc
    .ColumnHeaders.Add , , "Qtd.", 0.22 * .Width
    .ColumnHeaders.Add , , "Price", 0.34 * .Width
    .ColumnHeaders(2).Alignment = lvwColumnRight
    .ColumnHeaders.Add , , "Total", 0.34 * .Width
    .ColumnHeaders(3).Alignment = lvwColumnRight
    .LabelEdit = lvwManual
    .View = lvwReport

End With

creatDB
OpenDB
Do
DoEvents
Me.Visible = True
Loop Until Me.Visible = True
topMost HWND_TOPMOST, frmPass

Do
DoEvents
frmPass.Visible = True
Loop Until frmPass.Visible = True
frmPass.txtPass(0).SetFocus
frmPass.Visible = False
frmPass.Show vbModal, Me


If logged = False Then
Unload Me
End
End If

End Sub

Private Function generateCards()
Screen.MousePointer = 11
Me.Enabled = False
tell "Generating Cards..."
    Dim numb As Double, card$
    Me.cbc.Enabled = False
    Me.cmdGen.Enabled = False
    Me.List1.Clear
    Me.lhm(0).Clear
    Dim itr&
    For itr = 1 To CLng(Me.cbc.Text)
    'Me.lhm(0).AddItem Format(itr, "@@@")
    Randomize 9999999
    numb = CDbl(Format(Now, "ddmmyhssmm"))
    Do
    DoEvents
    numb = CDbl(Format(Now, "ddmmyyssmm")) * itr * Rnd * (9999999)
    card$ = Format(numb, "0000-0000-0000-0000")
    card = Right(card, 19)
    card = Replace(card$, "0000", Format(Rnd * (9999), "0000"))
    Loop Until Len(card) = 19 And InStr(1, card, "0000") = 0
    Debug.Print card
    Dim exist&
    exist = SendMessage(Me.List1.hwnd, LB_SELECTSTRING, -1, ByVal card$)
    If exist <> -1 Then
    tell "Already exist " & card
    Else
    Me.List1.AddItem card$
    End If
    
    tell Format(Me.List1.ListCount / CLng(Me.cbc.Text), "0%")
    Next
    Me.cbc.Enabled = True
    Me.cmdGen.Enabled = True
    tell "Cards Generated and Checked"
Screen.MousePointer = 0
Me.Enabled = True
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If logged = True Then
Cancel = True
doend
End If

End Sub

Private Sub doend()
Me.tcp.Close
Unload frmdet
Unload frmSup
End
End Sub
Private Sub lbp_Change()
Me.tmTime.Enabled = False
Me.tmTime.Enabled = True
End Sub

Private Sub lscard_Click()
If Me.lscard.ListCount = 0 Then Exit Sub
Dim mycard As CARDI
Me.lcData.Selected(Me.lscard.ListIndex) = True
'Me.lcData.ListIndex = -1
Dim tcrd As String
tcrd = Me.lcData.Text
CopyMemory mycard, ByVal tcrd, 51

With mycard
Me.lbDet.Text = "Id" & vbTab & .id & vbCrLf
Me.lbDet.Text = Me.lbDet.Text & "Code" & vbTab & .code & vbCrLf
Me.lbDet.Text = Me.lbDet.Text & "Created" & vbTab & Format(.date, "00-00-00") & vbCrLf
Me.lbDet.Text = Me.lbDet.Text & "Age" & vbTab & .life & " day" & IIf(.life > 1, "s", "") & vbCrLf
Me.lbDet.Text = Me.lbDet.Text & "New" & vbTab & IIf(.flag = "F", "Yes", "No") & vbCrLf
Me.lbDet.Text = Me.lbDet.Text & "Balance" & vbTab & Format(CLng(.tbal), "#,##0.00") & vbCrLf
Me.lbDet.Text = Me.lbDet.Text & "Used" & vbTab & Format(CLng(.tusd), "#,##0.00") & vbCrLf
Me.lbDet.Text = Me.lbDet.Text & "Remain" & vbTab & Format(CLng(.tbal) - CLng(.tusd), "#,##0.00") & vbCrLf
Me.lbDet.Text = Me.lbDet.Text & "Charge" & vbTab & trasnBytes(CLng(.bytes))

End With

End Sub

Private Sub opsearch_Click(Index As Integer)
findFlag = Index
Me.txtFind.Text = ""
Select Case Index
    Case 0
        tell "Use Selection Input Text to Specify the Date  ex. dd/mm/yy or mm/yy"
    Case 1, 2, 3
        tell "Use Selection Input Text to Specify the price reference"
    Case 5, 4
        Me.txtFind.Text = "Used = " & IIf(Index = 4, "True", "False")
    Case 6
        Me.txtFind.Text = "Error"
    Case 7
        Me.txtFind.Text = "All"
End Select

End Sub

Private Sub opsearch_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Select Case Index
    Case 0
    Case 1, 2, 3
        tell "Use Selection Input Text to Specify the price reference"
    Case 5, 4
    
        
End Select
End If

End Sub

Private Sub tcp_DataArrival(ByVal bytesTotal As Long)
Dim dta$
If tcp.State = 7 Then
tcp.GetData dta$

Select Case Left(dta$, 3)
    Case "MSG"
        tell "RMSG: " & Right(dta, Len(dta) - 3)
    Case "CNT"
    ccount = CLng(Right(dta, Len(dta) - 3))
    Debug.Print ccount
    Me.pgb.Min = 0
    
    Me.pgb.Max = IIf(ccount <> 0, ccount, 1)
    Me.pgb.Value = 0
        tell "Total Cards at " & Me.txtHost.Text & ": " & ccount
    Case "TAK"
        dta = Mid(dta, 4, 51)
        Dim md As CARDI
        Debug.Print Len(dta)
        rcard = rcard + 1
        Me.pgb.Value = rcard
        Me.lbp = "Receiving Data " & Format(rcard / ccount, "0%") & " " & rcard & "/" & ccount
        CopyMemory md, ByVal dta, 51
        
        Me.lscard.AddItem md.code
        Me.lcData.AddItem dta$
        Me.lscard.Selected(Me.lscard.ListCount - 1) = True
        Me.lscard.Selected(Me.lscard.ListCount - 1) = False
        
        Pause 0.04
        If rcard > ccount Then
        Pause 0.54
        tell "Operation completed..."
        Exit Sub
        
        End If
        
        talk "NEX"
    Case "NEX"
    gcmd = "NEX"
End Select

End If

End Sub

Private Sub Timer1_Timer()

Me.cmdTrans(3).Enabled = Me.lscard.ListCount <> 0
Me.cmdTrans(1).Enabled = Me.tcp.State = 7
Me.cmdTrans(2).Enabled = Me.lscard.ListCount <> 0 And Me.tcp.State = 7 And Me.lscard.SelCount <> 0
Me.cmdRebd(2).Enabled = Me.tcp.State = 7
Me.cmdRebd(3).Enabled = Me.tcp.State = 7
Me.cmdCard(1).Enabled = Me.lhm(0).ListIndex <> -1 And Me.lhm(1).ListIndex <> -1
Me.lbPrice = "Accepetd: " & ccard & ", Total Price " & Format(Val(Me.lstc.Tag), "#,##0.00")
Me.cmdCard(2).Enabled = Me.lstc.ListItems.Count <> 0 And Me.lhm(0).ListCount = 0
Me.cmdSelect.Enabled = Me.lscard.ListCount <> 0 And findFlag <> -2

Dim tps$
DoEvents
Select Case Me.tcp.State
    Case 0
    tps = "Disconnected"
    If Me.cmdConUn.Caption <> "&Connect" Then Me.cmdConUn.Caption = "&Connect"
   
    Case sckConnecting
    tps = "Connecting"
    Case 7
    tps = "Connected"
    If Me.cmdConUn.Caption <> "&Disconnect" Then Me.cmdConUn.Caption = "&Disconnect"
    Case 8, 9
    tps = "Error"
    Me.tcp.Close

End Select
tps = tps & ", Data Base:" & IIf(dbOpenned = True, "Openned", "Closed")
If Me.lbs <> tps Then Me.lbs = tps

Me.Frame3.Visible = Me.lscard.SelCount <> 0

End Sub

Private Sub Tmtime_Timer()
tmTime.Enabled = False

'tell "Terminal: \\" & Environ("COMPUTERNAME")
If Me.pgb.Max = Me.pgb.Value Then Me.pgb.Value = 0

End Sub

Private Sub txtFind_Change()
 If Me.txtFind.Text = "*" Or UCase(Me.txtFind.Text) = "ALL" Then
 findFlag = -1
 Me.cmdSelect.Enabled = True
 ElseIf Me.txtFind.Text = "" Then
 Me.cmdSelect.Enabled = False
 End If
 
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
Dim tmp$
If KeyCode = 13 Then
    If Me.txtFind.Text <> "" Then
        For KeyCode = 1 To Me.lscard.ListCount - 1
            tmp$ = Me.lscard.List(KeyCode)
                Me.lscard.Selected(KeyCode) = False
            If InStr(tmp$, Me.txtFind.Text) <> 0 Then
                Me.lscard.Selected(KeyCode) = True
            End If
            
        Next
            tell Me.lscard.SelCount & " card" & IIf(Me.lscard.SelCount > 1, "s", "") & " maching your input..."
    End If
End If

End Sub

Private Sub txtHost_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtHost.Text = Replace(Me.txtHost.Text, "\\", "")
    SaveSetting App.EXEName, "host", "name", Trim$(Me.txtHost.Text)
End If

End Sub
Public Function tell(ByRef what$)
'mostra mensagens para o utilizador
    If Me.lbp = what$ Then Exit Function
    
    Me.lbp = what$
    tmTime.Enabled = False
    tmTime.Enabled = True
    
End Function



