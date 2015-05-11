VERSION 5.00
Begin VB.Form frmp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2490
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConUn 
      Caption         =   "&Cancel"
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
      Left            =   2970
      TabIndex        =   3
      Top             =   2040
      Width           =   1275
   End
   Begin VB.CommandButton cmdConUn 
      Caption         =   "&Continue"
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
      Left            =   2970
      TabIndex        =   2
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "pType A4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Index           =   1
      Left            =   30
      TabIndex        =   1
      Top             =   1380
      Width           =   2895
      Begin VB.OptionButton Option1 
         Caption         =   "List [code price]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   420
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   735
         Index           =   0
         Left            =   1890
         Picture         =   "frmPrint.frx":000C
         Stretch         =   -1  'True
         Top             =   210
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PType Ticket"
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
      Height          =   1065
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   270
      Width           =   2895
      Begin VB.OptionButton Option2 
         Caption         =   "List [code price]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   645
         Index           =   1
         Left            =   2280
         Picture         =   "frmPrint.frx":08FA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConUn_Click(Index As Integer)
Select Case Index
    Case 0
    Me.cmdConUn(0).Enabled = False
        If Me.Option1.Value = True Then
        printmeA4
        End If
        
    Me.cmdConUn(0).Enabled = True
    Case 1
        Unload Me
End Select

End Sub

Private Sub printmeA4()
Dim Index

Dim lins As Long, page As Long
page = 1
printHead page

For Index = 1 To frmdet.lstdet.ListItems.Count
        DoEvents
        With frmdet.lstdet
          If .ListItems(Index).Checked = True Then
            
              Printer.Print .ListItems(Index).SubItems(1) & "   " & Format(Val(.ListItems(Index).SubItems(5)) * 100, "0$00") & " | ";
              If Index Mod 3 = 0 Then
              
              lins = lins + 1
                    If lins > 29 Then
                      lins = 0
                      Printer.NewPage
                      
                      page = page + 1
                      printHead page
                    Else
                    Printer.Print vbCrLf
                    Printer.Print "     ";
                    End If
              End If
            
         End If
        End With
Next
Printer.EndDoc

End Sub

Private Sub printHead(ByVal page&)
Dim tot As Double
tot = (frmdet.lstdet.ListItems.Count / (28 * 3))
If CLng(tot) < tot Then
tot = CLng(tot) + 1
Else
tot = CLng(tot)
End If

If tot = 0 Then tot = 1
Printer.CurrentX = 4
Printer.CurrentY = 0
Printer.FontName = "arial"
Printer.FontSize = 8
Printer.FontBold = True
Printer.Print "List of Cards Printed at " & Format(Now, "dd-mm-yy hh:mm:ss") & vbTab & vbTab & "Edson Martins Support 9978468, microbodix@hotmail.com" & vbTab & vbTab & vbTab & "Page " & page & " of " & tot

Printer.Print ""
Printer.FontSize = 12
Printer.FontBold = False
Printer.Print "     ";
End Sub

Private Sub Form_Load()
Me.Option1.Value = True
End Sub

