VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PassWord Screen"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4455
      Left            =   -240
      TabIndex        =   2
      Top             =   -360
      Width           =   4815
      Begin VB.PictureBox Picture9 
         Height          =   975
         Left            =   2400
         ScaleHeight     =   915
         ScaleWidth      =   1035
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
         Begin VB.Label Label900 
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            Height          =   1815
            Left            =   -240
            TabIndex        =   29
            Top             =   -240
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "                                 9"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   -360
            TabIndex        =   20
            Top             =   -120
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture8 
         Height          =   975
         Left            =   1320
         ScaleHeight     =   915
         ScaleWidth      =   1035
         TabIndex        =   10
         Top             =   2280
         Width           =   1095
         Begin VB.Label Label800 
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            Height          =   1815
            Left            =   -120
            TabIndex        =   28
            Top             =   -240
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "                                  8"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   -360
            TabIndex        =   19
            Top             =   -120
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture7 
         Height          =   975
         Left            =   240
         ScaleHeight     =   915
         ScaleWidth      =   1035
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
         Begin VB.Label Label700 
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            Height          =   1815
            Left            =   -360
            TabIndex        =   27
            Top             =   -360
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "                                  7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   -360
            TabIndex        =   18
            Top             =   -120
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture6 
         Height          =   975
         Left            =   2400
         ScaleHeight     =   915
         ScaleWidth      =   1035
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
         Begin VB.Label Label600 
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            Height          =   1815
            Left            =   -480
            TabIndex        =   26
            Top             =   -480
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "                                 6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   -360
            TabIndex        =   17
            Top             =   -120
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture5 
         Height          =   975
         Left            =   1320
         ScaleHeight     =   915
         ScaleWidth      =   1035
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
         Begin VB.Label Label500 
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            Height          =   1815
            Left            =   -240
            TabIndex        =   25
            Top             =   -240
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "                                  5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   -360
            TabIndex        =   16
            Top             =   -120
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   975
         Left            =   240
         ScaleHeight     =   915
         ScaleWidth      =   1035
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
         Begin VB.Label Label400 
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            Height          =   1815
            Left            =   -360
            TabIndex        =   24
            Top             =   -360
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "                                  4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   -360
            TabIndex        =   15
            Top             =   -120
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   975
         Left            =   2400
         ScaleHeight     =   915
         ScaleWidth      =   1035
         TabIndex        =   5
         Top             =   360
         Width           =   1095
         Begin VB.Label Label3000 
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            Height          =   1815
            Left            =   -360
            TabIndex        =   30
            Top             =   -240
            Width           =   1695
         End
         Begin VB.Label Label300 
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            Height          =   1815
            Left            =   -240
            TabIndex        =   23
            Top             =   -240
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "                                  3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   -360
            TabIndex        =   14
            Top             =   -120
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   975
         Left            =   1320
         ScaleHeight     =   915
         ScaleWidth      =   1035
         TabIndex        =   4
         Top             =   360
         Width           =   1095
         Begin VB.Label Label200 
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            Height          =   1815
            Left            =   -600
            TabIndex        =   22
            Top             =   -840
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "                                  2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   -360
            TabIndex        =   13
            Top             =   -120
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   975
         Left            =   240
         ScaleHeight     =   915
         ScaleWidth      =   1035
         TabIndex        =   3
         Top             =   360
         Width           =   1095
         Begin VB.Label Label100 
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            Height          =   1455
            Left            =   -360
            TabIndex        =   21
            Top             =   -240
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "                           1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   -240
            TabIndex        =   12
            Top             =   -120
            Width           =   1335
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Almost There"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label100.Visible = True
Label200.Visible = False
Label200.Visible = False
Label300.Visible = False
Label400.Visible = False
Label500.Visible = False
Label600.Visible = False
Label700.Visible = False
Label800.Visible = False
Label900.Visible = False
Label3000.Visible = False
End Sub

Private Sub Label100_Click()
Label100.Visible = True
Label200.Visible = False
Label200.Visible = False
Label300.Visible = True
Label400.Visible = False
Label500.Visible = False
Label600.Visible = False
Label700.Visible = False
Label800.Visible = False
Label900.Visible = False
Label3000.Visible = False
End Sub

Private Sub Label300_Click()
Label100.Visible = True
Label200.Visible = False
Label200.Visible = False
Label300.Visible = True
Label400.Visible = False
Label500.Visible = False
Label600.Visible = False
Label700.Visible = False
Label800.Visible = False
Label900.Visible = False
Label3000.Visible = True
End Sub

Private Sub Label3000_Click()
Label100.Visible = True
Label200.Visible = False
Label200.Visible = False
Label300.Visible = True
Label400.Visible = False
Label500.Visible = False
Label600.Visible = False
Label700.Visible = True
Label800.Visible = False
Label900.Visible = False
Label3000.Visible = True

End Sub

Private Sub Label700_Click()
Form2.Hide
MsgBox "Access Granted"
Form1.Show
End Sub
