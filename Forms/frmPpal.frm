VERSION 5.00
Begin VB.Form frmPpal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   8985
      TabIndex        =   3
      Top             =   0
      Width           =   8985
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Choose files"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7215
         Left            =   4950
         TabIndex        =   32
         Top             =   60
         Width           =   3780
         Begin VB.DriveListBox DrvDD 
            Height          =   315
            Left            =   600
            TabIndex        =   36
            Top             =   375
            Width           =   3045
         End
         Begin VB.DirListBox DirDD 
            Appearance      =   0  'Flat
            Height          =   2790
            Left            =   90
            TabIndex        =   35
            Top             =   1110
            Width           =   3570
         End
         Begin VB.FileListBox FileDD 
            Appearance      =   0  'Flat
            Height          =   2175
            Left            =   90
            MultiSelect     =   2  'Extended
            System          =   -1  'True
            TabIndex        =   34
            Top             =   4290
            Width           =   3570
         End
         Begin VB.TextBox txtFilePattern 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   2970
            TabIndex        =   33
            Text            =   "*.*"
            Top             =   780
            Width           =   690
         End
         Begin Rename.SOfficeButton SOffBtnSel 
            Height          =   465
            Index           =   1
            Left            =   2670
            TabIndex        =   37
            Top             =   6615
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   820
            BackColor       =   16777215
            Caption         =   "Select All"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayIcon        =   0   'False
            MouseIcon       =   "frmPpal.frx":2372
            MousePointer    =   99
            ShadowText      =   -1  'True
            SystemColor     =   0   'False
            TipBackColor    =   14811135
            TipForeColor    =   0
         End
         Begin Rename.SOfficeButton SOffBtnSel 
            Height          =   465
            Index           =   0
            Left            =   1440
            TabIndex        =   38
            Top             =   6615
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   820
            BackColor       =   16777215
            Caption         =   "Select None"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayIcon        =   0   'False
            MouseIcon       =   "frmPpal.frx":268C
            MousePointer    =   99
            ShadowText      =   -1  'True
            SystemColor     =   0   'False
            TipBackColor    =   14811135
            TipForeColor    =   0
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Directories"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   90
            TabIndex        =   42
            Top             =   870
            Width           =   765
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Drive:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   41
            Top             =   420
            Width           =   435
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Files in this directory"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   40
            Top             =   4035
            Width           =   1470
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pattern:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   2340
            TabIndex        =   39
            Top             =   840
            Width           =   600
         End
      End
      Begin VB.ListBox lstPreview 
         Appearance      =   0  'Flat
         Height          =   2955
         Left            =   90
         TabIndex        =   29
         Top             =   4320
         Width           =   4725
      End
      Begin VB.Frame framOper 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "File Operations"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3585
         Left            =   90
         TabIndex        =   4
         Top             =   60
         Width           =   4725
         Begin VB.TextBox txtFileOper 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   0
            Left            =   1785
            TabIndex        =   19
            Top             =   390
            Width           =   2820
         End
         Begin VB.TextBox txtInitialPos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1020
            TabIndex        =   18
            Text            =   "1"
            Top             =   780
            Width           =   330
         End
         Begin VB.CheckBox chkUpper 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ProperCase Words"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2955
            TabIndex        =   17
            Top             =   2460
            Width           =   1695
         End
         Begin VB.TextBox txtFileOper 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   1
            Left            =   1785
            TabIndex        =   16
            Top             =   1170
            Width           =   1635
         End
         Begin VB.OptionButton OptOpc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Replace All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   15
            Top             =   2925
            Width           =   1095
         End
         Begin VB.OptionButton OptOpc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Set &Before"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1425
            TabIndex        =   14
            Top             =   2925
            Width           =   1095
         End
         Begin VB.OptionButton OptOpc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Set &After"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   13
            Top             =   3210
            Width           =   1095
         End
         Begin VB.TextBox txtCounter 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1305
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "1"
            Top             =   2070
            Width           =   285
         End
         Begin VB.HScrollBar HSInc 
            Height          =   270
            Left            =   1665
            Max             =   100
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2085
            Value           =   1
            Width           =   480
         End
         Begin VB.TextBox txtFileOper 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   2
            Left            =   1170
            TabIndex        =   10
            ToolTipText     =   "You can use | for separate one or more characters."
            Top             =   1560
            Width           =   1500
         End
         Begin VB.TextBox txtFileOper 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   3
            Left            =   3105
            TabIndex        =   9
            Top             =   1560
            Width           =   1500
         End
         Begin VB.CheckBox chkLower 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lower case"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3525
            TabIndex        =   8
            Top             =   1230
            Width           =   1185
         End
         Begin VB.TextBox txtFinalPos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   2265
            TabIndex        =   7
            Text            =   "-"
            Top             =   780
            Width           =   330
         End
         Begin VB.CheckBox chkRDblSpace 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Only spaces at time"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2925
            TabIndex        =   6
            Top             =   825
            Width           =   1725
         End
         Begin VB.OptionButton OptOpc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&None"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1425
            TabIndex        =   5
            Top             =   3210
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Changed Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   28
            Top             =   420
            Width           =   1155
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Initial Pos:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   27
            Top             =   810
            Width           =   750
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Changed Ext:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   26
            Top             =   1200
            Width           =   990
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Options                                           "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   25
            Top             =   2595
            Width           =   2490
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Begin Counter:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   165
            TabIndex        =   24
            Top             =   2115
            Width           =   1080
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Incremental counter by ?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   2280
            TabIndex        =   23
            Top             =   2115
            Width           =   2280
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Replace this:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   165
            TabIndex        =   22
            Top             =   1590
            Width           =   930
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   2835
            TabIndex        =   21
            Top             =   1635
            Width           =   90
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Final Pos:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1470
            TabIndex        =   20
            Top             =   810
            Width           =   690
         End
         Begin VB.Image imgLogo 
            Height          =   480
            Left            =   4200
            MouseIcon       =   "frmPpal.frx":29A6
            MousePointer    =   99  'Custom
            Picture         =   "frmPpal.frx":2CB0
            Top             =   2940
            Width           =   480
         End
      End
      Begin Rename.SOfficeButton SOffBtnRun 
         Height          =   465
         Left            =   3450
         TabIndex        =   30
         Top             =   3780
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   820
         BackColor       =   16777215
         Caption         =   "&Run Preview"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayIcon        =   0   'False
         MouseIcon       =   "frmPpal.frx":5022
         MousePointer    =   99
         ShadowText      =   -1  'True
         SystemColor     =   0   'False
         TipBackColor    =   14811135
         TipForeColor    =   0
      End
      Begin Rename.SOfficeButton SOffBtnRename 
         Height          =   465
         Left            =   1980
         TabIndex        =   31
         Top             =   3780
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   820
         BackColor       =   16777215
         Caption         =   "R&ename"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayIcon        =   0   'False
         MouseIcon       =   "frmPpal.frx":533C
         MousePointer    =   99
         ShadowText      =   -1  'True
         SystemColor     =   0   'False
         TipBackColor    =   14811135
         TipForeColor    =   0
      End
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   300
      Picture         =   "frmPpal.frx":5656
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   2
      Top             =   900
      Width           =   900
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   1
      Left            =   300
      Picture         =   "frmPpal.frx":6157
      ScaleHeight     =   120
      ScaleWidth      =   840
      TabIndex        =   1
      Top             =   720
      Width           =   840
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   270
      Picture         =   "frmPpal.frx":638F
      ScaleHeight     =   390
      ScaleWidth      =   3750
      TabIndex        =   0
      Top             =   270
      Width           =   3750
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************'
'*        All rights Reserved © HACKPRO TM 2006        *'
'*******************************************************'
'*                   Version 1.0.0                     *'
'*******************************************************'
'* Control:       File Renamer                         *'
'*******************************************************'
'* Author:        Heriberto Mantilla Santamaría        *'
'*******************************************************'
'* Description:   Rename any file with special options.*'
'*-----------------------------------------------------*'
'* NOTE                                                *'
'*                Based on the submission Swertvaegher *'
'*                Stephan.                             *'
'*                                                     *'
'*                See the post: [CodeId = 3163].       *'
'*-----------------------------------------------------*'
'*                Now support skin, thx to John Under- *'
'*                hill (Steppenwolfe).                 *'
'*                                                     *'
'*                See the post: [CodeId = 64357].      *'
'*******************************************************'
'* Started on:    Sunday, 05-mar-2006.                 *'
'*******************************************************'
'* Release date:  Friday, 13-mar-2006.                 *'
'*******************************************************'
'* Note:     Comments, suggestions, doubts or bug      *'
'*           reports are wellcome to these e-mail      *'
'*           addresses:                                *'
'*                                                     *'
'*                  heri_05-hms@mixmail.com or         *'
'*                  hcammus@hotmail.com                *'
'*                                                     *'
'*      Please rate my work on this application.       *'
'*             Of Colombia for the world.              *'
'*******************************************************'
'*        All rights Reserved © HACKPRO TM 2006        *'
'*******************************************************'

Option Explicit '<-- Hi Matt, you're always in my App's.
 
Private Type InfoChar
    Position  As Long
    Character As String
End Type
 
Private m_Neo           As cNeoClass

Private Temp      As String, T As Long, NewFile As String, MovText() As InfoChar
Private Temp2     As String, X As Long, OldFile As String, ContainApp As Boolean
Private NumberErr As Boolean
 
Private Sub DirDD_Change()
 FileDD.Path = DirDD.Path
 Call Selections
End Sub

Private Sub DrvDD_Change()
On Error GoTo DriveOut
 Temp = DrvDD.Drive
 DirDD.Path = Left$(DrvDD.Drive, 2) + "\"
 Exit Sub
DriveOut:
 Call MsgBox("Sorry, but the selected device isn't ready!", vbOKOnly & vbCritical, "File Renamer")
End Sub

Private Sub FileDrvDD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Selections
End Sub

Private Sub Form_Load()
 '* Thx to John Underhill (Steppenwolfe).
 Set m_Neo = New cNeoClass
 Call Set_Skin
 HSInc.Value = 1
 DrvDD.Drive = "C:\"
 DirDD.Path = "C:\"
 FileDD.Path = "C:\"
 Call lstPreview.Clear
 Call Selections
End Sub

Private Sub Selections()
 T = 0
 For X = 0 To FileDD.ListCount - 1
  If (FileDD.Selected(X) = True) Then T = T + 1
 Next
End Sub

Private Sub GetNames()
 Dim yMatrix As Variant, Y As Long, Carac  As String
 Dim xMatrix As Variant, i As Long, lPos   As Long
 Dim Tel     As Integer, N As Long, lCarac As String
 
 Tel = 0
 For Y = Len(FileDD.List(X)) To 1 Step -1
  Tel = Tel + 1
  If (Mid$(FileDD.List(X), Y, 1) = ".") Then
   Temp2 = Right$(FileDD.List(X), Tel) '* Extension.
   '* Put Extension lower.
   If (chkLower.Value = 1) Then Temp2 = LCase$(Temp2)
   Temp = Left$(FileDD.List(X), Len(FileDD.List(X)) - Tel) '* Real Name.
   '* Cut the text since the Initial Pos and Final Pos.
   If (Trim$(txtInitialPos.Text) <> "") And (Trim$(txtFinalPos.Text) <> "") And (IsNumeric(txtFinalPos.Text) = True) Then
    Temp = Mid$(Temp, txtInitialPos.Text, txtFinalPos.Text)
   ElseIf (Trim$(txtInitialPos.Text) <> "") And ((Trim$(txtFinalPos.Text) = "") Or (Trim$(txtFinalPos.Text) = "-")) Then
    Temp = Mid$(Temp, txtInitialPos.Text, Len(Temp))
   ElseIf (Trim$(txtInitialPos.Text) = "") And (Trim$(txtFinalPos.Text) <> "") And (IsNumeric(txtFinalPos.Text) = True) Then
    Temp = Mid$(Temp, 1, txtFinalPos.Text)
   End If
   '* Replace character for any text.
   xMatrix = Split(txtFileOper(2).Text, "|")
   yMatrix = Split(txtFileOper(3).Text, "|")
   If (UBound(xMatrix) > 0) Then
  On Error Resume Next
    For i = 0 To UBound(xMatrix)
     Temp = Replace$(Temp, xMatrix(i), yMatrix(i))
    Next
   ElseIf (Trim$(txtFileOper(2).Text) <> "") Then
    Temp = Replace$(Temp, txtFileOper(2).Text, txtFileOper(3).Text)
   End If
   If (ContainApp = True) Then '* Moved specific characters.
    For i = 0 To UBound(MovText)
     Temp = Replace$(Temp, MovText(i).Character, "")
     N = MovText(i).Position + Len(MovText(i).Character)
     Carac = Mid$(Temp, MovText(i).Position + 1, N)
     Temp = Left$(Temp, MovText(i).Position) & MovText(i).Character & Carac
    Next
   End If
   '* Remove Dbl spaces.
   lCarac = ""
   If (chkRDblSpace.Value = 1) Then
    For i = 1 To Len(Temp)
     Carac = Mid$(Temp, i, 1)
     If (lPos = 0) And (Carac = " ") Then
      lPos = lPos + 1
      lCarac = lCarac & Carac
     ElseIf (Carac <> " ") Then
      lPos = 0
      lCarac = lCarac & Carac
     End If
    Next
    '* Trim first and last space.
    Temp = Trim$(lCarac)
   End If
   '* Changed extension.
   If (Trim$(txtFileOper(1).Text) <> "") Then Temp2 = txtFileOper(1).Text
   If (chkUpper.Value = 1) Then Temp = UpperCapitalText(Temp)
   Exit For
  End If
 Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set m_Neo = Nothing
End Sub

Private Sub HSInc_Change()
 lblTitle(7).Caption = "Incremental counter by " & HSInc.Value
End Sub

Private Sub imgLogo_Click()
 Call frmAbout.Show(1)
End Sub

Private Sub SOffBtnRename_Click()
 '* Rename the select files.
On Error GoTo myErr:
 If (ValidatePorc = True) And (NumberErr = False) Then
  If (ContainApp = True) And (OptOpc(3).Value = False) Then
   Call MsgBox("Error in the file name.", vbCritical + vbOKOnly, "File Renamer")
  Else
   Call Execution(1)
   Call FileDD.Refresh
   Call SOffBtnSel_Click(0)
  End If
 Else
  Call MsgBox("Error in the file name.", vbCritical + vbOKOnly, "File Renamer")
 End If
 Exit Sub
myErr:
 Call FileDD.Refresh
 Call SOffBtnSel_Click(0)
 Call MsgBox("Error the file yet exist.", vbCritical + vbOKOnly, "File Renamer")
End Sub

Private Sub SOffBtnRun_Click()
 '* Run demo.
 If (ValidatePorc = True) And (NumberErr = False) Then
  If (ContainApp = True) And (OptOpc(3).Value = False) Then
   Call MsgBox("Error in the file name.", vbCritical + vbOKOnly, "File Renamer")
  Else
   Call Execution
  End If
 Else
  Call MsgBox("Error in the file name.", vbCritical + vbOKOnly, "File Renamer")
 End If
End Sub

Private Sub SOffBtnSel_Click(Index As Integer)
 Select Case Index
  Case 0 '* Select None.
   For X = 0 To FileDD.ListCount - 1
    FileDD.Selected(X) = False
   Next
   Call Selections
  Case 1 '* Select All.
   For X = 0 To FileDD.ListCount - 1
    FileDD.Selected(X) = True
   Next
   Call Selections
 End Select
End Sub

Private Sub txtCounter_KeyPress(KeyAscii As Integer)
 Dim Pos As Integer, CharT As String
 
 Pos = txtCounter.SelStart - 1
 If (Pos <= 0) Then Pos = 1
 CharT = Mid$(txtCounter.Text, Pos, 1)
 If (KeyAscii = 8) Then
  Exit Sub
 ElseIf (KeyAscii = 45) Then
  If (CharT = "-") Or (IsNumeric(CharT) = True) Then
   KeyAscii = 0
   Beep
  End If
  Exit Sub
 End If
 If (CharT = "-") Or (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0: Beep
End Sub

Private Sub txtFilePattern_Change()
 FileDD.Pattern = txtFilePattern.Text
 Call Selections
End Sub

Private Sub txtFinalPos_KeyPress(KeyAscii As Integer)
 Dim Pos As Integer, CharT As String
 
 Pos = txtCounter.SelStart - 1
 If (Pos <= 0) Then Pos = 1
 CharT = Mid$(txtFinalPos.Text, Pos, 1)
 If (KeyAscii = 8) Then
  Exit Sub
 ElseIf (KeyAscii = 45) Then
  If (CharT = "-") Or (IsNumeric(CharT) = True) Then
   KeyAscii = 0
   Beep
  End If
  Exit Sub
 End If
 If (CharT = "-") Or (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0: Beep
End Sub

Private Sub txtInitialPos_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 8) Then Exit Sub
 If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0: Beep
End Sub

Private Sub Execution(Optional ByVal What As Integer = 0)
 Dim Counter As Long, Increment As Long
 
 Call Selections
 If (T = 0) Then
  Call MsgBox("First select 1 or more files!", vbOKOnly + vbExclamation, "File Renamer")
  Exit Sub
 End If
 Call lstPreview.Clear
 Counter = Val(txtCounter.Text)
 Increment = Val(Mid$(lblTitle(7).Caption, Len("Incremental counter by "), Len(lblTitle(7).Caption)))
 For X = 0 To FileDD.ListCount - 1
  If (FileDD.Selected(X) = True) Then
   If (What = 1) Then OldFile = FileDD.Path & "\" & FileDD.List(X)
   Call GetNames
   If (OptOpc(0).Value = True) Then '* Replace
    If (What = 0) Then
     If (txtCounter.Text <> "-") Then
      Call lstPreview.AddItem(txtFileOper(0).Text & Format$(Str(Counter), "00") & Temp2)
     Else
      Call lstPreview.AddItem(txtFileOper(0).Text & Temp2)
     End If
    Else
     If (txtCounter.Text <> "-") Then
      NewFile = FileDD.Path & "\" & txtFileOper(0).Text & Format$(Str(Counter), "00") & Temp2
     Else
      NewFile = FileDD.Path & "\" & txtFileOper(0).Text & Temp2
     End If
     Name OldFile As NewFile
    End If
    Counter = Counter + Increment
   ElseIf (OptOpc(1).Value = True) Then '* Set Before.
    If (What = 0) Then
     If (txtCounter.Text <> "-") Then
      Call lstPreview.AddItem(txtFileOper(0).Text & Temp & Format$(Str(Counter), "00") & Temp2)
     Else
      Call lstPreview.AddItem(txtFileOper(0).Text & Temp & Temp2)
     End If
    Else
     If (txtCounter.Text <> "-") Then
      NewFile = FileDD.Path & "\" & txtFileOper(0).Text & Temp & Format$(Str(Counter), "00") & Temp2
     Else
      NewFile = FileDD.Path & "\" & txtFileOper(0).Text & Temp & Temp2
     End If
     Name OldFile As NewFile
    End If
    Counter = Counter + Increment
   ElseIf (OptOpc(2).Value = True) Then '* Set After.
    If (What = 0) Then
     If (txtCounter.Text <> "-") Then
      Call lstPreview.AddItem(Temp & txtFileOper(0).Text & Format$(Str(Counter), "00") & Temp2)
     Else
      Call lstPreview.AddItem(Temp & txtFileOper(0).Text & Temp2)
     End If
    Else
     If (txtCounter.Text <> "-") Then
      NewFile = FileDD.Path & "\" & Temp & txtFileOper(0).Text & Temp2
     Else
      NewFile = FileDD.Path & "\" & Temp & txtFileOper(0).Text & Format$(Str(Counter), "00") & Temp2
     End If
     Name OldFile As NewFile
    End If
    Counter = Counter + Increment
   ElseIf (What = 0) Then
    Call lstPreview.AddItem(Temp & Trim$(Temp2))
   Else
    NewFile = FileDD.Path & "\" & Temp & Trim$(Temp2)
    Name OldFile As NewFile
   End If
  End If
 Next
End Sub

Private Function UpperCapitalText(ByVal vText As String) As String
 '* Set the First Text in Upper.
 UpperCapitalText = StrConv(vText, vbProperCase)
End Function

Private Function ValidatePorc() As Boolean
 Dim xMatrix As Variant, i   As Long, iCount  As Integer, jPos As Long
 Dim iText   As String, jText As String, ValueT As String
 
 '* Count the / and set validate replace.
 ValidatePorc = True
 ContainApp = False
 xMatrix = Split(txtFileOper(0).Text, "/")
 ReDim Preserve MovText(0)
 ValueT = ""
 jPos = 0
 NumberErr = False
 If ((UBound(xMatrix) Mod 2) = 0) Then
  For i = 1 To Len(txtFileOper(0).Text)
   iText = Mid$(txtFileOper(0).Text, i, 1)
   If (iText = "/") Then
    iCount = iCount + 1
    If (IsNumeric(ValueT) = False) And (ValueT <> "") And (jPos > 0) Then
     MovText(jPos - 1).Character = ValueT
     ValueT = ""
    End If
    ContainApp = True
    If (ValueT <> "") And (IsNumeric(ValueT) = False) And (iCount = 2) Then
     ValidatePorc = False
     ContainApp = False
     NumberErr = True
     Exit For
    ElseIf (iCount = 2) Then
     iCount = 0
     ReDim Preserve MovText(jPos)
     MovText(jPos).Position = Val(ValueT)
     ValueT = ""
     jPos = jPos + 1
    End If
   Else
    ValueT = ValueT & Mid$(txtFileOper(0).Text, i, 1)
   End If
  Next
  If (jPos > 0) Then MovText(jPos - 1).Character = ValueT
 End If
End Function

Private Sub Set_Skin()
 '/* load skin
On Error Resume Next
 With m_Neo
  Set .p_ICaption = picImage(0).Picture
  Set .p_IBorders = picImage(1).Picture
  Set .p_ICBoxMin = picImage(2).Picture
  Set .p_ICBoxMax = picImage(2).Picture
  Set .p_ICBoxRst = picImage(2).Picture
  Set .p_ICBoxCls = picImage(2).Picture
  '/* skin command buttons
  Set .p_ImlRef = Nothing
  .p_SkinCommand = False
  .p_CmdFntClr = &H0
  .p_SkinForm = True
  .p_BorderHasInactive = False
  .p_ControlHasInactive = True
  .p_ButtonHeight = 20
  .p_ButtonWidth = 20
  .p_ControlButtonPosition = True
  .p_ButtonOffsetX = -4
  .p_ButtonOffsetY = 4
  .p_LeftEnd = 224
  .p_ActiveRight = 225
  .p_RightEnd = 250
  .p_Offset = 0
  .p_LeftBorderWidth = 10
  .p_RightBorderWidth = 10
  .p_BottomBorderHeight = 2
  .p_TopBorderHeight = 2
  Call .Attach(Me)
 End With
End Sub
