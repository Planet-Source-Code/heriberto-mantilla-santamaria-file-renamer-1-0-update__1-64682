VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Acerca de..."
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":2372
   ScaleHeight     =   3375
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Rename.SOfficeButton cmdExit 
      Height          =   405
      Left            =   60
      TabIndex        =   0
      Top             =   2760
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   714
      BackColor       =   16777215
      Caption         =   "    &Cerrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayIcon        =   0   'False
      MouseIcon       =   "frmAbout.frx":AE3E
      MousePointer    =   99
      Picture         =   "frmAbout.frx":B158
      PictureAlign    =   1
      SetBorder       =   -1  'True
      ShadowText      =   -1  'True
      ShowFocus       =   -1  'True
      SystemColor     =   0   'False
      TipBackColor    =   14811135
      TipForeColor    =   0
   End
End
Attribute VB_Name = "frmAbout"
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
'* Description:   About form.                          *'
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

Private Sub cmdExit_Click()
 Call Unload(frmAbout)
 Set frmAbout = Nothing
End Sub
