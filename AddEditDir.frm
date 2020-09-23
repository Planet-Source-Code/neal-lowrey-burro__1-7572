VERSION 5.00
Begin VB.Form AddEditDir 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Directory"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4605
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton AddEditCnx 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton AddEditDone 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton BrowseDir 
      Caption         =   "Browse"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox DirPath 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Path"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "AddEditDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddEditCnx_Click()
  UserOpts.Tag = ""
  Unload Me
End Sub

Private Sub AddEditDone_Click()
  UserOpts.Tag = DirPath.Text
  Unload Me
End Sub

Private Sub BrowseDir_Click()
  AddEditDir.Tag = DirPath.Text
  FindFolder.Show 1
  DirPath.Text = AddEditDir.Tag
End Sub

