VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFTP 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTP SERVER"
   ClientHeight    =   4575
   ClientLeft      =   1455
   ClientTop       =   3105
   ClientWidth     =   8355
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmFtp.frx":0000
   LinkTopic       =   "FtpServ"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4575
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox UsrCnt 
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Text            =   "0"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton EndCmd 
      Caption         =   "Close Connection"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Frame StatFrame 
      Caption         =   "Status Window"
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      Begin VB.ListBox LogWnd 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   3165
         ItemData        =   "frmFtp.frx":030A
         Left            =   120
         List            =   "frmFtp.frx":030C
         TabIndex        =   2
         Top             =   300
         Width           =   7815
      End
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4320
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10654
            MinWidth        =   10654
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2187
            MinWidth        =   2187
            TextSave        =   "4/25/00"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "1:50 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "# of Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.Menu mSetup 
      Caption         =   "Setup"
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code for the form frmMTMain.
Public MainApp As MainApp

Private Sub Form_Unload(Cancel As Integer)
  MainApp.Closing
  Set MainApp = Nothing
End Sub

Private Sub EndCmd_Click()
  Dim i As Integer
  For i = 1 To MAX_N_USERS    'close all connection
    If users(i).control_slot <> INVALID_SOCKET Then
      retf = closesocket(users(i).control_slot) 'close control slot
      Set users(i).Jenny = Nothing
    End If
    If users(i).data_slot <> INVALID_SOCKET Then
      retf = closesocket(users(i).data_slot) 'close data slot
    End If
  Next
  retf = closesocket(ServerSlot)
  If SaveProfile(App.Path & "\Burro.ini", True) Then
  End If
  Unload Me
End Sub


Private Sub mSetup_Click()
  UserOpts.Show 1
End Sub

