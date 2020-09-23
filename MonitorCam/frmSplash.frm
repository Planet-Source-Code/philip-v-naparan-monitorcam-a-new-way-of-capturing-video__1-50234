VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3480
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3360
      Top             =   720
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H002ED7AF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   -240
      ScaleHeight     =   585
      ScaleWidth      =   6345
      TabIndex        =   4
      Top             =   2880
      Width           =   6375
      Begin VB.Image Image3 
         Height          =   240
         Left            =   2925
         Picture         =   "frmSplash.frx":000C
         Top             =   150
         Width           =   240
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " NaparanSoft"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   150
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyrights 2003 by"
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
         Left            =   1440
         TabIndex        =   5
         Top             =   195
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   1080
      ScaleHeight     =   15
      ScaleWidth      =   4575
      TabIndex        =   3
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmSplash.frx":0396
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The way of capturing video using the computer monitor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00733C00&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1710
      Width           =   4575
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   600
      Picture         =   "frmSplash.frx":0A20
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H002ED7AF&
      BorderWidth     =   2
      Height          =   2775
      Left            =   -1800
      Shape           =   3  'Circle
      Top             =   -240
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H002ED7AF&
      BorderWidth     =   2
      Height          =   2775
      Left            =   -1600
      Shape           =   3  'Circle
      Top             =   -960
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H002ED7AF&
      BorderWidth     =   2
      Height          =   2775
      Left            =   -1680
      Shape           =   3  'Circle
      Top             =   -600
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " version 1.1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monitor Cam"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   1050
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
frm1.Show
Unload Me
End Sub
