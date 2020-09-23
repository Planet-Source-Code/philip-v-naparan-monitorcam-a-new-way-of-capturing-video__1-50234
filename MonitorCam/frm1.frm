VERSION 5.00
Begin VB.Form frm1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4815
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   4395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   2280
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H002ED7AF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   4455
      TabIndex        =   9
      Top             =   4800
      Width           =   4455
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H002ED7AF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   4380
      ScaleHeight     =   4815
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   240
      Width           =   15
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H002ED7AF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4695
      ScaleWidth      =   15
      TabIndex        =   7
      Top             =   240
      Width           =   15
   End
   Begin MonitorCam.NaparanButton NaparanButton2 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      Caption         =   "View my self in the monitor camera !"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BackColor       =   16777215
   End
   Begin MonitorCam.NaparanButton NaparanButton1 
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   450
      Caption         =   "X"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   0   'False
      BackColor       =   3069871
      ForeColor       =   0
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   4455
      TabIndex        =   2
      Top             =   300
      Width           =   4455
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H002ED7AF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Monitor Cam version 1.1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   50
         Width           =   2655
      End
   End
   Begin MonitorCam.NaparanButton NaparanButton3 
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      Caption         =   "About the inventor"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      BackColor       =   16777215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Read the Information Bellow:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm1.frx":068A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   1080
      TabIndex        =   10
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frm1.frx":07C1
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Read the Information Bellow:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub NaparanButton1_Click()
End
End Sub

Private Sub NaparanButton2_Click()
frm2.Show vbModal
frm3.Show vbModal
frm4.Show vbModal
End Sub

Private Sub NaparanButton3_Click()
frmabout.Show vbModal
End Sub

Private Sub Timer1_Timer()
Label3.Visible = Not Label3.Visible
Label4.Visible = Not Label4.Visible
End Sub
