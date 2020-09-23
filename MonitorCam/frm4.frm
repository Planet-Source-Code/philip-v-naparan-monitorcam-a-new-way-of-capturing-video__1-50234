VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form frm4 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6495
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   4455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   5760
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   270
      Picture         =   "frm4.frx":000C
      ScaleHeight     =   4260
      ScaleWidth      =   3885
      TabIndex        =   9
      Top             =   720
      Width           =   3920
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SF1 
         Height          =   4290
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   3915
         _cx             =   4201218
         _cy             =   4201871
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
         Stacking        =   "below"
      End
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
      TabIndex        =   8
      Top             =   6480
      Width           =   4455
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H002ED7AF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   4440
      ScaleHeight     =   6255
      ScaleWidth      =   15
      TabIndex        =   7
      Top             =   240
      Width           =   15
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H002ED7AF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   0
      ScaleHeight     =   6255
      ScaleWidth      =   15
      TabIndex        =   6
      Top             =   240
      Width           =   15
   End
   Begin MonitorCam.NaparanButton NaparanButton2 
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   5520
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1085
      Caption         =   "Click here to view your self Live in the video using your monitor !"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Status: VCCC is waiting for signal to capture a                  video."
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
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   3975
   End
End
Attribute VB_Name = "frm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MsgBox "VCCC ( Video Capturing Capability Chip ) is ready for capturing video from your monitor.", vbInformation, "Ready !"
SF1.Movie = App.Path & "\support.bin"
SF1.Stop
End Sub

Private Sub NaparanButton1_Click()
If NaparanButton2.Enabled = True Then
    MsgBox "Click the button bellow first before clicking this button.", vbInformation, "Instruction"
    Exit Sub
End If
MsgBox "Your soo cute in the video. You know what? your species are INDANGER !", vbInformation, "About you !"
MsgBox "1. Don't be MAD this is only a joke !!!!" & vbCrLf & _
       "2. Advanced Marry Christmas !!!" & vbCrLf & _
       "3. Please don't forget to vote this application at http://www.pscode.com !!!" _
       , vbInformation, "Message from the Author:"
End
End Sub
Private Sub NaparanButton2_Click()
SF1.Visible = True
Label2.Caption = " Status: Live video capturing from your self."
SF1.Play
NaparanButton2.Enabled = False
End Sub

Private Sub SF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
End
End Sub

