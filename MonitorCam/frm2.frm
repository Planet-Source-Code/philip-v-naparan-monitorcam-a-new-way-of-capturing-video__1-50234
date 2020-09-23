VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm2 
   Appearance      =   0  'Flat
   Caption         =   "Please wait..."
   ClientHeight    =   1080
   ClientLeft      =   270
   ClientTop       =   1710
   ClientWidth     =   5970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   1080
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2160
      Top             =   0
   End
   Begin VB.Label Label2 
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frm2.frx":000C
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Gathering monitor information..."
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
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim c As Byte


Private Sub Form_Load()
ProgressBar1.Min = 0
ProgressBar1.Max = 100
End Sub

Private Sub Timer1_Timer()
c = c + 1
Label2.Caption = c & "%"
ProgressBar1.Value = c
If c = 32 Then
    Label1.Caption = "Searching VCCC chip hardware location..."
    Timer2.Enabled = True
    Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
c = c + 1
Label2.Caption = c & "%"
ProgressBar1.Value = c
If c = 37 Then
    Label1.Caption = "Connecting to VCCC chip core..."
    Timer3.Enabled = True
    Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
c = c + 1
Label2.Caption = c & "%"
ProgressBar1.Value = c
If c = 77 Then
    Label1.Caption = "Updating VCCC chip core..."
    Timer4.Enabled = True
    Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
c = c + 1
Label2.Caption = c & "%"
ProgressBar1.Value = c
If c = 85 Then
    Label1.Caption = "Preparing VCCC chip core..."
    Timer5.Enabled = True
    Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
c = c + 1
Label2.Caption = c & "%"
ProgressBar1.Value = c
If c = 100 Then
    Unload Me
End If
End Sub
