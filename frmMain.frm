VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mousemat Installer"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   735
      ScaleWidth      =   7695
      TabIndex        =   25
      Top             =   0
      Width           =   7695
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Mat Installer Version 5.1.2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         Top             =   120
         Width           =   5775
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   150
      ImageHeight     =   150
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1244
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3970
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":434D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E0F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7695
      Begin VB.Label lblText 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   1080
         TabIndex        =   7
         Top             =   960
         Width           =   5415
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   7695
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   3840
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblText2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   1440
         TabIndex        =   23
         Top             =   840
         Width           =   4935
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   7695
      Begin VB.OptionButton optRight 
         Caption         =   "Centered (not recommended)"
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   22
         Top             =   3720
         Width           =   3855
      End
      Begin VB.OptionButton optRight 
         Caption         =   "At the left side of your keyboard"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   21
         Top             =   3000
         Width           =   3855
      End
      Begin VB.OptionButton optRight 
         Caption         =   "At the right side of your keyboard"
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   20
         Top             =   2280
         Value           =   -1  'True
         Width           =   3855
      End
      Begin VB.Label lblText1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         TabIndex        =   19
         Top             =   960
         Width           =   3615
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   7695
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   6000
         TabIndex        =   18
         Text            =   "500"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6000
         TabIndex        =   17
         Text            =   "500"
         Top             =   3480
         Width           =   1335
      End
      Begin VB.OptionButton optMat 
         Caption         =   "Mat free"
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
         Index           =   7
         Left            =   6000
         TabIndex        =   16
         Top             =   4440
         Width           =   1335
      End
      Begin VB.OptionButton optMat 
         Caption         =   "Mat blanc (almost)"
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
         Index           =   6
         Left            =   4080
         TabIndex        =   15
         Top             =   4440
         Width           =   1695
      End
      Begin VB.OptionButton optMat 
         Caption         =   "Mat home"
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
         Index           =   5
         Left            =   2160
         TabIndex        =   14
         Top             =   4440
         Width           =   1455
      End
      Begin VB.OptionButton optMat 
         Caption         =   "Mat oval"
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
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   4440
         Width           =   1335
      End
      Begin VB.OptionButton optMat 
         Caption         =   "Mat graphical"
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
         Left            =   6000
         TabIndex        =   12
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton optMat 
         Caption         =   "Mat edged"
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
         Left            =   4080
         TabIndex        =   11
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton optMat 
         Caption         =   "Mat calculator"
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
         Left            =   2160
         TabIndex        =   10
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton optMat 
         Caption         =   "Mat standard"
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
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Please enter the width/height of your mat in cm."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6000
         TabIndex        =   27
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Image imgMat 
         Height          =   615
         Index           =   6
         Left            =   3960
         Top             =   2640
         Width           =   735
      End
      Begin VB.Image imgMat 
         Height          =   615
         Index           =   5
         Left            =   2040
         Top             =   2640
         Width           =   735
      End
      Begin VB.Image imgMat 
         Height          =   615
         Index           =   4
         Left            =   120
         Top             =   2640
         Width           =   735
      End
      Begin VB.Image imgMat 
         Height          =   615
         Index           =   3
         Left            =   5880
         Top             =   240
         Width           =   735
      End
      Begin VB.Image imgMat 
         Height          =   615
         Index           =   2
         Left            =   3960
         Top             =   240
         Width           =   735
      End
      Begin VB.Image imgMat 
         Height          =   615
         Index           =   1
         Left            =   2040
         Top             =   240
         Width           =   735
      End
      Begin VB.Image imgMat 
         Height          =   615
         Index           =   0
         Left            =   120
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   3
      Left            =   0
      TabIndex        =   28
      Top             =   720
      Width           =   7695
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00808080&
         Height          =   3015
         Left            =   600
         ScaleHeight     =   2955
         ScaleWidth      =   6555
         TabIndex        =   32
         Top             =   1440
         Width           =   6615
         Begin VB.Shape Shape2 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   3120
            Shape           =   3  'Circle
            Top             =   1320
            Width           =   375
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   2775
            Left            =   140
            Shape           =   4  'Rounded Rectangle
            Top             =   90
            Width           =   6255
         End
      End
      Begin VB.Label lblY 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblx 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3360
         TabIndex        =   30
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1080
         TabIndex        =   29
         Top             =   480
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private StageCntr As Byte

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFinish_Click()
    Dim i As Integer
    Me.ProgressBar1.Max = 100
    For i = 1 To 100
        Me.ProgressBar1.Value = i
        Sleep 2
    Next
    Me.ProgressBar1.Value = 0
    For i = 1 To 100
        Me.ProgressBar1.Value = i
        If (i > 0) And (i <= 30) Then
            Me.lblText2.Caption = "Installing drivers.. please wait"
            Me.lblText2.Refresh
        ElseIf (i > 30) And (i <= 60) Then
            Me.lblText2.Caption = "Registering files and positioning mouse to mat"
            Me.lblText2.Refresh
        ElseIf (i > 60) Then
            Me.lblText2.Caption = "Some time to think about the PSC users ;-)"
            Me.lblText2.Refresh
        End If
        Sleep 20
    Next
    Me.lblText2.Caption = "Thank you for using our product, all we care about is you !!!"
    MsgBox "Mouse Mat Installation succeeded..", vbInformation, "Mouse mat Install OK"
    Me.ProgressBar1.Value = 0
End Sub

Private Sub cmdNext_Click()
    StageCntr = StageCntr + 1
    DoStageChange
End Sub

Private Sub cmdPrevious_Click()
    StageCntr = StageCntr - 1
    DoStageChange
End Sub

Private Sub Form_Load()
    StageCntr = 0
    InitialLoad
    DoStageChange
End Sub

Private Sub DoStageChange()
    Dim i As Byte
    For i = 0 To Me.fraMain.Count - 1
        Me.fraMain(i).Visible = False
    Next
    Me.fraMain(StageCntr).Visible = True
    Me.cmdPrevious.Enabled = (StageCntr > 0)
    Me.cmdNext.Enabled = (StageCntr < Me.fraMain.Count - 1)
    Me.cmdFinish.Enabled = (StageCntr = Me.fraMain.Count - 1)
End Sub

Private Sub InitialLoad()
    Me.lblText.Caption = "Welcome to the mouse mat installation program." & vbCrLf & _
        "Press next to begin the installation" & vbCrLf & vbCrLf & _
        "If the mouse mat is well installed you will notice the big change."
    Me.lblText1.Caption = "What is the position of your mousemat ?"
    Me.lblText2.Caption = "Mouse mat installer has now all information needed for installation of your mousemat" & vbCrLf & vbCrLf & _
        "Press finish to begin the installation"
    Me.Label3.Caption = "Sweet spot configuration" & vbCrLf & vbCrLf & _
        "Center your mouse on the mat and hit the red spot on the screen"
    Dim i As Byte
    For i = 0 To Me.ImageList1.ListImages.Count - 1
        Me.imgMat(i).Picture = Me.ImageList1.ListImages(i + 1).Picture
    Next
End Sub

Private Sub optMat_Click(Index As Integer)
    If Index = 7 Then
        Me.Text1.SetFocus
        Me.Text1.SelStart = 0
        Me.Text1.SelLength = Len(Me.Text1.Text)
    End If
End Sub

Private Sub Picture2_Click()
    MsgBox "Mouseposition set", vbInformation, "Mouse SET"
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblx.Caption = "X " & X
    Me.lblY.Caption = "Y " & Y
End Sub
