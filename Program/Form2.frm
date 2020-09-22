VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Weather Graphic Viewer 1.0 - File Opened"
   ClientHeight    =   3870
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4830
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3585
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   3975
         Left            =   -720
         Top             =   -360
         Width           =   5655
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnufileee 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnup 
      Caption         =   "Picture Configuration"
      Begin VB.Menu mnup1 
         Caption         =   "Pan North"
      End
      Begin VB.Menu mnup2 
         Caption         =   "Pan South"
      End
      Begin VB.Menu mnup3 
         Caption         =   "Pan East"
      End
      Begin VB.Menu mnup4 
         Caption         =   "Pan West"
      End
   End
   Begin VB.Menu mnuo 
      Caption         =   "Print This Picture"
   End
   Begin VB.Menu mnuh 
      Caption         =   "Reset Picture"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox Image1.Top
MsgBox Image1.Left
End Sub

Private Sub Form_Load()
Me.Move 5000, 1000
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Hide
End Sub

Private Sub Form_Terminate()
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
End Sub

Private Sub mnufileee_Click()
Form2.Hide
End Sub

Private Sub mnuh_Click()
If Form1.Tag = "2" Then Form2.Image1.Top = -1660: Form2.Image1.Left = -1420
If Form1.Tag = "1" Then Form2.Image1.Top = -350: Form2.Image1.Left = -720
End Sub

Private Sub mnuo_Click()
On Error Resume Next
Printer.Print Image1.Picture
Form2.Hide
Exit Sub
End Sub

Private Sub mnup1_Click()
On Error Resume Next
Image1.Top = Image1.Top + 100
End Sub

Private Sub mnup2_Click()
On Error Resume Next
Image1.Top = Image1.Top - 100
End Sub

Private Sub mnup3_Click()
On Error Resume Next
Image1.Left = Image1.Left - 100
End Sub

Private Sub mnup4_Click()
On Error Resume Next
Image1.Left = Image1.Left + 100
End Sub
