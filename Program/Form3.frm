VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Weather"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   100
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Location"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   100
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   350
      Left            =   120
      MaxLength       =   5
      TabIndex        =   1
      Top             =   130
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Text2.SetFocus
Kill App.Path + "\zipcode.dll"
Open App.Path + "\zipcode.dll" For Output As #1
Print #1, Text1.text
Close #1
Me.Hide
End Sub

Private Sub Command2_Click()
Text2.SetFocus
Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Move 5000, 2000
Open App.Path + "\zipcode.dll" For Input As #1
Input #1, zip$
Text1.text = zip$
Close #1
Me.Show
End Sub
