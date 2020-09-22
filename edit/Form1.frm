VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Weather 1.0"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   360
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "Save Changes"
      Top             =   100
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   350
      Left            =   1560
      TabIndex        =   4
      Text            =   "IL"
      ToolTipText     =   "New State"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   350
      Left            =   1080
      TabIndex        =   3
      Text            =   "CHI"
      ToolTipText     =   "New City"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   350
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Listed State"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   350
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Listed City"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "Text5"
      Top             =   1680
      Width           =   150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Text5.SetFocus
Kill App.Path + "\location.dll"
Open App.Path + "\location.dll" For Output As #1
Print #1, Text3.Text
Print #1, Text4.Text
Close #1
Call change
End Sub

Sub change()
Dim address
address = LCase(Drive1.Drive + "\Program Files\Accessories\WORDPAD.EXE list.dll")
Shell address, vbNormalFocus
SendKeys "^h"
SendKeys LCase(Text1.Text)
SendKeys "{tab}"
SendKeys LCase(Text3.Text)
SendKeys "%{c}"
SendKeys "%{a}"
SendKeys "{ENTER}"
SendKeys "{tab}"
SendKeys UCase(Text2.Text)
SendKeys "{tab}"
SendKeys UCase(Text4.Text)
SendKeys "%{a}"
SendKeys "{ENTER}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{ENTER}"
SendKeys "%f"
SendKeys "x"
SendKeys "{ENTER}"
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Move 5000, 2000
Open App.Path + "\location.dll" For Input As #1
Input #1, city$
Input #1, state$
Close #1
Text1.Text = city$
Text2.Text = state$
x = check(App.Path + "\list.dll")
If x = False Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub

Private Function check(stl)
On Error GoTo NoFile:
Open stl For Input As #1
Close #1
check = True
Exit Function
NoFile:
check = False
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

