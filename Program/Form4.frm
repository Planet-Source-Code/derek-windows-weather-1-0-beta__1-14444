VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Weather 1.0"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   4920
      Width           =   615
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   1200
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   1200
      TabIndex        =   9
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Height          =   135
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Height          =   135
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Time is :"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Temperature is"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reporting zip :"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zip$

Private Sub Command1_Click()
Call decode
End Sub

Private Sub Form_Load()
Me.Move 5000, 2000
Open App.Path + "\zipcode.dll" For Input As #1
Input #1, zip$
Close #1
Exit Sub
End Sub

Sub decode()
On Error Resume Next
Dim con
If Text1.Text = "Haze" Then con = -2
If Text1.Text = "Fog" Then con = -2
If Text1.Text = "Sunny" Then con = -1
If Text1.Text = "Cloudy" Then con = 0
If Text1.Text = "Mostly Cloudy" Then con = 1
If Text1.Text = "Partly Cloudy" Then con = 1
If Text1.Text = "Mostly Sunny" Then con = 1
If Text1.Text = "Partly Sunny" Then con = 1
If Text1.Text = "Showers" Then con = 2
If Text1.Text = "Light Drizzle" Then con = 2
If Text1.Text = "Light Rain" Then con = 2
If Text1.Text = "Drizzle" Then con = 2
If Text1.Text = "Heavy Rain" Then con = 3
If Text1.Text = "Rain" Then con = 3
If Text1.Text = "Light Snow" Then con = 4
If Text1.Text = "Snow Showers" Then con = 4
If Text1.Text = "Scattered Snow Showers" Then con = 4
If Text1.Text = "Heavy Snow" Then con = 5
If Text1.Text = "Snow" Then con = 5
If Text1.Text = "Rain and Snow" Then con = 6
If con = -2 Then url$ = "http://image.weather.com/weather/wx_icons/PFMSforecastTrans/21.gif"
If con = -1 Then url$ = "http://image.weather.com/weather/wx_icons/PFMSforecastTrans/32.gif"
If con = 0 Then url$ = "http://image.weather.com/weather/wx_icons/PFMSforecastTrans/26.gif"
If con = 1 Then url$ = "http://image.weather.com/weather/wx_icons/PFMSforecastTrans/30.gif"
If con = 2 Then url$ = "http://image.weather.com/weather/wx_icons/PFMSforecastTrans/11.gif"
If con = 3 Then url$ = "http://image.weather.com/weather/wx_icons/PFMSforecastTrans/12.gif"
If con = 4 Then url$ = "http://image.weather.com/weather/wx_icons/PFMSforecastTrans/14.gif"
If con = 5 Then url$ = "http://image.weather.com/weather/wx_icons/PFMSforecastTrans/16.gif"
If con = 6 Then url$ = "http://image.weather.com/weather/wx_icons/PFMSforecastTrans/5.gif"
Dim bytes() As Byte
bytes() = Inet1.OpenURL(url$, icByteArray)
Open App.Path + "\output.dll" For Binary As #1
Put #1, , bytes(): Close #1
Picture1.Picture = LoadPicture(App.Path + "\output.dll")
Kill App.Path + "\output.dll"
Form4.Show
Label1.Caption = "Reporting zip : " & zip$
Label2.Caption = "Temperature is : " & Text2.Text
Label3.Caption = "Time is : " & Time
Picture1.ToolTipText = Text1.Text
End Sub
