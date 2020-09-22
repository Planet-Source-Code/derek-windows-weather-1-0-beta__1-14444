VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{4F6FBF88-E23D-4958-91E1-D0EE647DD9A3}#1.0#0"; "LISTBOXOCX.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Weather"
   ClientHeight    =   3495
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2790
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   600
      Top             =   5520
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Downloading Information"
      Top             =   3000
      Width           =   2535
   End
   Begin Project2.CustomListBox list1 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5106
      BackColor       =   12632256
      BeginProperty FontInfo {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      Graphical       =   0   'False
      Picture         =   "Form1.frx":030A
      ScrollBarBackColor=   12632256
      ScrollBarBorderColor=   8421504
      SelBoxColor     =   14737632
      Sorted          =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   840
      Picture         =   "Form1.frx":0326
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Width           =   495
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   360
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnufileexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuconfig 
      Caption         =   "Configure"
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "Options"
      Begin VB.Menu mnuoo 
         Caption         =   "Auto On"
      End
      Begin VB.Menu mnuooo 
         Caption         =   "Auto Off"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim co: Dim first: Dim second: Dim change
Dim third: Dim fourth: Dim fifth: Dim six
Dim seven: Dim eight: Dim nine:

Private Sub Form_Load()
On Error Resume Next
list1.AddImage Picture1.Picture
Me.Move 1500, 1500: Me.Show
co = 0: Call read_files
End Sub

Sub read_files()
On Error Resume Next
x = check(App.Path + "\history.dll")
If x = True Then
Call compile_list
Else
Open App.Path + "\history.dll" For Output As #1
Print #1, "22098"
Close #1
Call download_list
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

Sub decode_File()
On Error Resume Next
Let co = co + 1
Open App.Path + "\list.dll" For Input As #1
Input #1, a$, b$, c$, d$, e$, f$, g$, h$, i$, j$:
If co = 1 Then Text1.Text = a$
If co = 2 Then Text1.Text = b$
If co = 3 Then Text1.Text = c$
If co = 4 Then Text1.Text = d$
If co = 5 Then Text1.Text = e$
If co = 6 Then Text1.Text = f$
If co = 7 Then Text1.Text = g$
If co = 8 Then Text1.Text = h$
If co = 9 Then Text1.Text = i$
If co = 10 Then Text1.Text = j$
Close #1
End Sub

Private Sub compile_list()
On Error Resume Next
decode_File
Dim Search As String: Dim Spot As Integer
Dim Spot2 As Integer: Dim Text As String
Search = "01": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "02")
first = Mid$(Text, Spot, Spot2 - Spot)
Search = "02": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "03")
list1.AddItem Mid$(Text, Spot, Spot2 - Spot) & " " & Left(Time, 4), list1.Picture
decode_File
Search = "04": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "05")
second = Mid$(Text, Spot, Spot2 - Spot)
Search = "05": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "06")
list1.AddItem Mid$(Text, Spot, Spot2 - Spot) & " " & Left(Time, 4), list1.Picture
decode_File
Search = "07": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "08")
third = Mid$(Text, Spot, Spot2 - Spot)
Search = "08": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "09")
list1.AddItem Mid$(Text, Spot, Spot2 - Spot) & " " & Left(Time, 4), list1.Picture
decode_File
Search = "10": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "11")
fourth = Mid$(Text, Spot, Spot2 - Spot)
Search = "11": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "12")
list1.AddItem Mid$(Text, Spot, Spot2 - Spot) & " " & Left(Time, 4), list1.Picture
decode_File
Search = "13": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "14")
fifth = Mid$(Text, Spot, Spot2 - Spot)
Search = "14": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "15")
list1.AddItem Mid$(Text, Spot, Spot2 - Spot) & " " & Left(Time, 4), list1.Picture
decode_File
Search = "16": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "17")
six = Mid$(Text, Spot, Spot2 - Spot)
Search = "17": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "18")
list1.AddItem Mid$(Text, Spot, Spot2 - Spot) & " " & Left(Time, 4), list1.Picture
decode_File
Search = "19": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "20")
seven = Mid$(Text, Spot, Spot2 - Spot)
Search = "20": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "21")
list1.AddItem Mid$(Text, Spot, Spot2 - Spot) & " " & Left(Time, 4), list1.Picture
decode_File
Search = "22": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "23")
eight = Mid$(Text, Spot, Spot2 - Spot)
Search = "23": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "24")
list1.AddItem Mid$(Text, Spot, Spot2 - Spot) & " " & Left(Time, 4), list1.Picture
decode_File
Search = "25": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "26")
nine = Mid$(Text, Spot, Spot2 - Spot)
Search = "26": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "27")
list1.AddItem Mid$(Text, Spot, Spot2 - Spot) & " " & Left(Time, 4), list1.Picture
decode_File
Search = "27": Text = Text1.Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "28")
Me.Enabled = True: Text2 = "Selected Files Downloaded"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_DblClick()
On Error Resume Next
Let change = list1.ListIndex + 1
If change = 1 Then Let Inet.Tag = first: Form1.Tag = "2"
If change = 2 Then Let Inet.Tag = second: Form1.Tag = "1"
If change = 3 Then Let Inet.Tag = third: Form1.Tag = "1"
If change = 4 Then Let Inet.Tag = fourth: Form1.Tag = "1"
If change = 5 Then Let Inet.Tag = fifth: Form1.Tag = "1"
If change = 6 Then Let Inet.Tag = six: Form1.Tag = "2"
If change = 7 Then Let Inet.Tag = seven: Form1.Tag = "1"
If change = 8 Then Let Inet.Tag = eight: Form1.Tag = "1"
If change = 9 Then Let Inet.Tag = nine: Form1.Tag = "1"
If Left$(Inet.Tag, 9) = "internal:" Then Call con: Exit Sub
Text2.Text = "Downloading Selected File"
Me.Enabled = False: Call download(Inet.Tag)
End Sub

Sub con()
On Error Resume Next
Me.Enabled = False
Text2.Text = "Downloading Selected File"
Open App.Path + "\zipcode.dll" For Input As #1
Input #1, zip$: Close #1:
Dim Text As String
Dim Search As String
Dim Spot As Integer
Dim Spot2 As Integer
Dim Spot3 As Integer
Dim Spot4 As Integer
Dim text22 As String
Search = "<FONT FACE=""Arial, Helvetica, Chicago, Sans Serif"" SIZE=3><B>"
Search2 = "<FONT FACE=""Arial, Helvetica, Chicago, Sans Serif"" SIZE=2"
Text = Inet.OpenURL("http://www.weather.com/weather/us/zips/" & zip$ & ".html")
text22 = Text
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "</B>")
Spot3 = InStr(1, text22, Search2) + Len(Search)
Spot4 = InStr(Spot3, text22, "&")
temp = Right(Mid$(text22, Spot3, Spot4 - Spot3), 2)
sky = Mid$(Text, Spot, Spot2 - Spot)
Form4.Text1.Text = sky
Form4.Text2.Text = tempco
Form4.Command1.Value = True
Text2.Text = "Selected File Downloaded"
Me.Enabled = True
Form4.Show
End Sub

Sub download(url As String)
On Error Resume Next
Dim bytes() As Byte
bytes() = Inet.OpenURL(url, icByteArray)
Open App.Path + "\output.dll" For Binary As #1
Put #1, , bytes()
Close #1
Text2.Text = "Selected File Downloaded"
Form2.Image1.Picture = LoadPicture(App.Path + "\output.dll")
If Form1.Tag = "2" Then Form2.Image1.Top = -1660: Form2.Image1.Left = -1420
If Form1.Tag = "1" Then Form2.Image1.Top = -350: Form2.Image1.Left = -720
Form2.Image1.Refresh
Form2.Show
Kill App.Path + "\output.dll"
Me.Enabled = True
End Sub

Private Sub mnuconfig_Click()
Form3.Show
End Sub

Private Sub mnufileexit_Click()
End
End Sub

Private Sub mnuoo_Click()
Timer1.Enabled = True
End Sub

Private Sub mnuooo_Click()
Timer1.Enabled = False
End Sub

Private Sub Text2_GotFocus()
Text1.SetFocus
End Sub

Sub download_list()
On Error GoTo nonet
Me.Show: Me.Enabled = False
Dim bytes() As Byte
bytes() = Inet.OpenURL("http://derekgregg.tripod.com/1.js", icByteArray)
Open App.Path + "\list.dll" For Binary As #1
Put #1, , bytes(): Close #1
Call compile_list
Exit Sub
nonet:
Text2.Text = "Unable to Find Connection"
End Sub

Private Sub Timer1_Timer()
List1_DblClick
End Sub
