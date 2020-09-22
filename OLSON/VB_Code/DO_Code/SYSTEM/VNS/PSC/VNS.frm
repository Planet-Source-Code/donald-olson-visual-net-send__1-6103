VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00800000&
   Caption         =   "Visual Net Send"
   ClientHeight    =   3210
   ClientLeft      =   4200
   ClientTop       =   3720
   ClientWidth     =   5280
   Icon            =   "VNS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5280
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      Picture         =   "VNS.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2760
      TabIndex        =   9
      Top             =   360
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4800
      Top             =   3360
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00800000&
      Caption         =   "Use Customized Header"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   2880
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Saved 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Picture         =   "VNS.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Saved 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Picture         =   "VNS.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Picture         =   "VNS.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\VNS\VNS.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Data"
      Top             =   4440
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      Picture         =   "VNS.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   765
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   -240
      Picture         =   "VNS.frx":198C
      Top             =   690
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Select Group:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Awaiting Message Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Send to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu Exit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu EditList 
      Caption         =   "&Edit List"
   End
   Begin VB.Menu Custom1 
      Caption         =   "&Customize"
   End
   Begin VB.Menu About1 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varClicks As Integer
Private Sub About1_Click()
About2.Show
End Sub

Private Sub Combo1_Click()

'Set up database variables
Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Data")
Set GetData3 = GetData1.OpenRecordset("Groups")

'Empty Name Combo Box
Combo2.Clear

'Populate Combo Box for Names
If Combo1 = "All" Then
GetData2.MoveFirst
Do Until GetData2.EOF
    Combo2.AddItem (GetData2!Name)
    GetData2.MoveNext
Loop
Exit Sub
End If

GetData2.MoveFirst
Do Until GetData2.EOF
    If GetData2!Group = Combo1 Then
        Combo2.AddItem (GetData2!Name)
    End If
    GetData2.MoveNext
Loop

End Sub

Private Sub Combo2_Click()

Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Data")
GetData2.MoveFirst
Do Until GetData2.EOF
    If GetData2!Name = Combo2 Then
        varDesc = GetData2!Description
        Label4 = "Preparing message for " & varDesc
    End If
    GetData2.MoveNext
Loop

End Sub

Private Sub Command1_Click()

'Check for valid entries
If Combo2 = "" Then
    MsgBox "Receipient Must Be Selected!", 0, "Select Receipient Name"
    Exit Sub
End If

If Text1 = "" Then
    MsgBox "You must enter a message!", 0, "Enter Message"
    Exit Sub
End If

'Gather database information
Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Data")
Set GetData3 = GetData1.OpenRecordset("Custom")
Set GetData4 = GetData1.OpenRecordset("Previous")

GetData2.MoveFirst
CompName = "blank"
Do Until GetData2.EOF
    If GetData2!Name = Combo2 Then
        CompName = GetData2!Computer
        RecName = GetData2!Name
    End If
    GetData2.MoveNext
Loop

'Build Net Send Statement

If Check1.Value = 1 Then
    varMessage = GetData3!Header
    varMessage = varMessage + Chr(13) + Chr(10) + Chr(13) + Chr(10)
    varMessage = varMessage & Text1
Else:
    varMessage = Text1
End If

Message1 = "net send " & CompName & " " & varMessage

'Send Message
Call Shell(Message1)

'Animate Envelope
Timer1.Interval = 20

'Add to Previous Messages
GetData4.AddNew
GetData4!Message = Text1
GetData4.Update

varClicks = 0

'Update Label
MSG1 = "Message Sent to " & RecName
Label4 = MSG1




End Sub

Private Sub Command2_Click()
Text1 = ""
Label4 = "Awaiting Message Entry"
varClicks = 0
End Sub

Private Sub Command3_Click()

Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Custom")

GetData2.MoveFirst
Text1 = ""
Text1 = GetData2!Message1


End Sub

Private Sub Command4_Click()
Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Custom")

GetData2.MoveFirst
Text1 = ""
Text1 = GetData2!Message2

End Sub

Private Sub Command5_Click()

On Error GoTo Here1

varClicks = varClicks + 1

'Set Database variables
Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Previous")

'Count Records
GetData2.MoveFirst
Do Until GetData2.EOF
    varCount = varCount + 1
    GetData2.MoveNext
Loop

'Get Last Message
GetData2.MoveFirst
For varCounter = 1 To varCount - varClicks
    GetData2.MoveNext
Next
Text1 = GetData2!Message

Exit Sub

Here1:


End Sub

Private Sub Custom1_Click()
Custom.Show
End Sub

Private Sub DBCombo1_Change()

Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Data")
GetData2.MoveFirst
Do Until GetData2.EOF
    If GetData2!Name = DBCombo1 Then
        varDesc = GetData2!Description
        Label4 = "Preparing message for " & varDesc
    End If
    GetData2.MoveNext
Loop

End Sub

Private Sub EditList_Click()
List1.Show

End Sub

Private Sub Exit_Click()
varResults = MsgBox("Are You Sure You Want to Exit?", vbYesNo, "Exiting VNS")
    If varResults = vbYes Then
        End
    End If
End Sub

Private Sub Form_Load()

Timer1.Interval = 0
Image1.Visible = False

Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Custom")
Set GetData3 = GetData1.OpenRecordset("Previous")
Set GetData4 = GetData1.OpenRecordset("Groups")
Set GetData5 = GetData1.OpenRecordset("Data")

GetData2.MoveFirst
Command3.Caption = GetData2!Caption1
Command4.Caption = GetData2!Caption2

Do Until GetData3.EOF
    GetData3.Delete
    GetData3.MoveNext
Loop

'Load up Group Combo Box
Combo1.AddItem ("All")
GetData4.MoveFirst
Do Until GetData4.EOF
    Combo1.AddItem (GetData4!GroupName)
    GetData4.MoveNext
Loop

Combo1.Text = "All"

'Load up Send to Box
GetData5.MoveFirst
Do Until GetData5.EOF
    Combo2.AddItem (GetData5!Name)
    GetData5.MoveNext
Loop


End Sub

Private Sub Text1_Click()

Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Data")
GetData2.MoveFirst
Do Until GetData2.EOF
    If GetData2!Name = DBCombo1 Then
        varDesc = GetData2!Description
        Label4 = "Preparing message for " & varDesc
    End If
    GetData2.MoveNext
Loop

End Sub

Private Sub Timer1_Timer()
If Image1.Left > 5160 Then
    Image1.Left = -240
    Image1.Visible = False
    Timer1.Interval = 0
    Exit Sub
ElseIf Image1.Left = -240 Then
    Image1.Visible = False
    Image1.Move Image1.Left + 150
Else:
    Image1.Visible = True
    Image1.Move Image1.Left + 150
End If
End Sub
