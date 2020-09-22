VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form List1 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VNS Receipient List"
   ClientHeight    =   4905
   ClientLeft      =   4380
   ClientTop       =   2940
   ClientWidth     =   4965
   Icon            =   "List1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4965
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\VNS\VNS.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Data"
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\VNS\VNS.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Groups"
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Caption         =   "Add / Edit Recipient List"
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
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4695
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "Update List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2880
         Width           =   2775
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "List1.frx":0442
         DataSource      =   "Data2"
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Name"
         Text            =   ""
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Edit Existing Recipient"
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
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Add New Recipient"
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
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "List1.frx":0456
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "GroupName"
         Text            =   ""
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "Machine Name:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Group:"
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
         Left            =   2400
         TabIndex        =   11
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Name:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Description:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Add a New Group Name"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "Add Group Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Names"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Groups"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Menu close2 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "List1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close2_Click()
Unload Me
End Sub


Private Sub Command1_Click()

Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Data")

'test for entries
If DBCombo1 = "" Then
    MsgBox "Group Must Be Given!", 0, "Enter Data"
    Exit Sub
End If
If Text4 = "" Then
    MsgBox "Machine Name Must Be Given!", 0, "Enter Data"
    Exit Sub
End If
If Text3 = "" Then
    MsgBox "Description Must Be Given!", 0, "Enter Data"
    Exit Sub
End If

'Add New Record if that choice was selected
If Option1(0).Value = True Then
    GetData2.AddNew
    GetData2!Name = Text2
    GetData2!Group = DBCombo1
    GetData2!Computer = Text4
    GetData2!Description = Text3
    GetData2.Update
    Main.Combo2.AddItem (Text2)
    MsgBox "Recipient Data Has Been Added!", 0, "Recipient Added"
Else:
    GetData2.MoveFirst
    Do Until GetData2.EOF
        If GetData2!Name = DBCombo2 Then
            GetData2.Edit
            GetData2!Name = DBCombo2
            GetData2!Group = DBCombo1
            GetData2!Computer = Text4
            GetData2!Description = Text3
            GetData2.Update
        End If
    GetData2.MoveNext
    Loop
    MsgBox "Recipient Data Has Been Updated!", 0, "Recipient Updated"
End If

Main.Data1.Refresh
DBCombo2.Refresh

End Sub

Private Sub Command2_Click()

'Set up database variables
Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Groups")

'test for entry
If Text1 = "" Then
    MsgBox "Group Name Must Be Given!", 0, "Enter Data"
    Exit Sub
End If

GetData2.AddNew
GetData2!GroupName = Text1
GetData2.Update
Main.Combo1.AddItem (Text1)

MsgBox "Group Name Added", 0, "Group Added"

Data1.Refresh
DBCombo1.Refresh

End Sub

Private Sub DBCombo2_change()
'Gather database information
Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Data")

GetData2.MoveFirst
Do Until GetData2.EOF
    If GetData2!Name = DBCombo2 Then
        Text4 = GetData2!Computer
        Text3 = GetData2!Description
        DBCombo1 = GetData2!Group
        Exit Sub
    Else:
        GetData2.MoveNext
    End If
Loop
        
End Sub


Private Sub Form_Load()
DBCombo2.Visible = False
Option1(0).Value = True
Text2.Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).Value = True Then
    DBCombo2.Visible = False
    Text2.Visible = True
Else:
    DBCombo2.Visible = True
    Text2.Visible = False
End If


DBCombo1 = ""
DBCombo2 = ""
Text3 = ""
Text4 = ""
Text2 = ""
End Sub

