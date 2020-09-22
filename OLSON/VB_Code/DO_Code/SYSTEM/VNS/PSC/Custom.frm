VERSION 5.00
Begin VB.Form Custom 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customize VNS"
   ClientHeight    =   3285
   ClientLeft      =   4575
   ClientTop       =   3795
   ClientWidth     =   4680
   Icon            =   "Custom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Close"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Save Settings"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Saved Messages"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Text            =   "Saved Message #2"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Text            =   "Saved 2"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Text            =   "Saved Message #1"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Text            =   "Saved 1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800000&
         Caption         =   "Button #1 Message:"
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
         TabIndex        =   12
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "Button #1 Message:"
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
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Button #2 Caption:"
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
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Button #1 Caption:"
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
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Message Sent by VNS User"
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Message Header:"
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
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Custom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Custom")

GetData2.MoveFirst
GetData2.Edit
GetData2!Header = Text1
GetData2!Caption1 = Text2
GetData2!Message1 = Text3
GetData2!Caption2 = Text4
GetData2!Message2 = Text5
GetData2.Update

Main.Command3.Caption = Text2
Main.Command4.Caption = Text4

MsgBox "Customized Data Updated", 0, "Data Updated"

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

'Set up database variables
Set GetData1 = OpenDatabase("c:\Program Files\VNS\VNS.mdb")
Set GetData2 = GetData1.OpenRecordset("Custom")

GetData2.MoveFirst
Text1 = GetData2!Header
Text2 = GetData2!Caption1
Text3 = GetData2!Message1
Text4 = GetData2!Caption2
Text5 = GetData2!Message2

End Sub
