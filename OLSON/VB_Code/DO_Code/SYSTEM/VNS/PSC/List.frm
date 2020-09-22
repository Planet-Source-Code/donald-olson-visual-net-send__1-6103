VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form List 
   BackColor       =   &H00800000&
   Caption         =   "Edit List"
   ClientHeight    =   2850
   ClientLeft      =   8355
   ClientTop       =   3240
   ClientWidth     =   3705
   Icon            =   "List.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3705
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "List.frx":0442
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "List.frx":0452
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Data Data1 
      Caption         =   "List Data"
      Connect         =   "Access"
      DatabaseName    =   "C:\VNS\VNS.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Data"
      Top             =   3000
      Width           =   2475
   End
   Begin VB.Menu Close1 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close1_Click()
Unload Me
Main.Data1.Refresh
End Sub

