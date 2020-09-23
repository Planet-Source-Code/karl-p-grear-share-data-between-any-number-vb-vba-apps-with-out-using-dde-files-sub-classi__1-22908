VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Get"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "ShareData.frx":0000
      Top             =   720
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send Data"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Data"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************
'* Paste in Form    *
'* Add 2 Text Boxes *
'* Add 4 Buttons    *
'********************
Dim MemorySpace1 As New clsMemoryMap
Dim MemorySpace2 As New clsMemoryMap


Private Sub Command1_Click()
'Get memory.
    Text1.Text = MemorySpace1.Peek
End Sub

Private Sub Command2_Click()
'Store memory.
    MemorySpace1.Poke (Text1.Text)
End Sub


Private Sub Command3_Click()
'Store memory.
    MemorySpace2.Poke (Text2.Text)
End Sub

Private Sub Command4_Click()
'Get memory.
    Text2.Text = MemorySpace2.Peek
End Sub

Private Sub Form_Load()
    MemorySpace1.OpenMemory ("Space1")
    MemorySpace2.OpenMemory ("Space2")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MemorySpace1.CloseMemory
    MemorySpace2.CloseMemory
End Sub
