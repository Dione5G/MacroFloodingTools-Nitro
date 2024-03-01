VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "       About"
   ClientHeight    =   3165
   ClientLeft      =   -15
   ClientTop       =   210
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":57E2
   ScaleHeight     =   5.119
   ScaleMode       =   0  'User
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Youtube"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Instagram"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Github"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   3840
      Picture         =   "Form2.frx":89EA
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Discord - dione5g"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1200
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Any problem you can tell me on one of my social networks by clicking on one of the contact buttons"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "- Each box has a limit of 100 characters so you don't have problems when casting Flood."
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "- In this version use copy and paste to send messages."
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MFT Nitro - 2024"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dione5G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   3720
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
Form2.Hide
End Sub
Private Sub Command2_Click()
ShellExecute Me.hWnd, "open", "https://github.com/Dione5G", "", "C:\\", SW_SHOWNORMAL
End Sub
Private Sub Command3_Click()
ShellExecute Me.hWnd, "open", "https://www.instagram.com/Dione5G", "", "C:\\", SW_SHOWNORMAL
End Sub
Private Sub Command4_Click()
ShellExecute Me.hWnd, "open", "https://www.youtube.com/channel/UC7L97ZkpVRVceO-sqE9NhJw", "", "C:\\", SW_SHOWNORMAL
End Sub

