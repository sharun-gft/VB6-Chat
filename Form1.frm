VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ChatWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtOut 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox txtIn 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   6495
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   1234
      LocalPort       =   1234
   End
End
Attribute VB_Name = "ChatWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MessageIn As String
Dim MessageOut As String
Dim UserName As String
Dim PlayMusic As Boolean

Private Declare Function Playwave Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Sub cmdSend_Click()
MessageOut = UserName & " : " & txtOut
Winsock1.SendData MessageOut
txtIn = UserName & " : " & txtOut & vbCrLf & txtIn
End Sub

Private Sub Form_GotFocus()
PlayMusic = False
End Sub

Private Sub Form_Load()

UserName = InputBox("Enter your name: ", "Enter Name")

ChatWindow.Caption = Winsock1.LocalIP & ":" & Winsock1.LocalPort & " - " & UserName

End Sub

Private Sub txtName_Click()
   If txtName = "Enter Name" Then
    txtName = ""
    End If
    
    
End Sub

Private Sub txtName_LostFocus()
If txtName = "" Then
txtName = "Enter Name"
End If

End Sub

Private Sub Form_LostFocus()
PlayMusic = True
End Sub

Private Sub txtOut_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call cmdSend_Click
txtOut = ""
End If


End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData MessageIn

If Left$(MessageIn, Len(UserName)) <> UserName Then
txtIn = MessageIn & vbCrLf & txtIn
If PlayMusic Then
Call Playwave("C:\Users\admin\Music\WAV\Speech On.wav", 3)
End If
End If

End Sub

