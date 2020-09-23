VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agent Example"
   ClientHeight    =   765
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMove 
      Caption         =   "&Move"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdGreet 
      Caption         =   "&Greet"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSpeak 
      Caption         =   "&Speak"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   4080
      Top             =   120
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Begin VB.Menu mnuWrite 
         Caption         =   "&Write"
      End
      Begin VB.Menu mnuRead 
         Caption         =   "&Read"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ms Agent Example
'================
'Very easy to undestand and use
'Use this withe Ms Agent tutorial availble at
'www.planet-source-code.com/vb
'Tutorial and Example code by Mahangu Weerasinghe
'vbdude777@email.com
'If you like this example please take time to vote for
'it!


Dim char As IAgentCtlCharacterEx 'Declare the String char as the Character file
Dim Anim As String 'Declare the Anim string which we will use later on (declaring this will make it easy for us to change the character with ease, later on)

Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
char.Stop 'This makes the Character stop what he is doing
char.Play "Restpose" 'This makes the character return to his restpose

If Button = vbRightButton Then
frmMain.PopupMenu mnuMain
End If
End Sub

Private Sub cmdGreet_Click()
char.Play "Greet"

End Sub

Private Sub cmdMove_Click()
char.MoveTo 300, 300
End Sub

Private Sub cmdSpeak_Click()
char.Speak "Hello"
End Sub

Private Sub Form_Load()

Anim = "Peedy"    'We set the Anim String to "Peedy" . You can set this to Genie, or Merlin, or Robby too.

Agent1.Characters.Load Anim, Anim & ".acs"    'This is how we tell VB where to find the character's acs file. VB by default looks in the C:\Windows\MsAgent\Chars\ folder for the character file

Set char = Agent1.Characters(Anim)       'Remember we declared the char string earlier? Now we set char to equal Agent1.Charachters property. Note that the because we used the Anim string we can now change the character by changing only one line of code.

char.AutoPopupMenu = False





char.Show 'Shows the Character File (If set to "Peedy" he comes flying out of the background)




End Sub

Private Sub Text1_Change()

End Sub

Private Sub mnuRead_Click()
char.Play "Read"

End Sub

Private Sub mnuWrite_Click()
char.Play "Write" 'This makes the Character write!
End Sub
