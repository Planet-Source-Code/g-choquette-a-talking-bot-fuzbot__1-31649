VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form frmFuzbot 
   Caption         =   "Fuzbot"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmFuzbot.frx":0000
      Top             =   360
      Width           =   4575
   End
   Begin HTTSLibCtl.TextToSpeech TextToSpeech1 
      Height          =   1455
      Left            =   1440
      OleObjectBlob   =   "frmFuzbot.frx":009E
      TabIndex        =   0
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   3000
      Shape           =   2  'Oval
      Top             =   3360
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C96C34&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2880
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   2520
      Shape           =   2  'Oval
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   1560
      Shape           =   2  'Oval
      Top             =   3360
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C96C34&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   1080
      Shape           =   2  'Oval
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmFuzbot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
  SpeakTextbox
End Sub

Private Sub Form_Load()
  TextToSpeech1.Interfaces 2
  TextToSpeech1.MouthHeight = 0
  SpeakTextbox
End Sub

Private Sub Text1_DblClick()
  SpeakTextbox
End Sub

Private Sub TextToSpeech1_ClickIn(ByVal x As Long, ByVal y As Long)
  SpeakTextbox
End Sub

Private Sub TextToSpeech1_SpeakingDone()
  TextToSpeech1.MouthHeight = 0
End Sub

Private Sub SpeakTextbox()
  With TextToSpeech1
    If Text1.Text <> "" Then
      .Speak Text1.Text
    End If
  End With
End Sub
