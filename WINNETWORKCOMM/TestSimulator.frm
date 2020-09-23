VERSION 5.00
Object = "*\AWindowsNetworkCommunication.vbp"
Begin VB.Form TestSimulator 
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin WinNetworkComm.WinPoPupEmulator WinPoPupEmulator1 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
      _extentx        =   2778
      _extenty        =   1296
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   3720
      TabIndex        =   0
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox lABEL1 
      Height          =   1455
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3600
      Width           =   2055
   End
End
Attribute VB_Name = "TestSimulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cpt As Integer

Private Sub Command1_Click()
    WinPoPupEmulator1.SendToWinPopUp WinPoPupEmulator1.MailSlotName, "Vertex", "test" & cpt
    cpt = cpt + 1
End Sub

Private Sub Command2_Click()
    WinPoPupEmulator1.SendToWinPopUp "serge", "VERTEX03", "test" & cpt
    cpt = cpt + 1
    WinPoPupEmulator1.SendToWinPopUp "serge", "VERTEXX", "test" & cpt
    cpt = cpt + 1
End Sub

Private Sub Form_Load()
    cpt = 1
    WinPoPupEmulator1.MailSlotName = "BBB"
    WinPoPupEmulator1.MailSlotEnvoie = "AAA"
    WinPoPupEmulator1.Initialisation
    Debug.Print WinPoPupEmulator1.MailSlotHandle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WinPoPupEmulator1.CloseSimulator
End Sub

Private Sub WinPoPupEmulator1_MessageWaiting(NbrMessageWaiting As Integer)
Dim x As Integer
For x = 1 To NbrMessageWaiting
    lABEL1.Text = lABEL1.Text & vbCrLf & "Message From " & WinPoPupEmulator1.MessageFrom(1) & vbCrLf & WinPoPupEmulator1.MessageText(1)
    WinPoPupEmulator1.ClearMessage (1)
Next
End Sub
