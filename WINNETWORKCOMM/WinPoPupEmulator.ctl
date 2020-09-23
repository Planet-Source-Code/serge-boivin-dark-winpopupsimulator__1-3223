VERSION 5.00
Begin VB.UserControl WinPoPupEmulator 
   BackColor       =   &H000000FF&
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   Picture         =   "WinPoPupEmulator.ctx":0000
   ScaleHeight     =   705
   ScaleWidth      =   1515
   Begin VB.Timer TimerTestComm 
      Enabled         =   0   'False
      Interval        =   65000
      Left            =   360
      Tag             =   "0"
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1320
      Top             =   0
   End
   Begin VB.Timer TimerLedReceive 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2520
      Top             =   480
   End
   Begin VB.Timer TimerLedSend 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2520
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   4560
      Picture         =   "WinPoPupEmulator.ctx":1382
      Top             =   4080
      Width           =   1110
   End
   Begin VB.Shape LedReceive 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape LedSend 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   75
      Width           =   255
   End
   Begin VB.Label LblCaption 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1105
   End
End
Attribute VB_Name = "WinPoPupEmulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1
Private Type ChaineConfirmation
  ChaineToSend As String
  Flags As Boolean
  NewNow As Date
  SendTo As String
  SendBy As String
End Type
  
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type
Const MAILSLOT_WAIT_FOREVER = (-1)
Const OPEN_EXISTING = 3
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Const GENERIC_EXECUTE = &H20000000
Const GENERIC_ALL = &H10000000
Const INVALID_HANDLE_VALUE = -1
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hHandle As Long) As Long
Private Declare Function WriteFile Lib "Kernel32" (ByVal hFileName As Long, ByVal lpBuff As Any, ByVal nNrBytesToWrite As Long, lpNrOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwAccess As Long, ByVal dwShare As Long, ByVal lpSecurityAttrib As Long, ByVal dwCreationDisp As Long, ByVal dwAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetMailslotInfo Lib "Kernel32" (ByVal hMailslot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Private Declare Function ReadFile Lib "Kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateMailslot Lib "kernel32.dll" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Dim MHandle As Long
'Default Property Values:
Dim m_def_NoVersion As String  'Ici je triche ma valeur par default
Const m_def_IsLog = True
Const m_def_WorkGroup = "VERTEX"
Const m_def_IsOn = True
Const m_def_OnlyLightFunction = True
Const m_def_Flashing = 20
Const m_def_MailSlotEnvoie = "messngr"
Const m_def_MailSlotName = "messngr"
Const m_def_MailSlotHandle = 0
Const m_def_MessageFrom = ""
Const m_def_MessageText = ""

'Property Variables:
Dim m_NoVersion As String
Dim m_IsLog As Boolean
Dim m_WorkGroup As String
Dim m_IsOn As Boolean
Dim m_OnlyLightFunction As Boolean
Dim m_Flashing As Integer
Dim m_NbrMessage As Integer
Dim m_MailSlotEnvoie As String
Dim m_MailSlotName As String
Dim m_MailSlotHandle As Long
Dim m_MessageFrom() As String
Dim m_MessageText() As String
Dim m_PathToLog As String
Dim WorkGroupName As String
'Event Declarations:
Event MessageWaiting(NbrMessageWaiting As Integer)
Function SendToWinPopUp(PopFrom As String, MsgText As String, Optional ByVal NoteMe As Boolean = True) As Long
    Dim rc As Long
    Dim mshandle As Long
    Dim msgtxt As String
    Dim byteswritten As Long
    Dim MailSlotName As String
    Dim Pos As Long
    Dim PopTo As String
    ' name of the mailslot
    PopTo = WorkGroupName
    If Not m_IsOn Then Exit Function
    If ExtractChaineEntreSlash(MsgText, 1, False) <> "CONFIRMATION" And NoteMe Then
      Pos = NoteForConfirmation(MsgText, MailSlotEnvoie, MailSlotName)
      MsgText = MsgText & Pos & "/"
    End If
    LedSend.FillColor = vbBlue
    LedSend.Refresh
    MailSlotName = "\\" + PopTo + "\mailslot\" & MailSlotEnvoie
    msgtxt = PopFrom + Chr(0) + PopTo + Chr(0) + MsgText + Chr(0)
    mshandle = CreateFile(MailSlotName, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, -1)
    rc = WriteFile(mshandle, msgtxt, Len(msgtxt), byteswritten, 0)
    rc = CloseHandle(mshandle)
    TimerLedSend.Enabled = True
    LogMe MailSlotEnvoie, MsgText, True
    
End Function
Private Sub ReadToWinPoPup()
Dim NextSize As Long
Dim Waiting As Long
Dim Buffer() As Byte
Dim ReadSize As Long
Dim FHandle As Long
Dim TempsWaiting As Long
Dim tempo As String
Dim x As Integer
Dim passe As Integer
Dim FromTo As String
Dim Message As String
Dim R As Long
  If Not m_IsOn Then Exit Sub
 TempsWaiting = MAILSLOT_WAIT_FOREVER
 ' look for message
 FHandle = GetMailslotInfo(MHandle, R, NextSize, Waiting, TempsWaiting)
 'if message go read this
 If Waiting <> 0 Then
   LedReceive.FillColor = vbYellow
   LedReceive.Refresh
   TimerLedReceive.Enabled = True
   Dim y As Integer
   For y = 1 To Waiting
     If m_MessageFrom(UBound(m_MessageFrom)) <> "" Then
       ReDim Preserve m_MessageFrom(UBound(m_MessageFrom) + 1)
       ReDim Preserve m_MessageText(UBound(m_MessageText) + 1)
     End If
     ReDim Buffer(NextSize)
     FHandle = ReadFile(MHandle, Buffer(1), NextSize, ReadSize, ByVal 0&)
     passe = 1
     For x = 1 To ReadSize
       If Buffer(x) <> 0 Then
         tempo = tempo & Chr(Buffer(x))
       Else
         Select Case passe
           Case 1
             m_MessageFrom(UBound(m_MessageFrom)) = tempo
             passe = 2
           Case 2
             passe = 3
           Case 3
             m_MessageText(UBound(m_MessageText)) = tempo
         End Select
         tempo = ""
       End If
     Next
     If m_MessageText(UBound(m_MessageText)) <> "" Then
      LogMe m_MessageFrom(UBound(m_MessageFrom)), m_MessageText(UBound(m_MessageText)), False
      If ExtractChaineEntreSlash(m_MessageText(UBound(m_MessageText)), 1, False) = "CONFIRMATION" Then
        ConfimationReceive m_MessageText(UBound(m_MessageText))
        ClearMessage UBound(m_MessageText)
      ElseIf ExtractChaineEntreSlash(m_MessageText(UBound(m_MessageText)), 1, False) = "TESTCOMM" Then
        TimerTestComm.Tag = 0
        ClearMessage UBound(m_MessageText)
        m_NbrMessage = (UBound(m_MessageFrom))
        Exit Sub
      Else
        SendConfimation m_MessageFrom(UBound(m_MessageFrom)), m_MessageText(UBound(m_MessageText))
      End If
    End If
   Next
   If Not (UBound(m_MessageFrom) = 1 And m_MessageFrom(1) = "") Then
     m_NbrMessage = (UBound(m_MessageFrom))
     RaiseEvent MessageWaiting(UBound(m_MessageFrom))
   End If
 End If
End Sub
Private Sub ConfimationReceive(ByVal Chaine As String)
Dim NoConfirmation As Long
Dim Ch As ChaineConfirmation
Dim NoFile As Integer
  'Prendre Le Numero De Confirmation ... Qui Est Toujours Le Troisieme Parametre
  NoConfirmation = ExtractChaineEntreSlash(Chaine, 3, False)
  NoFile = FreeFile
  Open m_PathToLog & MailSlotName & ".txt" For Random As #NoFile Len = 600
  Get NoFile, NoConfirmation, Ch
  'Ici On Le Flag Dans Le Fichier Comme Etant Confirmer
  Ch.Flags = False
  Put NoFile, NoConfirmation, Ch
  Close #NoFile
End Sub
Private Sub SendConfimation(ByVal From As String, ByVal Message As String)
  MailSlotEnvoie = From
  SendToWinPopUp MailSlotName, "CONFIRMATION/" & ExtractChaineEntreSlash(Message, 1, False) & "/" & ExtractChaineEntreSlash(Message, GetPosConfirmationNumber(Message), False) & "/", False
End Sub
Private Sub Timer1_Timer()
    ReadToWinPoPup
End Sub
Public Sub Initialisation()
Dim MaxMessage As Long
Dim MesssageTimer As Long
Dim t As SECURITY_ATTRIBUTES
    If Not IsOn Then
      LedReceive.Visible = False
      LedSend.Visible = False
      Exit Sub
    End If
    WorkGroupName = FindWorkGroupName(SV_TYPE_DOMAIN_ENUM, WorkGroup)
    t.nLength = Len(t)
    t.bInheritHandle = False
    MaxMessage = 0
    MesssageTimer = MAILSLOT_WAIT_FOREVER
    MHandle = CreateMailslot("\\.\mailslot\" & m_MailSlotName, MaxMessage, MesssageTimer, t)
    ReDim m_MessageFrom(1)
    ReDim m_MessageText(1)
    m_MessageFrom(1) = m_def_MessageFrom
    m_MessageText(1) = m_def_MessageText
    m_MailSlotHandle = MHandle
    If MHandle = -1 Then
      LedReceive.FillColor = vbBlack
      LedSend.FillColor = vbBlack
    Else
      Timer1.Enabled = True
      Timer2.Enabled = True
    End If
    TimerTestComm.Enabled = True
End Sub

Public Property Get MessageFrom(Index As Integer) As String
    MessageFrom = m_MessageFrom(Index)
End Property
Public Property Get MessageText(Index As Integer) As String
    MessageText = m_MessageText(Index)
End Property
Public Sub ClearMessage(Index As Integer)
Dim tempo() As String
Dim tempo2() As String
Dim x As Integer
Dim y As Integer
Dim z As Integer
    If UBound(m_MessageFrom) = 1 Then
        m_MessageFrom(1) = ""
        m_MessageText(1) = ""
        m_NbrMessage = 1
        Exit Sub
    End If
    ReDim tempo(UBound(m_MessageFrom) - 1)
    ReDim tempo2(UBound(m_MessageFrom) - 1)
    y = 1
    z = 1
    For x = 1 To UBound(m_MessageFrom)
        If x <> Index Then
            tempo(z) = m_MessageFrom(y)
            tempo2(z) = m_MessageText(y)
            z = z + 1
        End If
        y = y + 1
    Next
    ReDim m_MessageFrom(UBound(tempo))
    ReDim m_MessageText(UBound(tempo2))
    For x = 1 To UBound(tempo)
        m_MessageFrom(x) = tempo(x)
        m_MessageText(x) = tempo2(x)
    Next
    m_NbrMessage = UBound(tempo)
End Sub
Public Sub CloseSimulator()
     Timer1.Enabled = False
     CloseHandle MHandle
     TimerTestComm.Enabled = False
End Sub

Public Property Get MailSlotHandle() As Long
    MailSlotHandle = m_MailSlotHandle
End Property
Public Property Get NbrMessage() As Integer
    NbrMessage = m_NbrMessage
End Property
  

Public Property Get MailSlotName() As String
    MailSlotName = m_MailSlotName
End Property

Public Property Let MailSlotName(ByVal New_MailSlotName As String)
    m_MailSlotName = New_MailSlotName
    PropertyChanged "MailSlotName"
End Property
Public Property Let PathToLog(ByVal New_PathToLog As String)
  m_PathToLog = New_PathToLog
  If Dir(m_PathToLog, vbDirectory) = "" Then
    MkDir m_PathToLog
  End If
  m_PathToLog = m_PathToLog
End Property

Private Sub Timer2_Timer()
Dim NoFile As Integer
Dim Alarme As String
Dim NoRec As Integer
Dim UneTrouver As Boolean
Dim Temp As String
Dim Ch As ChaineConfirmation
Dim CptSend As Integer
  'Recherche Du Fichier Ou Sont Stoker Les Donnees Deja Envoyer
  If Dir(m_PathToLog & MailSlotName & ".txt") <> "" Then
    NoFile = FreeFile
    Open m_PathToLog & MailSlotName & ".txt" For Random As #NoFile Len = 600
    'Verifier Si Le Fichier Contient Reelement Des Donnees
    If EOF(NoFile) Then
      'Si Aucunne Donnee ... Qu'est-ce Qui Fou La Lui ... On Le Flush
      Close #NoFile
      Kill m_PathToLog & MailSlotName & ".txt"
      Exit Sub
    End If
    NoRec = 0
    UneTrouver = False
    'On a Des Donnee Dans Le Fichier Alors On Va Les Lire Hein !
    CptSend = 0
    Do While Not EOF(NoFile)
      NoRec = NoRec + 1
      Get #NoFile, NoRec, Ch
      'Est-ce Qu'il A Ete Confimer ???
      If Ch.Flags Then
        CptSend = CptSend + 1
        'Ca Ben L'Air Que Non ... Alors Est-Ce Que Le Temps Est Venu De Le Renvoyer ??
        If Abs(DateDiff("s", Ch.NewNow, Now)) > 30 Then
          MailSlotEnvoie = Ch.SendTo
          SendToWinPopUp MailSlotName, Ch.ChaineToSend & NoRec & "/", False
          'Bon Ben Je Les Renvoyer ... Je Note La NOW Du Nouvelle Envoie Au Cas Ou
          'L'Autre Serais Dure De La Feuille Encore Une Fois ...
          Ch.NewNow = Now
          Put #NoFile, NoRec, Ch
        End If
        UneTrouver = True
      End If
      If CptSend >= 5 Then
        Exit Do
      End If
    Loop
    Close #NoFile
    'Y'Avais-Tu  Encore Des Donnee Non Confirmee ??
    If Not UneTrouver Then
      'Ben Non .. Donc On Flush Le Fichier On A N'a Plus De Besoin ... Bon Debarra
      Kill m_PathToLog & MailSlotName & ".txt"
    End If
  End If
  
  
End Sub

Private Sub TimerLedReceive_Timer()
  LedReceive.FillColor = vbGreen
  LedReceive.Refresh
  TimerLedReceive.Enabled = False
End Sub

Private Sub TimerLedSend_Timer()
  LedSend.FillColor = vbGreen
  LedSend.Refresh
  TimerLedSend.Enabled = False
End Sub

Private Sub TimerTestComm_Timer()
Dim Temp As String
  If TimerTestComm.Tag = 0 Then
    Temp = m_MailSlotEnvoie
    m_MailSlotEnvoie = m_MailSlotName
    SendToWinPopUp m_MailSlotName, "TESTCOMM/", False
    TimerTestComm.Tag = 1
    m_MailSlotEnvoie = Temp
  ElseIf TimerTestComm.Tag = 1 Then
    'Pas De Reponse ...
     CloseHandle MHandle
     Initialisation
     TimerTestComm.Tag = 0
  End If
End Sub

Private Sub UserControl_Initialize()
  If Dir("C:\Log\", vbDirectory) = "" Then
    MkDir "C:\Log\"
  End If
  m_PathToLog = "C:\Log\"
  m_def_NoVersion = App.Major & "." & App.Revision
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_MailSlotName = m_def_MailSlotName
  m_MailSlotEnvoie = m_def_MailSlotEnvoie
  m_Flashing = m_def_Flashing
  
  m_OnlyLightFunction = m_def_OnlyLightFunction
  m_IsOn = m_def_IsOn
  m_WorkGroup = m_def_WorkGroup
  m_IsLog = m_def_IsLog
  m_NoVersion = m_def_NoVersion
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_MailSlotName = PropBag.ReadProperty("MailSlotName", m_def_MailSlotName)
    m_MailSlotEnvoie = PropBag.ReadProperty("MailSlotEnvoie", m_def_MailSlotEnvoie)
  LblCaption.Caption = PropBag.ReadProperty("Caption", "")
  m_Flashing = PropBag.ReadProperty("Flashing", m_def_Flashing)
  m_OnlyLightFunction = PropBag.ReadProperty("OnlyLightFunction", m_def_OnlyLightFunction)
  m_IsOn = PropBag.ReadProperty("IsOn", m_def_IsOn)
  m_WorkGroup = PropBag.ReadProperty("WorkGroup", m_def_WorkGroup)
  m_IsLog = PropBag.ReadProperty("IsLog", m_def_IsLog)
  m_NoVersion = PropBag.ReadProperty("NoVersion", m_def_NoVersion)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("MailSlotName", m_MailSlotName, m_def_MailSlotName)
    Call PropBag.WriteProperty("MailSlotEnvoie", m_MailSlotEnvoie, m_def_MailSlotEnvoie)
  Call PropBag.WriteProperty("Caption", LblCaption.Caption, "")
  Call PropBag.WriteProperty("Flashing", m_Flashing, m_def_Flashing)
  Call PropBag.WriteProperty("OnlyLightFunction", m_OnlyLightFunction, m_def_OnlyLightFunction)
  Call PropBag.WriteProperty("IsOn", m_IsOn, m_def_IsOn)
  Call PropBag.WriteProperty("WorkGroup", m_WorkGroup, m_def_WorkGroup)
  Call PropBag.WriteProperty("IsLog", m_IsLog, m_def_IsLog)
  Call PropBag.WriteProperty("NoVersion", m_NoVersion, m_def_NoVersion)
End Sub

Public Property Get MailSlotEnvoie() As String
    MailSlotEnvoie = m_MailSlotEnvoie
End Property

Public Property Let MailSlotEnvoie(ByVal New_MailSlotEnvoie As String)
    m_MailSlotEnvoie = New_MailSlotEnvoie
    PropertyChanged "MailSlotEnvoie"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LblCaption,LblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
  Caption = LblCaption.Caption
End Property


Public Property Let Caption(ByVal New_Caption As String)
  LblCaption.Caption() = New_Caption
  PropertyChanged "Caption"
End Property

Public Property Get Flashing() As Integer
  Flashing = m_Flashing
End Property

Public Property Let Flashing(ByVal New_Flashing As Integer)
  m_Flashing = New_Flashing
  PropertyChanged "Flashing"
  TimerLedReceive.Interval = New_Flashing
  TimerLedSend.Interval = New_Flashing
End Property


Private Sub LogMe(ByVal SendToFrom As String, ByVal Chaine As String, ByVal IsSend As Boolean)
Dim NoFile As Integer
  If IsLog Then
    NoFile = FreeFile
    Open m_PathToLog & MailSlotName & ".log" For Append As #NoFile
    If IsSend Then
      Print #NoFile, "Envoyer A:" & SendToFrom & "  " & Chaine
    Else
      Print #NoFile, "Recu De  :" & SendToFrom & "  " & Chaine
    End If
    Close #NoFile
  End If
End Sub
'
Public Property Get OnlyLightFunction() As Boolean
  OnlyLightFunction = m_OnlyLightFunction
End Property

Public Property Let OnlyLightFunction(ByVal New_OnlyLightFunction As Boolean)
  m_OnlyLightFunction = New_OnlyLightFunction
  PropertyChanged "OnlyLightFunction"
  
  If New_OnlyLightFunction Then
    LedReceive.Left = 0
    LedSend.Left = 0
    UserControl.Width = 270
    UserControl.Picture = LoadPicture()
    UserControl.BackStyle = 0
    LblCaption.Visible = False
  Else
    LedReceive.Left = 1200
    LedSend.Left = 1200
    UserControl.Width = 1530
    LblCaption.Visible = True
    UserControl.BackStyle = 1
    UserControl.Picture = Image1.Picture
  End If
End Property
Private Function GetPosConfirmationNumber(ByVal Chaine As String) As Long
Dim x As Integer
'Ici On Veut Compter Le Nombre De "/", Pour Savoir Ou Se Trouve Notre Numero De
'Confirmation Qui est Toujours Le Dernier ....
  GetPosConfirmationNumber = 0
  For x = 1 To Len(Chaine)
    If Mid(Chaine, x, 1) = "/" Then
      GetPosConfirmationNumber = GetPosConfirmationNumber + 1
    End If
  Next
End Function

Private Function NoteForConfirmation(ByVal StringToNote As String, ByVal SendTo As String, ByVal SendBy As String) As Long
Dim NoFile As Integer
Dim Ch As ChaineConfirmation
Dim Pos As Integer
'Ici on fais ... a Pis aller Voir La Fonction Envoie .. C'est Elle Qui L'Appel
'Et Je Viens Tout Juste De Decrire L'Appel De Cette Fonction
  NoFile = FreeFile
  Open m_PathToLog & MailSlotName & ".txt" For Random As #NoFile Len = 600
  Ch.ChaineToSend = StringToNote
  Ch.Flags = True
  Ch.NewNow = Now
  Ch.SendTo = SendTo
  Ch.SendBy = SendBy
  Pos = GetFirstNotFlagPosition()
  Put #1, Pos, Ch
  Close #NoFile
  NoteForConfirmation = Pos
End Function

'***************************************************************************
'*Cette Fonction Trouve Le Premier Record Qui N'Est Plus Flager            *
'***************************************************************************
Private Function GetFirstNotFlagPosition() As Long
Dim NoFile As Integer
Dim Ch As ChaineConfirmation
Dim Pos As Long
  NoFile = FreeFile
  Open m_PathToLog & MailSlotName & ".txt" For Random As #NoFile Len = 600
  Pos = 1
  Do While Not EOF(NoFile)
    'Recherche De La Premiere Position Deja Confirmer Pour La Remplacer Par La Nouvelle
    'Si Pas Trouver On Apprend Au Fichier Tse Pas Plus Fou Qu'Un Autre
    Get NoFile, Pos, Ch
    If Not Ch.Flags Then
      Exit Do
    End If
    Pos = Pos + 1
  Loop
  Close #NoFile
  GetFirstNotFlagPosition = Pos
End Function
Public Function ExtractChaineEntreSlash(ByVal Chaine As String, Position As Integer, AllReste As Boolean) As String
'chaine= chaine de caractère
'Position = numero du slash  -1 a rechercher pour debut d'extraction
'allReste = Si vrai prend tout ce qui est apres la position
'Function qui extrait un parti de la chaine de caractère en commencant a la position
'du slash -1 demander
'Explication du -1 : c'est que la chaine de commence pas par un slash donc si je donne
'2 comme position il va me donné le deuxième mots et celui ci est apres le premier
'le premier slash
'Example :
'Debug.print ExtractChaineEntreSlash("allo/toi/ça/va",3,false)
'le resultat est = ça
Dim Pos As Integer
Dim x As Integer
  For x = 1 To Position - 1
    Pos = InStr(1, Chaine, "/", vbTextCompare)
    Chaine = Mid(Chaine, Pos + 1)
  Next
  Pos = InStr(1, Chaine, "/", vbTextCompare)
  If AllReste Then
    ExtractChaineEntreSlash = Chaine
  Else
    ExtractChaineEntreSlash = Mid(Chaine, 1, Pos - 1)
  End If
End Function
Public Property Get IsOn() As Boolean
  IsOn = m_IsOn
End Property

Public Property Let IsOn(ByVal New_IsOn As Boolean)
  m_IsOn = New_IsOn
  PropertyChanged "IsOn"
End Property

Public Property Get WhatOs() As String
  WhatOs = WinOs
End Property
Public Property Let WhatOs(ByVal New_WhatOs As String)
  WinOs = New_WhatOs
End Property



Public Property Get WorkGroup() As String
  WorkGroup = m_WorkGroup
End Property

Public Property Let WorkGroup(ByVal New_WorkGroup As String)
  m_WorkGroup = New_WorkGroup
  PropertyChanged "WorkGroup"
End Property

Public Property Get IsLog() As Boolean
  IsLog = m_IsLog
End Property

Public Property Let IsLog(ByVal New_IsLog As Boolean)
  m_IsLog = New_IsLog
  PropertyChanged "IsLog"
End Property

Public Property Get NoVersion() As String
  NoVersion = m_NoVersion
End Property

Public Property Let NoVersion(ByVal New_NoVersion As String)
  If Ambient.UserMode = False Then Err.Raise 394
  If Ambient.UserMode Then Err.Raise 393
  m_NoVersion = New_NoVersion
  PropertyChanged "NoVersion"
End Property

