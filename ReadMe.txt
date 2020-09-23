'Start Ocx 
	'Parameter For Just See Light In RunTime
	  FrmStatut.LO_CELCA.OnlyLightFunction = True
	'Parameter For Open MailSlot ( It is false, And You Use SendToWinPopop Methode,Nothing happend)
	  FrmStatut.LO_CELCA.IsOn = readinifile(FrmStatut.LO_CELCA.MailSlotName, "IsOn", "True", "Celect Card 2000.ini")
	'Parameter For Loggin Communication
	  FrmStatut.LO_CELCA.IsLog = readinifile(FrmStatut.LO_CELCA.MailSlotName, "IsLog", "True", "Celect Card 2000.ini")
	;Prameter  For If Win98, WinOs=98 Else WinOs=2000
	  FrmStatut.LO_CELCA.WhatOs = WinOs
	;For Some Reason, Handle Not Always Close , Then Read Previous Handle From IniFile
	  PrecivousHandle = readinifile(FrmStatut.LO_CELCA.MailSlotName, "Handle", "-1", "Celect Card 2000.ini")
	;Need This Line OnLy For Win98
	  FrmStatut.LO_CELCA.WorkGroup = readinifile(FrmStatut.LO_CELCA.MailSlotName, "WorkGroup", "Vertex", "Celect Card 2000.ini")
	;Close Previous Handle
	  CloseHandle PrecivousHandle
	;Path For Log File
	  FrmStatut.LO_CELCA.PathToLog = "C:\LOG\"
	;Start Ocx
	  FrmStatut.LO_CELCA.Initialisation
	;Check If Mailslot Work 
	  If FrmStatut.LO_CELCA.MailSlotHandle = -1 Then
	    ErrorTrap "FrmMain!InitialiseMessagerie", "ERREUR LORS DE L'OUVERTURE DE LA BOITE D'ENVOIE EN MODE RESEAU"
	  End If
	;Keep Handle In IniFile
	  Tempo = Trim(Str(FrmStatut.LO_CELCA.MailSlotHandle))
	  WritePrivateProfileString FrmStatut.LO_CELCA.MailSlotName, "Handle", Tempo, "Celect Card 2000.ini"
;Send Method Example
	Public Sub Send(ByVal Chaine As String, ByVal SendTo As String, ByVal MailSlot As Object)
	  MailSlot.MailSlotEnvoie = SendTo
	  MailSlot.SendToWinPopUp MailSlot.MailSlotName, Chaine
	End Sub
;Send Call Example
	Send "It's A Test", MailSlotDestination, MailSlotSenderObject
;Message Receive Event Sample
	Private Sub LO_CELCA_MessageWaiting(NbrMessageWaiting As Integer)
	  For X = 1 To NbrMessageWaiting       
	    'True Message ?? Not A Bad Events ?? 
	    If LO_CELCA.MessageText(y) <> "" Then
	      'Yep Then .. extract Some Parameter
	       'Message From ?? 
	       FromLogiciel = LO_CELCA.MessageFrom(y)
	       Message=LO_CELCA.MessageText(y)
	       'Enter Your Code Here 
	    End If
	    'Need Keep This Message In Memory ?? 
	    If Not (Commande = "REPONSEPIGE" And AreInPige) Then
	      'Nah .. Then Clear It
	      LO_CELCA.ClearMessage (y)
	    Else 
	      'Ok Ok .. take Next Message, Don't Clear It	
	      y = y + 1
	    End If
	  Next
	End Sub
;Note .... Ocx Send Message To Him Each 65 second, if he don't receive here message after 65
second, ocx close handle and open a new handle


