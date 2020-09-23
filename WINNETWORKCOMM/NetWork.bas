Attribute VB_Name = "NetWork"
Option Explicit
Public Const SV_TYPE_DOMAIN_ENUM = &H80000000
Private Type SERVER_INFO_101
    dw_platform_id As Long
    ptr_name As Long
    dw_ver_major As Long
    dw_ver_minor As Long
    dw_type As Long
    ptr_comment As Long
End Type
Global WinOs As String
Private Declare Function NetServerEnum Lib "Netapi32.dll" (vServername As Any, ByVal lLevel As Long, vBufptr As Any, lPrefmaxlen As Long, lEntriesRead As Long, lTotalEntries As Long, vServerType As Any, ByVal sDomain As String, vResumeHandle As Any) As Long
Private Declare Sub RtlMoveMemory Lib "Kernel32" (dest As Any, vSrc As Any, ByVal lSize&)
Private Declare Sub lstrcpyW Lib "Kernel32" (vDest As Any, ByVal sSrc As Any)
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
Public Function FindWorkGroupName(ByVal lType As Long, ByVal WorkGroupDefault As String) As String
        
Dim Server_Info As Long
Dim lEntries As Long
Dim lTotal As Long
Dim lMax As Long
Dim vResume As Variant
Dim tServer_info_101 As SERVER_INFO_101
Dim sServer As String
Dim sDomain As String
Dim lServerInfo101StructPtr As Long
Dim i As Long
Dim bBuffer(512) As Byte
Dim lReturn As Long
'NetServerEnum N'Existe Pas En win 98
'Donc Dans Le Cas D'Un Win98 Ben Je Mets Le Nom De Groupe a Vertex
  If WinOs <> "98" Then
    lReturn = NetServerEnum(ByVal 0&, 101, Server_Info, lMax, lEntries, lTotal, ByVal lType, sDomain, vResume)
    If lReturn <> 0 Then
        'Erreur
       ' Kappel_ErrorHandler "NetWork.bas!Public Function FindWorkGroupName(lType As Long) As String", lReturn
        Exit Function
    End If
    lServerInfo101StructPtr = Server_Info
    'DoEvents
    RtlMoveMemory tServer_info_101, ByVal lServerInfo101StructPtr, Len(tServer_info_101)
    lstrcpyW bBuffer(0), tServer_info_101.ptr_name
     
    i = 0
    Do While bBuffer(i) <> 0
      sServer = sServer & Chr$(bBuffer(i))
      i = i + 2
      'DoEvents
    Loop
    FindWorkGroupName = sServer
    NetApiBufferFree lServerInfo101StructPtr
  Else
    FindWorkGroupName = WorkGroupDefault
  End If
End Function
