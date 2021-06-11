Attribute VB_Name = "modGlobal"
Option Explicit

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal SounName As String, ByVal uFlags As Long) As Long
Public Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal HostName As String) As Long
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal Handle As Long) As Boolean
Public Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean

Public Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function OpenProcess Lib "KERNEL32" (ByVal dwAccess As Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
Public Const Infini = -1&

Global Const AppSite = "Template sarl"

Public Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Public Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Public Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type
Public Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type

Type TypeConnexions
    strPoste     As String
    strPseudo    As String
    strIP        As String
    iNoControl   As Integer
End Type
Type TypeTexte
    lCouleur    As Long
    strFonte    As String
    dTaille     As Double
    bBold       As Boolean
    bItalic     As Boolean
End Type

Public Connexions(200)  As TypeConnexions
Public AffHeure         As TypeTexte
Public AffPseudo        As TypeTexte
Public AffSystem        As TypeTexte
Public AffPerso         As TypeTexte
Public AffAutres        As TypeTexte

Public strNomMachine    As String
Public strNomUser       As String
Public AppAdresse       As String
Public iCompteMinutes   As Integer
Public iPeriodeRafr     As Integer
Public Sys              As Object
Public LogFile          As String
Public OkLogFile        As Boolean
Public iPurgeLog        As Integer
Public bDebug           As Boolean
Public sTimeDebut       As Single
Public bNouveauMessage  As Boolean
Public bNouveauMessageVu As Boolean

Public bCarteSon        As Boolean
Public strSonMessage    As String
Public iAvecSonMessage  As Integer
Public strSonArriv�e    As String
Public iAvecSonArriv�e  As Integer

Public sTimeTouche      As Single
Public bTimeTouche      As Boolean

Public strIPextraNet    As String


Public Sub Main()
    
    ' D�tecte si on demande le mode "debug" dans la ligne de commande
    bDebug = False
    If LCase(Element(Command(), 1, " ")) = "debug" Then bDebug = True
    
    AppAdresse = "13 rue Lecarnier" & vbCr & _
                 "76700 HARFLEUR" & vbCr & _
                 "Contact : Jacques Millet : 02.35.47.74.22"
    
    Forme.Show

End Sub

Public Function Element(ByVal Texte As String, _
                        ByVal Numero As Integer, _
                        ByVal S�parateur As String) As String
    
    '--- Cette fonction renvoie le texte equivalent au Xeme element (Numero de 1 a X)
    '    de la chaine Texte. Chaque element etant s�par� par des S�parateurs

    Dim Debut As Integer, r As Integer, No As Integer
    
    If Right(Texte, Len(S�parateur)) <> S�parateur Then Texte = Texte & S�parateur
    
    Debut = 1
    No = 1

Element_0:
    r = InStr(Debut, Texte, S�parateur)
    If r = 0 Then GoTo Element_Fin
    If Numero = No Then GoTo Element_10
    No = No + 1
    Debut = r + Len(S�parateur)
    If r >= Len(Texte) Then GoTo Element_Fin
    DoEvents
    GoTo Element_0
    
Element_10:
    Element = Mid(Texte, Debut, r - Debut)
    
Element_Fin:
    
End Function

Public Sub Ecrit_Log(ByVal Texte As String, Optional ByVal Erreur As Boolean = False)

    ' Cette sub assure l'ecriture du Texte dans le fichier de suivi.
    ' Ces lignes sont horodat�es.
    ' Leur format differe s'il s'agit d'une info (Erreur=False) ou d'une erreur (Erreur=True)
    '   31/12/99 13:04:15  ERREUR : texte de l'erreur
    '   31/12/99 13:04:15           texte du message

    Dim fflog As Integer
    Dim Temp As String, r As Integer
    
    ' Au d�marrage, test liaison avec fichier Log
    If Texte = "TEST" Then
        OkLogFile = Sys.FileExists(LogFile)
        Exit Sub
    End If
    
    ' Si l'option est "sans" fichier log MAIS que les donn�es � �crire
    ' sont de type SYSt�me, on �crit quand m�me
    If Left(Texte, 3) = "SYS" Then
        Texte = Right(Texte, Len(Texte) - 3)
        GoTo Suite
    End If
    
    ' Sinon, on n'�crit rien si l'option est "sans" fichier log
    If Not Forme.mnuLogFile.Checked Then Exit Sub
    
Suite:
    fflog = FreeFile
    On Error GoTo Log_Erreur
    Open LogFile For Append Access Write As #fflog
    
    ' 1�re ligne : Date et Heure + Mot "Erreur" + Texte
    Temp = Format(Now, "dd/mm/yy Hh:Nn:Ss  ")
    If Erreur = True Then
        Print #fflog, Temp & "ERREUR : " & Element(Texte, 1, vbCr)
    Else
        Print #fflog, Temp & "         " & Element(Texte, 1, vbCr)
    End If
    ' Si le texte comporte plusieurs lignes, on �crit les suivantes avec un d�calage
    ' pour se retrouver � droite de la date
    r = 2
boucle_10:
    Temp = Element(Texte, r, vbCr)
    If Temp <> "" And Temp <> vbCr Then
        Print #fflog, "                            " & Temp
        r = r + 1
        GoTo boucle_10
    End If
    
    OkLogFile = True
    GoTo Log_Fin

Log_Erreur:
    MsgBox "Impossible d'�crire dans le fichier de suivi" & vbCr & LogFile, _
        vbOKOnly Or vbCritical, App.Title & " - Fichier LOG"
    OkLogFile = False
    Resume Log_Fin

Log_Fin:
    On Error Resume Next
    Close #fflog

End Sub

Public Sub LanceEtAttendShell(ByVal cmdLine As String, Style As VbAppWinStyle)

    ' Cette routine lance la commande donn�e dans cmdLine,
    ' puis attend la fin de son execution avant de rendre la main

    Dim retVal As Long, PiD As Long, pHandle As Long
    
    PiD = Shell(cmdLine, Style)
'    pHandle = OpenProcess(&H100000, True, PiD)
'    retVal = WaitForSingleObject(pHandle, Infini)

End Sub

