VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Forme 
   Caption         =   "Chat"
   ClientHeight    =   6420
   ClientLeft      =   1020
   ClientTop       =   6420
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLanChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   11430
   Begin MSComctlLib.ImageList imgListe 
      Left            =   9240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLanChat.frx":000C
            Key             =   "Ex Vert"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLanChat.frx":09D8
            Key             =   "ZZzz"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLanChat.frx":0B34
            Key             =   "Vert"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLanChat.frx":1988
            Key             =   "Rouge"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLanChat.frx":27DC
            Key             =   "ex Rouge"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLanChat.frx":2938
            Key             =   "Bleu"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLanChat.frx":3304
            Key             =   "Noir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstPostes 
      Height          =   3135
      Left            =   7080
      TabIndex        =   17
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   5530
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      Icons           =   "imgListe"
      SmallIcons      =   "imgListe"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer tmrFlash 
      Interval        =   300
      Left            =   240
      Top             =   5400
   End
   Begin VB.Timer tmrSurveillance 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   4920
   End
   Begin RichTextLib.RichTextBox lstMsgRecus 
      Height          =   2775
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmLanChat.frx":3460
   End
   Begin VB.Frame fGroupeDebug 
      Caption         =   "Debug"
      Height          =   2775
      Left            =   840
      TabIndex        =   2
      Top             =   3360
      Width           =   10095
      Begin VB.ListBox lstPostesCache 
         Height          =   2160
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   1455
      End
      Begin MSComctlLib.TreeView tvwReseau 
         Height          =   2415
         Left            =   7320
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   4260
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imlNWImages 
         Left            =   9480
         Top             =   2160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":34D2
               Key             =   "directory"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":3824
               Key             =   "root"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":3B76
               Key             =   "group"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":3EC8
               Key             =   "ndscontainer"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":421A
               Key             =   "network"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":456C
               Key             =   "server"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":48BE
               Key             =   "tree"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":4C10
               Key             =   "domain"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":4F62
               Key             =   "share"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":52B4
               Key             =   "adminshare"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":5606
               Key             =   "printer"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":5718
               Key             =   "folder"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLanChat.frx":5A6A
               Key             =   "file"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblConn 
         Caption         =   "Label1"
         Height          =   255
         Index           =   9
         Left            =   1680
         TabIndex        =   16
         Top             =   2520
         Width           =   5535
      End
      Begin VB.Label lblConn 
         Caption         =   "Label1"
         Height          =   255
         Index           =   8
         Left            =   1680
         TabIndex        =   15
         Top             =   2280
         Width           =   5535
      End
      Begin VB.Label lblConn 
         Caption         =   "Label1"
         Height          =   255
         Index           =   7
         Left            =   1680
         TabIndex        =   14
         Top             =   2040
         Width           =   5535
      End
      Begin VB.Label lblConn 
         Caption         =   "Label1"
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   13
         Top             =   1800
         Width           =   5535
      End
      Begin VB.Label lblConn 
         Caption         =   "Label1"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   12
         Top             =   1560
         Width           =   5535
      End
      Begin VB.Label lblConn 
         Caption         =   "Label1"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   11
         Top             =   1320
         Width           =   5535
      End
      Begin VB.Label lblConn 
         Caption         =   "Label1"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   10
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Label lblConn 
         Caption         =   "Label1"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   9
         Top             =   840
         Width           =   5535
      End
      Begin VB.Label lblConn 
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   8
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label xLabel1 
         Caption         =   "Prochaine mise à jour dans"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label xLabel1 
         Caption         =   "Liste cachée"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6165
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   450
      SimpleText      =   "Prêt"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Init"
            TextSave        =   "Init"
            Object.ToolTipText     =   "Action en cours"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1499
            TextSave        =   "11/06/2021"
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   979
            MinWidth        =   970
            TextSave        =   "9:18"
            Object.ToolTipText     =   "Heure"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wskConnexions 
      Index           =   0
      Left            =   240
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskEcoute 
      Left            =   240
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrMaJPostes 
      Interval        =   60000
      Left            =   240
      Top             =   4440
   End
   Begin VB.TextBox txtMessage 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   6735
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter!"
   End
   Begin VB.Menu mnuxOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuLogFile 
         Caption         =   "&Avec fichier LOG"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLogFileVoir 
         Caption         =   "&Voir le fichier de LOG"
      End
      Begin VB.Menu mnuzbarre 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuxExtranet 
      Caption         =   "&Extranet"
      Begin VB.Menu mnuAdresseExtraNet 
         Caption         =   "&Adresse IP"
      End
   End
   Begin VB.Menu mnuA_Propos 
      Caption         =   "&A propos de ..."
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "Forme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUsername Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private NetRoot As NetResource
Dim strLastMessages(20)     As String   ' Mémoire des messages tapés
Dim iIndexMessage           As Integer
Dim iIndexMessageOld        As Integer
Dim iPosMouseDown           As Integer


Private Sub Form_Load()
    
    Dim Temp As String
    
    Me.Caption = App.Title
    ' Manipulation si en mode debug
    If Not bDebug Then Me.WindowState = vbMinimized
    fGroupeDebug.Visible = bDebug
    If bDebug Then Debug.Print "-------------------------------------------------------------"
    DoEvents
    
    ' Accès aux fonctions fichier
    On Error GoTo Erreur
    Set Sys = CreateObject("Scripting.FileSystemObject")

    ' Tempo avant prochain rafraichissement
    sTimeDebut = Timer
    ' init Tempo détection touche appuyée
    sTimeTouche = Timer
    ' init messages en mémoire
    iIndexMessage = -1
    
    ' Fichier de log
    LogFile = App.Path & "\" & App.Title & ".log"
    Call Nettoie_Fichier_Log    ' Nettoie les lignes trop vieilles (voir Options)
    
    ' Test s'il existe une carte son
    bCarteSon = False
    If waveOutGetNumDevs <> 0 Then bCarteSon = True
    
    'Recherche caratéristiques de mon poste (strNomMachine et strNomUser)
    Call NomMachine
    
    ' ------- Récupère les paramètres de la dernière fois
    ' Position de la forme
    Me.Left = GetSetting(AppSite, App.Title, "MainPosGauche", 1000)
        If Me.Left > Screen.Width Then Me.Left = 0
    Me.Top = GetSetting(AppSite, App.Title, "MainPosHaut", 1000)
        If Me.Top > Screen.Height Then Me.Top = 0
    Me.Height = GetSetting(AppSite, App.Title, "MainHauteur", 7000)
    Me.Width = GetSetting(AppSite, App.Title, "MainLargeur", 10000)
    ' Pseudo
    strNomUser = GetSetting(AppSite, App.Title, "Pseudo", strNomUser)
    ' Avec/Sans fichier Log + purge
    mnuLogFile.Checked = GetSetting(AppSite, App.Title, "Avec LogFile", True)
    iPurgeLog = GetSetting(AppSite, App.Title, "Purge log (jours)", 7)
    ' Période de rafraichissement des connectés (en cas de perte de connexion)
    iPeriodeRafr = GetSetting(AppSite, App.Title, "Rafraichissement postes (min)", 10)
    ' Sons Message
    iAvecSonMessage = GetSetting(AppSite, App.Title, "Son Wave (nouveau message) Avec", vbChecked)
    strSonMessage = GetSetting(AppSite, App.Title, "Son Wave (nouveau message)", App.Path & "\Message.wav")
    ' Sons Arrivée
    iAvecSonArrivée = GetSetting(AppSite, App.Title, "Son Wave (nouveau contact) Avec", vbChecked)
    strSonArrivée = GetSetting(AppSite, App.Title, "Son Wave (nouveau contact)", App.Path & "\Arrivée.wav")

    ' paramètres Heure
    AffHeure.lCouleur = GetSetting(AppSite, App.Title, "Affichage Heure (couleur)", &H8000&)
    AffHeure.strFonte = GetSetting(AppSite, App.Title, "Affichage Heure (fonte)", "Arial")
    AffHeure.dTaille = GetSetting(AppSite, App.Title, "Affichage Heure (taille)", 8)
    AffHeure.bBold = GetSetting(AppSite, App.Title, "Affichage Heure (gras)", False)
    AffHeure.bItalic = GetSetting(AppSite, App.Title, "Affichage Heure (italique)", False)
    ' paramètres Pseudo
    AffPseudo.lCouleur = GetSetting(AppSite, App.Title, "Affichage Pseudo (couleur)", &HC00000)
    AffPseudo.strFonte = GetSetting(AppSite, App.Title, "Affichage Pseudo (fonte)", "Arial")
    AffPseudo.dTaille = GetSetting(AppSite, App.Title, "Affichage Pseudo (taille)", 8)
    AffPseudo.bBold = GetSetting(AppSite, App.Title, "Affichage Pseudo (gras)", False)
    AffPseudo.bItalic = GetSetting(AppSite, App.Title, "Affichage Pseudo (italique)", False)
    ' paramètres Système
    AffSystem.lCouleur = GetSetting(AppSite, App.Title, "Affichage System (couleur)", &H80000012)
    AffSystem.strFonte = GetSetting(AppSite, App.Title, "Affichage System (fonte)", "Arial")
    AffSystem.dTaille = GetSetting(AppSite, App.Title, "Affichage System (taille)", 8)
    AffSystem.bBold = GetSetting(AppSite, App.Title, "Affichage System (gras)", False)
    AffSystem.bItalic = GetSetting(AppSite, App.Title, "Affichage System (italique)", True)
    ' paramètres Perso
    AffPerso.lCouleur = GetSetting(AppSite, App.Title, "Affichage Perso (couleur)", &H80FF&)
    AffPerso.strFonte = GetSetting(AppSite, App.Title, "Affichage Perso (fonte)", "Arial")
    AffPerso.dTaille = GetSetting(AppSite, App.Title, "Affichage Perso (taille)", 8)
    AffPerso.bBold = GetSetting(AppSite, App.Title, "Affichage Perso (gras)", False)
    AffPerso.bItalic = GetSetting(AppSite, App.Title, "Affichage Perso (italique)", False)
    ' paramètres Autres
    AffAutres.lCouleur = GetSetting(AppSite, App.Title, "Affichage Autre (couleur)", vbBlack)
    AffAutres.strFonte = GetSetting(AppSite, App.Title, "Affichage Autre (fonte)", "Arial")
    AffAutres.dTaille = GetSetting(AppSite, App.Title, "Affichage Autre (taille)", 8)
    AffAutres.bBold = GetSetting(AppSite, App.Title, "Affichage Autre (gras)", False)
    AffAutres.bItalic = GetSetting(AppSite, App.Title, "Affichage Autre (italique)", False)
    
    Me.Refresh
    DoEvents
    
    ' Gestion des icones à gauche du nom des postes (cercle vert ou rouge, selon l'activité)
    lstPostes.SmallIcons = imgListe
    
    ' Initialisation de la liste des postes
    tvwReseau.ImageList = imlNWImages
    Set NetRoot = New NetResource
    
    'Init des WinSocks
    wskEcoute.LocalPort = 4012
    wskEcoute.Protocol = sckTCPProtocol
    wskEcoute.Listen
    
    ' Teste la présence d'un fichier de log
    Call Ecrit_Log("TEST")

    ' Mémo Info de connexion initiale
    Temp = "SYS----------- Lancement (" & strNomUser & ", " & strNomMachine & ", " & _
                                       wskEcoute.LocalIP & ":" & wskEcoute.LocalPort & ")"
    If bDebug Then Call Ecrit_Log(Temp)

    'Message de bienvenue
    With lstMsgRecus
        .SelStart = 0
        .SelBold = True
        .SelItalic = False
        .SelFontName = "Arial"
        .SelFontSize = "10"
        .SelColor = &H8000& ' vert
        .SelText = "Bonjour " & strNomUser & " et bienvenu(e) sur le chat " & App.Title
        .SelStart = Len(.Text)
        .SelBold = False
        .SelFontSize = "8"
        .SelText = " (v " & App.Major & "." & App.Minor & "." & App.Revision & ")" & vbCrLf
        .SelStart = Len(.Text)
        .SelColor = &H80FF& ' orangé
        .SelText = "Va dans le menu ""Options"" si tu veux changer ton Pseudo" & vbCrLf & vbCrLf
    End With
    
    ' Lance la première recherche des postes
    tmrSurveillance.Enabled = True
    
    Exit Sub
    
Erreur:
    If Err.Number = 10048 Then
        Temp = App.Title & " est déjà en cours d'exécution." & vbCr
        Temp = Temp & "Il n'est pas possible de lancer plusieurs cessions en même temps."
        MsgBox Temp, vbCritical Or vbOKOnly, App.Title & " (Main)"
    Else
        Temp = "Erreur " & Format(Err.Number) & vbCr
        Temp = Temp & Err.Description & vbCr & vbCr
        Call Ecrit_Log("(FormLoad) " & Temp, True)
        MsgBox Temp, vbCritical Or vbOKOnly, App.Title & " (Main)"
    End If
    End
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        Cancel = 1  ' Annule la fermeture
        Me.WindowState = vbMinimized
    End If

End Sub

Private Sub Form_Resize()
    
    Dim i As Double
    Dim Décalage As Double  ' Espace libre sous le txtMessage
    Dim MenuHtr As Double, BandeHtr As Double, BandeLgr As Double
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If bDebug Then
        Décalage = 3000 ' Zone laissée libre sous les 3 controles
    Else
        Décalage = 0
    End If
    
    ' Dimension mini de la feuille
    If Forme.Height < (4260 + Décalage) Then Forme.Height = 4260 + Décalage
    If Forme.Width < 7245 Then Forme.Width = 7245
    
    ' Dimension du Status en bas de la forme
    i = Forme.ScaleWidth - Status.Panels(2).Width - Status.Panels(2).Width - 200
    If i < 10 Then i = 10
    Status.Panels(1).Width = i
    
    ' Position et taille des controles de la feuille
    MenuHtr = 0   ' Hauteur du menu
    BandeLgr = 50   ' Largeur de la Bande entre deux controles
    BandeHtr = 40   ' Hauteur ...
    lstPostes.Width = 1815 ' Taille fixe
    txtMessage.Height = 315     ' Hauteur fixe
    lstMsgRecus.Top = MenuHtr
    lstMsgRecus.Left = BandeLgr
    lstPostes.Top = lstMsgRecus.Top
    lstPostes.Left = Forme.ScaleWidth - lstPostes.Width - 2 * BandeLgr
    lstMsgRecus.Width = lstPostes.Left - lstMsgRecus.Left - BandeLgr
        If lstMsgRecus.Width < 5110 Then lstMsgRecus.Width = 5110
    lstMsgRecus.Height = Forme.ScaleHeight - MenuHtr - Status.Height - txtMessage.Height - _
                         BandeHtr - Décalage
        If lstMsgRecus.Height < 2580 Then lstMsgRecus.Height = 2580
    txtMessage.Top = lstMsgRecus.Top + lstMsgRecus.Height + BandeHtr
    txtMessage.Width = lstMsgRecus.Width
    txtMessage.Left = lstMsgRecus.Left
    lstPostes.Height = Forme.ScaleHeight - MenuHtr - Status.Height - Décalage
    
    ' paramètrage de la fenêtre d'infos technique si mode Debug
    If bDebug = False Then Exit Sub
    fGroupeDebug.Top = txtMessage.Top + txtMessage.Height + BandeHtr
    fGroupeDebug.Left = BandeLgr
    fGroupeDebug.Height = Décalage - 2 * BandeHtr
    fGroupeDebug.Width = Forme.ScaleWidth - 2 * BandeLgr
    lstPostesCache.Height = fGroupeDebug.Height - 6 * BandeHtr - xLabel1(0).Height
    tvwReseau.Height = fGroupeDebug.Height - 8 * BandeHtr
    tvwReseau.Left = fGroupeDebug.Width - 2 * BandeLgr - tvwReseau.Width  ' à droite de fenêtre
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim r As Integer
    
    tmrSurveillance.Enabled = False
    tmrMaJPostes.Enabled = False
    
    ' Sauvegarde les paramètres
    If Me.WindowState <> vbMinimized Then
        SaveSetting AppSite, App.Title, "MainPosGauche", Me.Left
        SaveSetting AppSite, App.Title, "MainPosHaut", Me.Top
        SaveSetting AppSite, App.Title, "MainHauteur", Me.Height
        SaveSetting AppSite, App.Title, "MainLargeur", Me.Width
    End If
    SaveSetting AppSite, App.Title, "Avec LogFile", mnuLogFile.Checked
    
    ' Décharge tous les controles WinSock
    On Error Resume Next
    r = 1
    Do While r <= 200
        If Connexions(r).iNoControl <> 0 Then
            If wskConnexions(Connexions(r).iNoControl).State <> sckClosed Then _
                wskConnexions(Connexions(r).iNoControl).Close
            DoEvents
            Unload wskConnexions(Connexions(r).iNoControl)
        End If
        r = r + 1
        DoEvents
    Loop

    End

End Sub

Private Sub lstMsgRecus_KeyPress(KeyAscii As Integer)
    
    ' Pas de texte dans ce controle
    KeyAscii = 0

End Sub

Private Sub lstMsgRecus_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Len(lstMsgRecus.SelText) <> 0 Then
        txtMessage.Text = txtMessage.Text + lstMsgRecus.SelText
        lstMsgRecus.SelStart = Len(lstMsgRecus.Text)       ' On se remet à la fin
    End If
    txtMessage.SetFocus
    txtMessage.SelStart = Len(txtMessage.Text)

End Sub

Private Sub lstPostes_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    ' La sélection vient de changer
    ' Si <Tous> est coché, on décoche les autres
    ' Si un autre est coché, on décoche <Tous>
    
    Dim r As Integer, t As Boolean
    
    If lstPostes.ListItems.Count = 0 Then Exit Sub
    If Item.Index = 1 Then
        ' <Tous> vient de changer d'état
        If Item.Checked Then
            For r = 2 To lstPostes.ListItems.Count
                lstPostes.ListItems.Item(r).Checked = False
            Next r
        End If
    Else
        ' Sinon, décoche <Tous>
        lstPostes.ListItems.Item(1).Checked = False
    End If
    
    ' Coche <Tous> si aucun poste particulier n'est coché
    t = False
    For r = 2 To lstPostes.ListItems.Count
        If lstPostes.ListItems.Item(r).Checked = True Then t = True
    Next r
    If Not t Then lstPostes.ListItems.Item(1).Checked = True

End Sub

Private Sub lstPostes_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Item.Checked = Not Item.Checked
    Call lstPostes_ItemCheck(Item)
    
End Sub

Private Sub mnuA_Propos_Click()

    frmAbout.Show vbModeless, Me

End Sub

Private Sub mnuAdresseExtraNet_Click()

    Dim Temp As String, i As Integer, z As Integer, NoCtrl As Integer, Ret As Long
    
    strIPextraNet = ""
    frmExtraNet.Show vbModal, Me
    
    If strIPextraNet = "" Then Exit Sub
    
    mnuAdresseExtraNet.Enabled = False
    ' Essaie de se connecter à une adresse IP à l'extérieur du réseau
    i = 1   ' Recherche d'une connexion disponible dans le tableau
    Do While i <= 200
        If Connexions(i).strPoste = "" And _
            Connexions(i).iNoControl = 0 Then
            Exit Do
        Else
            i = i + 1
        End If
    Loop
            
    ' Recherche un numéro d'index pour le futur controle
    NoCtrl = 1
    For z = 1 To 200
        If Connexions(z).iNoControl = NoCtrl Then
            NoCtrl = NoCtrl + 1
            z = 0
            DoEvents
        End If
    Next z
            
    ' On créé un nouveau control WinSocks
    Load wskConnexions(NoCtrl)
    wskConnexions(NoCtrl).Protocol = sckTCPProtocol
    wskConnexions(NoCtrl).RemoteHost = strIPextraNet
    wskConnexions(NoCtrl).RemotePort = 4012
    DoEvents
    
    ' Fait une demande de connexion
    wskConnexions(NoCtrl).Connect
    DoEvents
    Do While wskConnexions(NoCtrl).State = sckConnecting
                            ' 0 sckClosed
                            ' 1 sckOpen
                            ' 3 sckConnectionPending
                            ' 4 sckResolvingHost
                            ' 5 sckHostResolved
                            ' 6 sckConnecting
                            ' 7 sckConnected
                            ' 8 sckClosing
                            ' 9 sckError
        DoEvents
    Loop
    If wskConnexions(NoCtrl).State <> sckConnected Then
        ' Si la connexion n'a pas répondu, on annule tout
        If wskConnexions(NoCtrl).State <> sckClosed Then wskConnexions(NoCtrl).Close
        DoEvents
        Unload wskConnexions(NoCtrl)
        Beep
        Status.Panels(1).Text = "La connexion à """ & strIPextraNet & """ a échouée."
    Else
        ' Envoi mes caractéristiques
        wskConnexions(NoCtrl).SendData "Machine" & Chr(0) & _
                                        strNomMachine & Chr(0) & _
                                        strNomUser & Chr(0) & Chr(1)
        DoEvents
        ' trace
        Temp = "Nouvelle connexion (ExtraNet " & strIPextraNet & ")"
            Status.Panels(1).Text = Temp
            Call Ecrit_Log(Temp)
        DoEvents
        ' Sinon, on mémorise les caractéristiques
        Connexions(i).iNoControl = NoCtrl
        Connexions(i).strPoste = strIPextraNet
        Connexions(i).strPseudo = strIPextraNet
        Connexions(i).strIP = strIPextraNet
        Call EcritMessage("Système", """" & strIPextraNet & """" & _
                    " vient de se connecter (Port " & wskConnexions(NoCtrl).LocalPort & ").")
        If lstPostes.ListItems.Count = 0 Then
            lstPostes.ListItems.Add 1, , "<Tous>", , "Noir"
            lstPostes.ListItems.Item(1).Checked = True
        End If
        lstPostes.ListItems.Add , , strIPextraNet, , "Rouge"
        On Error Resume Next
            If iAvecSonArrivée Then Ret = sndPlaySound(strSonArrivée, 1)
        On Error GoTo 0
    End If

    mnuAdresseExtraNet.Enabled = True

End Sub

Private Sub mnuLogFile_Click()

    mnuLogFile.Checked = Not mnuLogFile.Checked
    
    If mnuLogFile.Checked = False And OkLogFile = True Then _
        Call Ecrit_Log("SYSArrêt du fichier de suivi (LOG)")
    
    If mnuLogFile.Checked = True Then _
        Call Ecrit_Log("SYSMise en service du fichier de suivi (LOG)")

End Sub

Private Sub mnuLogFileVoir_Click()

    Dim Temp As String
    
    If Not Sys.FileExists(LogFile) Then
        Status.Panels(1).Text = "Le fichier LOG est vide."
        Exit Sub
    End If
    
    Temp = LogFile & " (Copie)"
    Sys.CopyFile LogFile, Temp, True
    Call LanceEtAttendShell("notepad " & Temp, vbNormalFocus)
    If Sys.FileExists(Temp) Then Sys.DeleteFile (Temp)

End Sub

Private Sub mnuOptions_Click()

    frmOptions.Show vbModeless, Me
    
End Sub

Private Sub mnuQuitter_Click()

    Form_Unload (False)

End Sub

Private Sub tmrFlash_Timer()

    Dim Ret As Long
    
    On Error Resume Next    ' au cas ou le fichier wav serait introuvable
    
    ' Fait clignoter la fenêtre ou son icône dans la barre des tâches
    If bNouveauMessage Then Call FlashWindow(Me.hwnd, True)
    
    ' Joue le son "Nouveau message" s'il vient d'arriver (le message)
    If bNouveauMessage And Not bNouveauMessageVu Then
        bNouveauMessageVu = True
        If bCarteSon And Me.hwnd <> GetForegroundWindow() And _
           iAvecSonMessage Then Ret = sndPlaySound(strSonMessage, 1)
    End If

End Sub

Private Sub tmrMaJPostes_Timer()
    
    iCompteMinutes = iCompteMinutes + 1
    If iCompteMinutes >= iPeriodeRafr Then
        Call PostesNom
        iCompteMinutes = 0
        ' init Tempo avant prochain rafraichissement
        sTimeDebut = Timer
    End If
    
End Sub

Private Sub tmrSurveillance_Timer()

    ' On fait le ménage si la connexion n'est plus Connectée
    
    Dim i As Integer, r As Integer, Temp As String, Ret As Long
    
    ' Fait clignoter la fenêtre si un nouveau message est arrivé et que
    ' l'on est en icône ou que la fenetre courante n'est pas nous même personnellement
    If Me.WindowState <> vbMinimized And Me.hwnd = GetForegroundWindow() Then
        bNouveauMessage = False
        bNouveauMessageVu = False
    End If

    ' Si c'est la première fois qu'on arrive ici, on recherche les postes connectés
    If Status.Panels(1) = "Init" Then
        Call PostesNom
        Status.Panels(1).Text = "Prêt"
        Exit Sub
    End If
    
    ' En mode debug, affiche temps restant avant prochaine mise à jour
    If bDebug Then
        Temp = Format(Fix((iPeriodeRafr * 60) - (Timer - sTimeDebut)))
        xLabel1(1).Caption = "Prochaine mise à jour dans " & Temp & " sec"
        For r = 1 To 9
            lblConn(r).Caption = Format(Connexions(r).iNoControl) & "-" & _
                                 Connexions(r).strIP & ", " & _
                                 Connexions(r).strPoste & ", " & _
                                 Connexions(r).strPseudo & _
                                 ", Etat : " & wskConnexions(Connexions(r).iNoControl).State
        Next r
    End If

    ' Supprime les connexions qui sont déconnectées
    i = 1
    Do While i <= 200
        If Connexions(i).iNoControl <> 0 Then
            If wskConnexions(Connexions(i).iNoControl).State <> sckConnected Then
                ' Ferme et décharge la connexion
                wskConnexions(Connexions(i).iNoControl).Close
                DoEvents
                Unload wskConnexions(Connexions(i).iNoControl)
                DoEvents
                ' Supprime le poste de la lstPostes
                r = 1
                Do While r <= lstPostes.ListItems.Count
                    If lstPostes.ListItems.Item(r) = Connexions(i).strPseudo Or _
                        lstPostes.ListItems.Item(r) = Connexions(i).strIP Then
                        Call EcritMessage("Système", """" & Connexions(i).strPseudo & """" & _
                                " vient de se déconnecter.")
                        lstPostes.ListItems.Remove r
                    Else
                        r = r + 1
                    End If
                Loop
                ' RaZ la mémoire
                Connexions(i).iNoControl = 0
                Connexions(i).strIP = ""
                Connexions(i).strPoste = ""
                Connexions(i).strPseudo = ""
            Else
                If Connexions(i).strPoste = "" Then _
                    wskConnexions(Connexions(i).iNoControl).SendData "Qui ?" & Chr(0) & Chr(1)
            End If
        End If
        i = i + 1
        DoEvents
    Loop
    
    ' Supprime <Tous> s'il est tout seul (le pauvre)
    If lstPostes.ListItems.Count = 1 Then
        If lstPostes.ListItems.Item(1) = "<Tous>" Then lstPostes.ListItems.Remove 1
    Else
        ' Recherche les IP ou nom de poste pour les remplacer par les Pseudos
        i = 2
        Do While i <= lstPostes.ListItems.Count
            r = 1
            Do While r <= 200
                If lstPostes.ListItems.Item(i) = Connexions(r).strIP Or _
                   lstPostes.ListItems.Item(i) = Connexions(r).strPoste Then
                    ' on a trouvé une IP ou un nom de poste dans la liste
                    '   (au lieu du pseudo) : On regarde si le pseudo est arrivé
                    If Connexions(r).strPoste <> "" And _
                       Connexions(r).strPseudo <> lstPostes.ListItems.Item(i) Then
                        Temp = """" & lstPostes.ListItems.Item(i) & _
                               """ change son pseudo en """ & _
                               Connexions(r).strPseudo & """"
                        Call EcritMessage("Système", Temp)
                        lstPostes.ListItems.Item(i) = Connexions(r).strPseudo
                        Exit Do
                    End If
                End If
                r = r + 1
            Loop
            i = i + 1
        Loop
        
        ' Si on a appuyé sur une touche (sTimeTouche est réinitialisé), on
        '   envoie le message "Message en cours de composition" (Ok)
        ' Si ça fait plus de 20 sec qu'on n'a pas appuyé sur une touche
        '   on envoie "Pas de touche enfoncée depuis 20 sec" (Pas Ok)
        If Timer - sTimeTouche > 20 Then
            ' Plus de 10 sec depuis le dernier appui sur une touche
            If bTimeTouche Then
                bTimeTouche = False
                i = 2
                Do While i <= lstPostes.ListItems.Count
                    r = 1
                    Do While r <= 200
                        If lstPostes.ListItems.Item(i) = Connexions(r).strPseudo Or _
                            lstPostes.ListItems.Item(i) = Connexions(r).strIP Then
                            If wskConnexions(Connexions(r).iNoControl).State = sckConnected Then
                                wskConnexions(Connexions(r).iNoControl).SendData _
                                    "Compose" & Chr(0) & "Pas Ok" & Chr(1)
                            Exit Do
                            End If
                        End If
                        r = r + 1
                    Loop
                    DoEvents
                    i = i + 1
                Loop
            End If
        Else
            ' Moins de 10 sec depuis le dernier appui sur une touche
            If Not bTimeTouche Then
                bTimeTouche = True
                i = 2
                Do While i <= lstPostes.ListItems.Count
                    r = 1
                    Do While r <= 200
                        If lstPostes.ListItems.Item(i) = Connexions(r).strPseudo Or _
                            lstPostes.ListItems.Item(i) = Connexions(r).strIP Then
                            If wskConnexions(Connexions(r).iNoControl).State = sckConnected Then
                                wskConnexions(Connexions(r).iNoControl).SendData _
                                    "Compose" & Chr(0) & "Ok" & Chr(1)
                            Exit Do
                            End If
                        End If
                        r = r + 1
                    Loop
                    DoEvents
                    i = i + 1
                Loop
            End If
        End If
    End If

End Sub

Private Sub tvwReseau_Expand(ByVal Node As MSComctlLib.Node)
    If Node.Tag = "N" Then
        ChercheNode Node, lstPostesCache
    End If
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim r As Integer, i As Integer
    
    ' Une touche a été frappée : On réinitialise le Chrono de 20 sec
    sTimeTouche = Timer
    
    ' Si on a tapé flèche-haut ou flèche-bas ET qu'il y a des messages en memoire
    If Shift = 0 And (KeyCode = 38 Or KeyCode = 40) And iIndexMessage <> -1 Then
        Select Case KeyCode
            Case 38 ' ----------- Flèche-haut
                iIndexMessageOld = iIndexMessageOld - 1   ' décrémente le n° du message
                If iIndexMessageOld < 0 Then
                    iIndexMessageOld = -1
                    txtMessage.Text = ""
                    Beep
                Else            ' Si on a déjà fait un tour complet
                    txtMessage.Text = strLastMessages(iIndexMessageOld)
                    txtMessage.SelStart = Len(txtMessage.Text)  ' on se place à la fin
                End If
            Case 40 ' ----------- Flèche-bas
                iIndexMessageOld = iIndexMessageOld + 1   ' incrémente le n° du message
                If iIndexMessageOld > UBound(strLastMessages) Then  ' Si on dépasse le tableau
                    iIndexMessageOld = UBound(strLastMessages)
                    txtMessage.Text = ""
                    Beep
                Else
                    If iIndexMessageOld > 0 Then
                        If strLastMessages(iIndexMessageOld - 1) = "" Then ' S'il n'y a rien a afficher
                            iIndexMessageOld = iIndexMessageOld - 1
                            txtMessage.Text = ""
                            Beep
                        Else
                            txtMessage.Text = strLastMessages(iIndexMessageOld)
                            txtMessage.SelStart = Len(txtMessage.Text)  ' on se place à la fin
                        End If
                    Else
                        txtMessage.Text = strLastMessages(iIndexMessageOld)
                        txtMessage.SelStart = Len(txtMessage.Text)  ' on se place à la fin
                    End If
                End If
        End Select
        KeyCode = 0 ' supprime touche fleche du buffer
    End If

    ' Si on a tapé autre chose d'autre que <Entrée> ou s'il n'y a pas de message
    If KeyCode <> Asc(vbCr) Or txtMessage.Text = "" Then Exit Sub
    
    ' Teste s'il y a des destinataires possibles
    If lstPostes.ListItems.Count = 0 Then
        'Beep
        Status.Panels(1).Text = "Pas de poste accessible actuellement."
        Exit Sub
    End If
    
    ' Mémorise se nouveau message
    iIndexMessage = iIndexMessage + 1
    If iIndexMessage > 100 Then iIndexMessage = 0
    strLastMessages(iIndexMessage) = txtMessage.Text
    
    ' Teste en fonction du/des destinataires
    ' Choix n° 1 (donc 0) = <Tous>
    
    ' Sélection = <Tous> ?
    If lstPostes.ListItems.Item(1).Checked Then
        For i = 1 To 200
            If Connexions(i).iNoControl <> 0 Then
                If wskConnexions(Connexions(i).iNoControl).State = sckConnected Then
                    wskConnexions(Connexions(i).iNoControl).SendData _
                            "Message" & Chr(0) & txtMessage.Text & Chr(1)
                End If
            End If
        Next i
        ' Un echo pour moi aussi
        Call EcritMessage(strNomUser, txtMessage.Text)
    Else
    ' Sélection du ou des destinataires :
        For i = 1 To 200
            If Connexions(i).iNoControl <> 0 Then
                For r = 1 To lstPostes.ListItems.Count
                    ' Recherche dans la liste le poste correspondant à 'i'
                    If lstPostes.ListItems.Item(r) = Connexions(i).strPseudo Then
                        If lstPostes.ListItems.Item(r).Checked Then
                            If wskConnexions(Connexions(i).iNoControl).State = sckConnected Then
                                wskConnexions(Connexions(i).iNoControl).SendData _
                                            "Message" & Chr(0) & "(privé) " & txtMessage.Text & Chr(1)
                                ' Un echo pour moi, mais en privé
                                Call EcritMessage(strNomUser & "-->" & Connexions(i).strPseudo, _
                                            txtMessage.Text)
                            End If
                        End If
                    End If
                Next r
            End If
            DoEvents
        Next i
    End If
    txtMessage.Text = ""

End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)

    If KeyAscii = Asc(vbCr) Then
        KeyAscii = 0
        iIndexMessageOld = iIndexMessage
    End If

End Sub

Private Sub wskConnexions_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'Des données arrivent

    Dim Données As String, Clé As String, Texte As String
    Dim Temp As String, NoCtrl As Integer, r As Integer, i As Integer, iItem As Integer
    
    'Recherche le user associé à cet index
    r = 0
    NoCtrl = 1
    Do While NoCtrl <= 200
        If Connexions(NoCtrl).iNoControl = Index Then
            r = NoCtrl
            Exit Do
        Else
            NoCtrl = NoCtrl + 1
        End If
    Loop
    ' On ressort si on n'a pas trouvé la mémoire associée à ce control
    If r = 0 Then Exit Sub

    '------- Les messages ont la structure suivante :
    ' MotClé chr(0) Donnée1 chr(0) Donnée2 chr(0) DonnéeX ... chr(1) MotClé chr(0) Donnée1 chr(0) Donnée2 chr(0) DonnéeX ... chr(1) ...
    ' C'est à dire que plusieurs messages peuvent arriver en même temps :
    '   Il sont isolés par le 'chr(1)'
    ' A l'intérieur de chaque message, les données sont séparées par des 'chr(0)'
    
    'Lecture des données
    wskConnexions(Index).GetData Temp, , bytesTotal
    r = 1 ' 1ère partie
Suite_10:
    ' Isole la 'r'ème partie du message
    Données = Element(Temp, r, Chr(1))
    ' On arrête s'il n'y a plus rien à traiter
    If Données = "" Then Exit Sub
    'Isole le mot clé initial et le contenu
    Clé = Element(Données, 1, Chr(0))
    Texte = Element(Données, 2, Chr(0))
    
    Select Case Clé
        Case "Machine"   ' Nom de la machine
            Connexions(NoCtrl).strPoste = Element(Données, 2, Chr(0))
            Connexions(NoCtrl).strPseudo = Element(Données, 3, Chr(0))
        Case "Qui ?"    ' Demande d'identification
            wskConnexions(Index).SendData "Machine" & Chr(0) & strNomMachine & _
                                            Chr(0) & strNomUser & Chr(0) & Chr(1)
        Case "Message"
            'Affiche le message avec heure et user en tête
            Call EcritMessage(Connexions(NoCtrl).strPseudo, Texte)
            bNouveauMessage = True
        Case "Compose"
            ' Recherche le nom du poste dans lstPostes
            i = 1
            Do While i <= lstPostes.ListItems.Count
                If lstPostes.ListItems.Item(i) = Connexions(NoCtrl).strPseudo Or _
                    lstPostes.ListItems.Item(i) = Connexions(NoCtrl).strIP Then
                    Exit Do
                Else
                    i = i + 1
                End If
            Loop
            ' Affecte un rond Vert si ça tape dur (sur le clavier)
            '   de l'autre côté, sinon Rouge
            If Texte = "Ok" Then
                lstPostes.ListItems.Item(i).SmallIcon = "Vert"
            Else
                lstPostes.ListItems.Item(i).SmallIcon = "Rouge"
            End If
    End Select
    r = r + 1
    GoTo Suite_10

End Sub

Private Sub wskConnexions_Error(Index As Integer, ByVal Number As Integer, _
                                Description As String, ByVal Scode As Long, _
                                ByVal Source As String, ByVal HelpFile As String, _
                                ByVal HelpContext As Long, CancelDisplay As Boolean)

    Dim Temp As String, r As Integer, i As Integer
    
    Exit Sub    '##################################################################"
    
    ' Recherche de quel poste il s'agit
    Temp = ""
    r = 1
    Do While r <= 200
        If Connexions(r).iNoControl = Index Then
            Temp = "(wskConn) " & Format(Number) & " - " & Description & _
                   " (" & Connexions(r).strPoste & ")"
            ' Ferme le control si besoin
            If wskConnexions(Index).State <> sckClosed Then wskConnexions(Index).Close
            DoEvents
            ' Décharge le controle
            Unload wskConnexions(Index)
            ' Supprime le poste de la liste
            i = 1
            Do While i <= lstPostes.ListItems.Count
                If lstPostes.ListItems.Item(i) = Connexions(r).strPseudo Or _
                    lstPostes.ListItems.Item(i) = Connexions(r).strIP Then
                    Call EcritMessage("Système", """" & Connexions(r).strPseudo & """" & _
                                " vient de se déconnecter.")
                    lstPostes.ListItems.Remove i
                Else
                    i = i + 1
                End If
            Loop
            ' Annule les mémoires
            Connexions(r).iNoControl = 0
            Connexions(r).strPoste = ""
            Connexions(r).strPseudo = ""
            Connexions(r).strIP = ""
            Exit Do
        End If
        r = r + 1
    Loop
    ' S'il ne reste que "<Tous>", on l'enlève
    If lstPostes.ListItems.Count = 1 Then
       If lstPostes.ListItems.Item(1) = "<Tous>" Then lstPostes.ListItems.Remove 1
    End If
        
    If Temp = "" Then Temp = "(wskConn) " & Format(Number) & " - " & Description & _
                             " (Index inconnu " & Format(Index) & ")"
    Call Ecrit_Log(Temp, True)
    Status.Panels(1).Text = Temp
    
End Sub

Private Sub wskEcoute_ConnectionRequest(ByVal RequestID As Long)

    Dim r As Integer, NoCtrl As Integer, Temp As String, Ret As Long
    
    On Error GoTo Erreur
    
    'Un poste distant demande une connexion
    
    ' Annule la demande si on a déjà une connexion avec cette IP
    r = 1
    Do While r <= 200
        If Connexions(r).strIP = wskEcoute.RemoteHostIP Then
            Temp = "(ConnReq) Connexion refusée (doublon) avec " & wskEcoute.RemoteHostIP
                Status.Panels(1).Text = Temp
                Call Ecrit_Log(Temp)
            Exit Sub
        End If
        r = r + 1
        DoEvents
    Loop
    
    'Recherche un numéro d'index pour le futur control
    NoCtrl = 1
    For r = 1 To 200
        If Connexions(r).iNoControl = NoCtrl Then
            NoCtrl = NoCtrl + 1
            r = 0
            DoEvents
        End If
    Next r

    'On créé un nouveau control WinSocks pour valider la connexion avec lui
    Load wskConnexions(NoCtrl)
    wskConnexions(NoCtrl).Protocol = sckTCPProtocol

    'Accepte la transaction avec le requestID qui va bien
    wskConnexions(NoCtrl).Accept (RequestID)
    DoEvents
    Do While wskConnexions(NoCtrl).State = sckConnecting
                            ' 0 sckClosed
                            ' 1 sckOpen
                            ' 3 sckConnectionPending
                            ' 4 sckResolvingHost
                            ' 5 sckHostResolved
                            ' 6 sckConnecting
                            ' 7 sckConnected
                            ' 8 sckClosing
                            ' 9 sckError
        DoEvents
    Loop
    If wskConnexions(NoCtrl).State <> sckConnected Then
        If wskConnexions(NoCtrl).State <> sckClosed Then wskConnexions(NoCtrl).Close
        DoEvents
        Unload wskConnexions(NoCtrl)
        Exit Sub
    Else
        'Envoie la trame d'identification
        wskConnexions(NoCtrl).SendData "Machine" & Chr(0) & strNomMachine & _
                                      Chr(0) & strNomUser & Chr(0) & Chr(1)
        DoEvents
        
        ' Mémorise les paramètres de la connexion : trouve une place dans le tableau
        r = 1
        Do While r <= 200
            If Connexions(r).strPoste = "" And Connexions(r).iNoControl = 0 Then
                Connexions(r).iNoControl = NoCtrl
                Connexions(r).strPoste = ""
                Connexions(r).strPseudo = ""
                Connexions(r).strIP = wskEcoute.RemoteHostIP
                If lstPostes.ListItems.Count = 0 Then
                    lstPostes.ListItems.Add 1, , "<Tous>", , "Noir"
                    lstPostes.ListItems.Item(1).Checked = True
                End If
                lstPostes.ListItems.Add , , wskEcoute.RemoteHostIP, , "Rouge"
                On Error Resume Next
                    If iAvecSonArrivée Then Ret = sndPlaySound(strSonArrivée, 1)
                On Error GoTo 0
                Temp = "Réponse à connexion : " & wskEcoute.RemoteHostIP & ":" & _
                                                  wskEcoute.RemotePort & " - " & _
                                                  wskEcoute.RemoteHost
                Ecrit_Log (Temp)
                Exit Do
            End If
            r = r + 1
            DoEvents
        Loop
    End If
    Exit Sub
    
Erreur:
    Temp = "Erreur " & Format(Err.Number) & vbCr
    Temp = Temp & Err.Description & vbCr & vbCr
    On Error Resume Next
    Temp = Temp & "Demande de " & wskEcoute.RemoteHostIP
    MsgBox Temp, vbCritical Or vbOKOnly, App.Title & " (ConnReq)"
    Call Ecrit_Log("(ConnReq) " & Temp, True)
    
End Sub

Private Sub wskEcoute_Error(ByVal Number As Integer, Description As String, _
                            ByVal Scode As Long, ByVal Source As String, _
                            ByVal HelpFile As String, ByVal HelpContext As Long, _
                            CancelDisplay As Boolean)

    Dim Temp As String
    
    
    Exit Sub    '###############################################################"
    
    
    
    Temp = "(wskEcoute) " & Format(Number) & " - " & Description & vbCr
    Temp = Temp & "(Scode " & Format(Scode) & ") Source : " & Source
    Call Ecrit_Log(Temp, True)
    Temp = "Erreur Ecoute (Scode " & Format(Scode) & ") Source : " & Source
    Status.Panels(1).Text = Temp

End Sub

Private Sub PostesNom()

    Dim r As Integer, i As Integer, z As Integer, NoCtrl As Integer
    Dim Trouvé As Integer, Temp As String, strIP As String, Ret As Long
    
    ' point rouge
    If lstPostes.ListItems.Count > 0 Then
        lstPostes.ListItems.Item(1).SmallIcon = "Bleu"
        lstPostes.Refresh
    End If
    ' Efface la liste actuelle
    lstPostesCache.Clear
    ' initialise le treeview du réseau pour une nouvelle recherche
    Call InitTreeView
    ' Explore tous les noeuds
    With tvwReseau
        r = 1
        Do While r <= .Nodes.Count
            .Nodes(r).EnsureVisible
            r = r + 1
            DoEvents
        Loop
    End With
    
    '----------------------------------------------
    ' La liste 'lstPostesCache' contient tous les postes accessibles
    ' La liste 'lstPostes' contient la liste des connexions actives
    '   (Dans cette liste, le 1er choix = <Tous>)
    
Suite_10:
    '----------------------------------------------
    ' Teste s'il faut ajouter un item de la liste
    ' On regarde si tous les items de lstPostesCache sont dans lstPostes
    For r = 0 To lstPostesCache.ListCount - 1
        ' Recherche dans le tableau le pseudo associé au nom du poste
        Trouvé = 0
        NoCtrl = 1
        Do While NoCtrl <= 200
            If Connexions(NoCtrl).strPoste = lstPostesCache.List(r) Then
                Trouvé = 1
                Exit Do
            Else
                NoCtrl = NoCtrl + 1
            End If
            DoEvents
        Loop
        ' L'item pointé par r n'existe pas dans l'ancienne liste
        ' --> Ajout de la connexion
        If Trouvé = 0 Then
            i = 1   ' Recherche d'une connexion disponible
            Do While i <= 200
                If Connexions(i).strPoste = "" And _
                   Connexions(i).iNoControl = 0 Then
                    ' Recherche un numéro d'index pour le futur controle
                    NoCtrl = 1
                    For z = 1 To 200
                        If Connexions(z).iNoControl = NoCtrl Then
                            NoCtrl = NoCtrl + 1
                            z = 0
                            DoEvents
                        End If
                    Next z
                    ' Recherche l'IP à partir du nom
                    strIP = IPduPoste(lstPostesCache.List(r))
                    If Not (strIP = "" Or strIP = "Erreur") Then
                        ' On créé un nouveau control WinSocks
                        Load wskConnexions(NoCtrl)
                        wskConnexions(NoCtrl).Protocol = sckTCPProtocol
                        wskConnexions(NoCtrl).RemoteHost = strIP
                        wskConnexions(NoCtrl).RemotePort = 4012
                        DoEvents
                        ' Fait une demande de connexion
                        wskConnexions(NoCtrl).Connect
                        DoEvents
                        Do While wskConnexions(NoCtrl).State = sckConnecting
                            ' 0 sckClosed
                            ' 1 sckOpen
                            ' 3 sckConnectionPending
                            ' 4 sckResolvingHost
                            ' 5 sckHostResolved
                            ' 6 sckConnecting
                            ' 7 sckConnected
                            ' 8 sckClosing
                            ' 9 sckError
                            DoEvents
                        Loop
                        If wskConnexions(NoCtrl).State <> sckConnected Then
                            If wskConnexions(NoCtrl).State <> sckClosed Then _
                                wskConnexions(NoCtrl).Close
                            DoEvents
                            Unload wskConnexions(NoCtrl)
                        Else
                            wskConnexions(NoCtrl).SendData "Machine" & Chr(0) & _
                                                            strNomMachine & Chr(0) & _
                                                            strNomUser & Chr(0) & Chr(1)
                            DoEvents
                            ' trace
                            Temp = "Nouvelle connexion (" & lstPostesCache.List(r) & ")"
                                Status.Panels(1).Text = Temp
                                Call Ecrit_Log(Temp)
                            DoEvents
                            Connexions(i).iNoControl = NoCtrl
                            Connexions(i).strPoste = lstPostesCache.List(r)
                            Connexions(i).strPseudo = lstPostesCache.List(r)
                            Connexions(i).strIP = strIP
                            Call EcritMessage("Système", """" & lstPostesCache.List(r) & """" & _
                                " vient de se connecter (Port " & wskConnexions(NoCtrl).LocalPort & ").")
                            If lstPostes.ListItems.Count = 0 Then
                                lstPostes.ListItems.Add 1, , "<Tous>", , "Noir"
                                lstPostes.ListItems.Item(1).Checked = True
                            End If
                            lstPostes.ListItems.Add , , lstPostesCache.List(r), , "Rouge"
                            On Error Resume Next
                                If iAvecSonArrivée Then Ret = sndPlaySound(strSonArrivée, 1)
                            On Error GoTo 0
                        End If
                    End If
                    Exit Do
                End If
                i = i + 1
            Loop
        End If
    Next r
    If lstPostes.ListItems.Count = 1 Then
       If lstPostes.ListItems.Item(1) = "<Tous>" Then lstPostes.ListItems.Remove 1
    End If
    
Suite_fin:
    ' point vert
    If lstPostes.ListItems.Count > 0 Then
        lstPostes.ListItems.Item(1).SmallIcon = "Noir"
        lstPostes.Refresh
    End If
    
    Exit Sub
    
End Sub

Private Sub ChercheNode(Node As MSComctlLib.Node, Combo As Control)
' Développe tous les noeuds, insère les noms des objects du type "server" dans le "Combo" désigné

Dim FSO As Scripting.FileSystemObject
Dim NWFolder As Scripting.Folder
Dim FilX As Scripting.File, DirX As Scripting.Folder
Dim tNod As Node, isFSFolder As Boolean
Dim Temp As String, Liste As String

Screen.MousePointer = vbHourglass

' Enlève le node Fake utilisé pour faire apparaitre le signe "+"
On Error Resume Next
tvwReseau.Nodes.Remove Node.Key + "_FAKE"

' If this node is marked as a share is it a proper networked directory?
' need to make this check since NDS marks some containers (wrongly, in my opinion) as shares when they're not applicable to
' file system directories (i.e. the two containers demarking NDS and Novell FileServers are marked as shares)

' ########## On ne prend pas les données en dessous de "server" (poste)
If Not (Node.SelectedImage = "server" Or _
        Node.SelectedImage = "directory" Or _
        Node.SelectedImage = "adminshare" Or _
        Node.SelectedImage = "folder" Or _
       (Node.SelectedImage = "share" And isFSFolder = True)) Then
    ' Search up through the tree, noting the node keys so that we can then locate the
    '   NetResource object under NetRoot.
    Dim pS As String, kPath() As String, nX As NetResource, i As Integer, tX As NetResource
    Set tNod = Node ' Start at the node that was expanded
    Do While Not tNod.Parent Is Nothing ' Proceed up the tree using parent references, each time saving the node key to the string pS
        pS = tNod.Key + "|" + pS
        Set tNod = tNod.Parent
    Loop
    ' String pS is now of the form "<Node Key>|<Node Key>|<Node Key>"
    ' Split this into an array using the VB6 Split function
    kPath = Split(pS, "|")
    Set nX = NetRoot
    ' Now loop through this array, this time following down the tree of NetResource objects from
    '   NetRoot to the child NetResource object that corresponds to the node the user clicked
    For i = 0 To UBound(kPath) - 1
        Set nX = nX.Children(kPath(i))
    Next
    ' Now that we know both the node and the corresponding NetResource we can enumerate the children
    ' and add the nodes
    For Each tX In nX.Children
        Set tNod = tvwReseau.Nodes.Add(nX.RemoteName, _
                                        tvwChild, _
                                        tX.RemoteName, _
                                        tX.ShortName, _
                                        LCase(tX.ResourceTypeName), LCase(tX.ResourceTypeName))
        tNod.Tag = "N"
        ' Ne retient que les Child recherchés
        If LCase(tX.ResourceTypeName) = "server" And tX.ShortName <> "FAKE" Then
            Temp = LCase(tX.ShortName)
            ' remplace les espaces par des tirets
            'While InStr(1, Temp, " ") <> 0
            '    Mid(Temp, InStr(1, Temp, " "), 1) = "-"
            '    DoEvents
            'Wend
            ' Met en majuscule la première lettre
            If Mid(Temp, 1, 1) Like "[a-z]" Then Mid(Temp, 1, 1) = Chr(Asc(Mid(Temp, 1, 1)) - 32)
            ' Insère le nom (si ce n'est pas moi)
            If Temp <> strNomMachine Then Combo.AddItem Temp
        End If
        
        ' Add fake nodes to all new nodes except when they're printers
        ' (you can always be sure a printer never has children)
        If tX.ResourceType <> Printer Then tvwReseau.Nodes.Add tX.RemoteName, _
                                                                   tvwChild, _
                                                                   tX.RemoteName + "_FAKE", _
                                                                   "FAKE", _
                                                                   "server", "server"
    Next
    tvwReseau.Refresh  ' Refresh the view
    Node.Tag = "Y"  ' Set the tag to "Y" to denote that this node has been expanded and populated
End If

Screen.MousePointer = vbDefault
End Sub

Private Sub InitTreeView()
    
    Dim nX As NetResource, nodX As Node
    
    ' Efface le contenu actuel
    tvwReseau.Nodes.Clear
    
    ' Ajoute un premier node
    Set nodX = tvwReseau.Nodes.Add(, , "_ROOT", "Réseau global", "root", "root")
    
    ' On signale qu'il a été exploré (fait juste après)
    nodX.Tag = "Y"
    
    ' Explore le premier node "Réseau global"
    For Each nX In NetRoot.Children
        Set nodX = tvwReseau.Nodes.Add("_ROOT", _
                                        tvwChild, _
                                        nX.RemoteName, _
                                        nX.ShortName, _
                                        LCase(nX.ResourceTypeName), _
                                        LCase(nX.ResourceTypeName))
        
        ' Signale nouveau node non exploré
        nodX.Tag = "N"
        
        ' Créé un node bidon pour obtenir le signe "+"
        tvwReseau.Nodes.Add nodX.Key, _
                             tvwChild, _
                             nodX.Key + "_FAKE", _
                             "FAKE", _
                             "server", _
                             "server"
        nodX.EnsureVisible
    Next
    
End Sub

Private Sub NomMachine()
    
    On Error Resume Next
    
    'Créé un buffer
    strNomUser = String(255, Chr$(0))
    'Cherche le nom d'utiisateur
    GetUsername strNomUser, 255
    'Isole le nom dans le buffer
    strNomUser = Left$(strNomUser, InStr(strNomUser, Chr$(0)) - 1)
    
    'Créé un buffer
    strNomMachine = String(255, Chr$(0))
    GetComputerName strNomMachine, 255
    'Isole le nom dans le buffer
    strNomMachine = LCase(Left$(strNomMachine, InStr(1, strNomMachine, Chr$(0)) - 1))
    'Passe la première lettre en majuscule (si entre a et z)
    If Left(strNomMachine, 1) Like "[a-z]" Then _
        Mid(strNomMachine, 1, 1) = Chr(Asc(Mid(strNomMachine, 1, 1)) - 32)

End Sub

Private Function IPduPoste(ByVal HostName As String) As String

    Dim hFile As Long, lpWSAdata As WSAdata
    Dim hHostent As Hostent, AddrList As Long
    Dim Address As Long, rIP As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    Const SOCKET_ERROR = 0

    Call WSAStartup(&H101, lpWSAdata)
    
    If GetHostByName(HostName + String(64 - Len(HostName), 0)) <> SOCKET_ERROR Then
        CopyMemory hHostent.h_name, ByVal GetHostByName(HostName + String(64 - Len(HostName), 0)), Len(hHostent)
        CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
        CopyMemory Address, ByVal AddrList, 4
    End If
    
    hFile = IcmpCreateFile()
    
    If hFile = 0 Then
        IPduPoste = "Erreur"
        Exit Function
    End If
    OptInfo.TTL = 255
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
        rIP = CStr(EchoReply.Address(0)) + "." + _
              CStr(EchoReply.Address(1)) + "." + _
              CStr(EchoReply.Address(2)) + "." + _
              CStr(EchoReply.Address(3))
    Else
        IPduPoste = "TimeOut"
    End If
    If EchoReply.Status = 0 Then
        IPduPoste = rIP
        'MsgBox "Reply from " + HostName + " (" + rIP + ") recieved after " + Trim$(CStr(EchoReply.RoundTripTime)) + "ms"
    Else
        IPduPoste = "Erreur"
    End If
    Call IcmpCloseHandle(hFile)
    Call WSACleanup

End Function

Private Sub EcritMessage(ByVal User As String, ByVal Texte As String)

    ' On affiche le texte reçu ou envoyé dans le RichText en utilisant
    ' des couleurs et des fontes personnalisées
    
    With lstMsgRecus
        ' On se place à la fin
        .SelStart = Len(.Text)
        
        '--- Heure en vert
        .SelBold = AffHeure.bBold
        .SelItalic = AffHeure.bItalic
        .SelFontName = AffHeure.strFonte
        .SelFontSize = AffHeure.dTaille
        .SelColor = AffHeure.lCouleur   ' &H8000&
        .SelText = Format(Now, "\[hh:nn:ss\]") & " "
        
        Select Case User
            Case "Système"
                '--- Tout en violet
                .SelBold = AffSystem.bBold
                .SelItalic = AffSystem.bItalic
                .SelFontName = AffSystem.strFonte
                .SelFontSize = AffSystem.dTaille
                .SelColor = AffSystem.lCouleur  ' &H80000012
                .SelText = Texte & vbCrLf
            Case strNomUser
                '--- Nom utilisateur (moi)
                .SelBold = AffPseudo.bBold
                .SelItalic = AffPseudo.bItalic
                .SelFontName = AffPseudo.strFonte
                .SelFontSize = AffPseudo.dTaille
                .SelColor = AffPseudo.lCouleur   ' &H80FF&
                .SelText = User & " > "
                '--- Texte en orangé
                .SelBold = AffPerso.bBold
                .SelItalic = AffPerso.bItalic
                .SelFontName = AffPerso.strFonte
                .SelFontSize = AffPerso.dTaille
                .SelColor = AffPerso.lCouleur   ' &H80FF&
                .SelText = Texte & vbCrLf
            Case Else
                '--- Nom utilisateur
                .SelBold = AffPseudo.bBold
                .SelItalic = AffPseudo.bItalic
                .SelFontName = AffPseudo.strFonte
                .SelFontSize = AffPseudo.dTaille
                .SelColor = AffPseudo.lCouleur   ' &H80FF&
                .SelText = User & " > "
                ' Cas d'un message privé
                If Left(User, Len(strNomUser)) = strNomUser Then
                    '--- Texte en orangé
                    .SelBold = AffPerso.bBold
                    .SelItalic = AffPerso.bItalic
                    .SelFontName = AffPerso.strFonte
                    .SelFontSize = AffPerso.dTaille
                    .SelColor = AffPerso.lCouleur   ' &H80FF&
                    .SelText = Texte & vbCrLf
                Else
                    '--- Texte en noir
                    .SelBold = AffAutres.bBold
                    .SelItalic = AffAutres.bItalic
                    .SelFontName = AffAutres.strFonte
                    .SelFontSize = AffAutres.dTaille
                    .SelColor = AffAutres.lCouleur  ' vbBlack
                    .SelText = Texte & vbCrLf
                End If
        End Select
    End With
    
End Sub

Private Sub Nettoie_Fichier_Log()

    ' Permet de supprimer du fichier LOG les lignes qui sont trop vieilles
    ' de plus de 'iPurgeLog' jours
    ' Cette routine est lancée à chaque lancement

    Dim Date_lue As Date
    Dim ff1 As Integer, ff2 As Integer
    Dim Log_New As String
    Dim Temp As String
    Dim r As Integer
    
    ' Pas de purge si Délai est 999
    If iPurgeLog = 999 Then GoTo Nettoie_fin_2
    
    On Error GoTo Nettoie_Erreur
    Screen.MousePointer = vbHourglass
    
    ' Créé un nouveau fichier LOG dans lequel on ne gardera que les lignes utiles
    Log_New = Left(LogFile, InStrRev(LogFile, ".")) + "TMP"
    ff1 = FreeFile
    Open Log_New For Output As #ff1
    
    ' Ouvre le LogFile en Lecture
    ff2 = FreeFile
    Open LogFile For Input As #ff2

    ' Boucle de lecture
    Do While Not EOF(ff2)
        Line Input #ff2, Temp
        If Temp = "" Then GoTo fin_loop
        If Left(Temp, 8) = "        " Then GoTo fin_loop
        Date_lue = DateValue(Left(Temp, 8)) ' Format 11/12/99
        Date_lue = Format(Date_lue, "dd/mm/yy")
        ' ecrit la ligne si nombre de jours inférieur à seuil
        r = DateDiff("d", Date_lue, Now)
        If r <= iPurgeLog Then Print #1, Temp
fin_loop:
    Loop
    Close #ff1
    Close #ff2

    On Error Resume Next
    ' Supprime l'original
    If Sys.FileExists(LogFile) Then Sys.DeleteFile (LogFile)
    ' Renomme le nouveau en logfile
    Sys.CopyFile Log_New, LogFile, True

Nettoie_fin:
    ' Supprime le fichier temporaire
    If Sys.FileExists(Log_New) Then Sys.DeleteFile (Log_New)
Nettoie_fin_2:
    Screen.MousePointer = vbDefault
    Exit Sub

Nettoie_Erreur:
    On Error Resume Next
    Close #ff1
    Close #ff2
    Resume Nettoie_fin

End Sub

