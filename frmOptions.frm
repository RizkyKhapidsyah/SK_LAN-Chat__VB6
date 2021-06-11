VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   7770
   ClientLeft      =   10890
   ClientTop       =   6795
   ClientWidth     =   7455
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Tag             =   "Options"
   Begin VB.Frame Frame1 
      Caption         =   "Identification "
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   7215
      Begin VB.TextBox txtPseudo 
         Height          =   285
         Left            =   2760
         TabIndex        =   25
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtMachine 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   24
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Votre Pseudo"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   615
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nom de votre machine"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   255
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Son à jouer"
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Width           =   7215
      Begin VB.CheckBox chkAvecSonArrivée 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "Avec ou Sans ce son"
         Top             =   1140
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox txtSonArrivée 
         Height          =   285
         Left            =   480
         TabIndex        =   40
         Top             =   1125
         Width           =   5775
      End
      Begin VB.CommandButton btnTestSonArrivée 
         Caption         =   "&Test"
         Height          =   375
         Left            =   6360
         TabIndex        =   39
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox chkAvecSonMessage 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Avec ou Sans ce son"
         Top             =   540
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox txtSonMessage 
         Height          =   285
         Left            =   480
         TabIndex        =   37
         Top             =   525
         Width           =   5775
      End
      Begin VB.CommandButton btnTestSonMessage 
         Caption         =   "&Test"
         Height          =   375
         Left            =   6360
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "lors de la connexion d'un nouveau contact"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Width           =   6975
      End
      Begin VB.Label Label2 
         Caption         =   "lors de l'arrivée d'un nouveau message alors que le Chat est en icône"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Couleurs et Polices "
      Height          =   2415
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   7215
      Begin VB.CommandButton btnInitPolice 
         Caption         =   "&Valeurs par défaut"
         Height          =   375
         Left            =   4920
         TabIndex        =   36
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton btnSystemePolice 
         Caption         =   "Messages &Système"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton btnMoiPolice 
         Caption         =   "Messages &Personnels"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton btnAutrePolice 
         Caption         =   "Messages &Reçus"
         Height          =   375
         Left            =   2520
         TabIndex        =   33
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton btnHeurePolice 
         Caption         =   "&Heure"
         Height          =   375
         Left            =   2520
         TabIndex        =   32
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton btnPseudoPolice 
         Caption         =   "Ps&eudos"
         Height          =   375
         Left            =   4920
         TabIndex        =   31
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton btnSystemeCouleur 
         Height          =   375
         Left            =   2160
         Picture         =   "frmOptions.frx":164A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton btnMoiCouleur 
         Height          =   375
         Left            =   2160
         Picture         =   "frmOptions.frx":1794
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton btnAutreCouleur 
         Height          =   375
         Left            =   4440
         Picture         =   "frmOptions.frx":18DE
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton btnHeureCouleur 
         Height          =   375
         Left            =   4440
         Picture         =   "frmOptions.frx":1A28
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton btnPseudoCouleur 
         Height          =   375
         Left            =   6840
         Picture         =   "frmOptions.frx":1B72
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1440
         Width           =   255
      End
      Begin RichTextLib.RichTextBox txtExemples 
         Height          =   1095
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1931
         _Version        =   393217
         TextRTF         =   $"frmOptions.frx":1CBC
      End
   End
   Begin MSComDlg.CommonDialog dialPoliceCouleur 
      Left            =   480
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Périodicité "
      Height          =   1965
      Left            =   120
      TabIndex        =   8
      Tag             =   "Exemple 1"
      Top             =   1080
      Width           =   7185
      Begin VB.TextBox txtRafraichissement 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4920
         TabIndex        =   15
         Text            =   "60"
         ToolTipText     =   "Valeur entre 1 et 60 minutes."
         Top             =   555
         Width           =   495
      End
      Begin VB.TextBox txtLog_Purge 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4920
         TabIndex        =   9
         Text            =   "7"
         ToolTipText     =   $"frmOptions.frx":1D85
         Top             =   195
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "minutes"
         Height          =   195
         Index           =   15
         Left            =   5520
         TabIndex        =   16
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Nota :"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblNota 
         Caption         =   "Il n'est pas nécessaire ..."
         Height          =   975
         Left            =   840
         TabIndex        =   13
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "jours d'ancienneté."
         Height          =   195
         Index           =   12
         Left            =   5490
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Supprimer les lignes du fichier de LOG qui ont plus de"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Faire la recherche des nouveaux postes sur le réseau toutes les"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   390
      Left            =   2520
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   7305
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Annuler"
      Height          =   390
      Left            =   3840
      TabIndex        =   1
      Tag             =   "Annuler"
      Top             =   7305
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Exemple 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   7
         Tag             =   "Exemple 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Exemple 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   6
         Tag             =   "Exemple 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Exemple 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   4
         Tag             =   "Exemple 2"
         Top             =   305
         Width           =   2033
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strAncienPseudo As String

Private Sub btnAutreCouleur_Click()

    With dialPoliceCouleur
        .Flags = cdlCCFullOpen
        .Color = AffAutres.lCouleur
        .ShowColor
        AffAutres.lCouleur = .Color
    End With
    
    ' Remet à jour l'affichage avec ces nouveaux paramètres
    Call TestCouleurPolice

End Sub

Private Sub btnAutrePolice_Click()

    With dialPoliceCouleur
        .Flags = cdlCFTTOnly + cdlCFScreenFonts   ' cdlCFBoth +
        .FontName = AffAutres.strFonte
        .FontSize = AffAutres.dTaille
        .FontBold = AffAutres.bBold
        .FontItalic = AffAutres.bItalic
        .Color = AffAutres.lCouleur
        .DialogTitle = App.Title & " - Affichage de l'Heure"
            .ShowFont
        If .FontName <> "" Then AffAutres.strFonte = .FontName
        If .FontSize <> 0 Then AffAutres.dTaille = .FontSize
        AffAutres.lCouleur = .Color
        AffAutres.bBold = .FontBold
        AffAutres.bItalic = .FontItalic
    End With
    
    ' Remet à jour l'affichage avec ces nouveaux paramètres
    Call TestCouleurPolice

End Sub

Private Sub btnHeureCouleur_Click()

    With dialPoliceCouleur
        .Flags = cdlCCFullOpen
        .Color = AffHeure.lCouleur
        .ShowColor
        AffHeure.lCouleur = .Color
    End With
    
    ' Remet à jour l'affichage avec ces nouveaux paramètres
    Call TestCouleurPolice

End Sub

Private Sub btnHeurePolice_Click()

    With dialPoliceCouleur
        .Flags = cdlCFTTOnly + cdlCFScreenFonts   ' cdlCFBoth +
        .FontName = AffHeure.strFonte
        .FontSize = AffHeure.dTaille
        .FontBold = AffHeure.bBold
        .FontItalic = AffHeure.bItalic
        .Color = AffHeure.lCouleur
        .DialogTitle = App.Title & " - Affichage de l'Heure"
            .ShowFont
        If .FontName <> "" Then AffHeure.strFonte = .FontName
        If .FontSize <> 0 Then AffHeure.dTaille = .FontSize
        AffHeure.lCouleur = .Color
        AffHeure.bBold = .FontBold
        AffHeure.bItalic = .FontItalic
    End With
    
    ' Remet à jour l'affichage avec ces nouveaux paramètres
    Call TestCouleurPolice

End Sub

Private Sub btnInitPolice_Click()

    ' Paramètres par défaut (que j'aime bien)
        
    ' paramètres Heure
    AffHeure.lCouleur = &H8000&
    AffHeure.strFonte = "Arial"
    AffHeure.dTaille = 8
    AffHeure.bBold = False
    AffHeure.bItalic = False
    ' paramètres Pseudo
    AffPseudo.lCouleur = &HC00000
    AffPseudo.strFonte = "Arial"
    AffPseudo.dTaille = 8
    AffPseudo.bBold = False
    AffPseudo.bItalic = False
    ' paramètres Système
    AffSystem.lCouleur = &H80000012
    AffSystem.strFonte = "Arial"
    AffSystem.dTaille = 8
    AffSystem.bBold = False
    AffSystem.bItalic = True
    ' paramètres Perso
    AffPerso.lCouleur = &H80FF&
    AffPerso.strFonte = "Arial"
    AffPerso.dTaille = 8
    AffPerso.bBold = False
    AffPerso.bItalic = False
    ' paramètres Autres
    AffAutres.lCouleur = vbBlack
    AffAutres.strFonte = "Arial"
    AffAutres.dTaille = 8
    AffAutres.bBold = False
    AffAutres.bItalic = False

    ' Remet à jour l'affichage avec ces nouveaux paramètres
    Call TestCouleurPolice

End Sub

Private Sub btnMoiCouleur_Click()

    With dialPoliceCouleur
        .Flags = cdlCCFullOpen
        .Color = AffPerso.lCouleur
        .ShowColor
        AffPerso.lCouleur = .Color
    End With
    
    ' Remet à jour l'affichage avec ces nouveaux paramètres
    Call TestCouleurPolice

End Sub

Private Sub btnMoiPolice_Click()

    With dialPoliceCouleur
        .Flags = cdlCFTTOnly + cdlCFScreenFonts   ' cdlCFBoth +
        .FontName = AffPerso.strFonte
        .FontSize = AffPerso.dTaille
        .FontBold = AffPerso.bBold
        .FontItalic = AffPerso.bItalic
        .Color = AffPerso.lCouleur
        .DialogTitle = App.Title & " - Affichage de l'Heure"
            .ShowFont
        If .FontName <> "" Then AffPerso.strFonte = .FontName
        If .FontSize <> 0 Then AffPerso.dTaille = .FontSize
        AffPerso.lCouleur = .Color
        AffPerso.bBold = .FontBold
        AffPerso.bItalic = .FontItalic
    End With
    
    ' Remet à jour l'affichage avec ces nouveaux paramètres
    Call TestCouleurPolice

End Sub

Private Sub btnPseudoCouleur_Click()

    With dialPoliceCouleur
        .Flags = cdlCCFullOpen
        .Color = AffPseudo.lCouleur
        .ShowColor
        AffPseudo.lCouleur = .Color
    End With
    
    ' Remet à jour l'affichage avec ces nouveaux paramètres
    Call TestCouleurPolice

End Sub

Private Sub btnPseudoPolice_Click()

    With dialPoliceCouleur
        .Flags = cdlCFTTOnly + cdlCFScreenFonts   ' cdlCFBoth +
        .FontName = AffPseudo.strFonte
        .FontSize = AffPseudo.dTaille
        .FontBold = AffPseudo.bBold
        .FontItalic = AffPseudo.bItalic
        .Color = AffPseudo.lCouleur
        .DialogTitle = App.Title & " - Affichage de l'Heure"
            .ShowFont
        If .FontName <> "" Then AffPseudo.strFonte = .FontName
        If .FontSize <> 0 Then AffPseudo.dTaille = .FontSize
        AffPseudo.lCouleur = .Color
        AffPseudo.bBold = .FontBold
        AffPseudo.bItalic = .FontItalic
    End With
    
    ' Remet à jour l'affichage avec ces nouveaux paramètres
    Call TestCouleurPolice

End Sub

Private Sub btnSystemeCouleur_Click()

    With dialPoliceCouleur
        .Flags = cdlCCFullOpen
        .Color = AffSystem.lCouleur
        .ShowColor
        AffSystem.lCouleur = .Color
    End With
    
    ' Remet à jour l'affichage avec ces nouveaux paramètres
    Call TestCouleurPolice

End Sub

Private Sub btnSystemePolice_Click()

    With dialPoliceCouleur
        .Flags = cdlCFTTOnly + cdlCFScreenFonts   ' cdlCFBoth +
        .FontName = AffSystem.strFonte
        .FontSize = AffSystem.dTaille
        .FontBold = AffSystem.bBold
        .FontItalic = AffSystem.bItalic
        .Color = AffSystem.lCouleur
        .DialogTitle = App.Title & " - Affichage de l'Heure"
            .ShowFont
        If .FontName <> "" Then AffSystem.strFonte = .FontName
        If .FontSize <> 0 Then AffSystem.dTaille = .FontSize
        AffSystem.lCouleur = .Color
        AffSystem.bBold = .FontBold
        AffSystem.bItalic = .FontItalic
    End With
    
    ' Remet à jour l'affichage avec ces nouveaux paramètres
    Call TestCouleurPolice

End Sub

Private Sub btnTestSon_Click()
End Sub

Private Sub btnTestSonArrivée_Click()

    Dim Ret As Long
    
    On Error Resume Next
    If bCarteSon Then Ret = sndPlaySound(txtSonArrivée.Text, 1)

End Sub

Private Sub btnTestSonMessage_Click()

    Dim Ret As Long
    
    On Error Resume Next
    If bCarteSon Then Ret = sndPlaySound(txtSonMessage.Text, 1)

End Sub

Private Sub chkAvecSonArrivée_Click()

    If chkAvecSonArrivée.Value = vbGrayed Then chkAvecSonArrivée.Value = vbUnchecked

End Sub

Private Sub chkAvecSonMessage_Click()

    If chkAvecSonMessage.Value = vbGrayed Then chkAvecSonMessage.Value = vbUnchecked

End Sub

Private Sub cmdCancel_Click()
    
    Unload Me

End Sub

Private Sub cmdOK_Click()
        
    Dim r As Integer, i As Integer, Temp As String
    
    Screen.MousePointer = vbHourglass
    '---- Vérifie les valeurs min/max
    r = Val(txtRafraichissement.Text)
        If r < 1 Then txtRafraichissement.Text = "1"
        If r > 60 Then txtRafraichissement.Text = "60"
    iPeriodeRafr = Val(txtRafraichissement.Text)
    
    r = Val(txtLog_Purge.Text)
        If r < 1 Then txtLog_Purge.Text = "1"
        If r > 10 And r <> 999 Then txtLog_Purge.Text = "10"
    Me.Refresh
    iPurgeLog = Val(txtLog_Purge.Text)
    
    '---- Ecrit ces valeurs dans la base de registres ------------------------------------------------
    SaveSetting AppSite, App.Title, "Rafraichissement postes (min)", txtRafraichissement.Text
    SaveSetting AppSite, App.Title, "Purge log (jours)", txtLog_Purge.Text
    
    ' Son arrivée nouveau message
    If Sys.FileExists(txtSonMessage.Text) Then
        SaveSetting AppSite, App.Title, "Son Wave (nouveau message)", txtSonMessage.Text
        strSonMessage = txtSonMessage.Text
    Else
        If Sys.FileExists(App.Path & "\Message.wav") Then
            SaveSetting AppSite, App.Title, "Son Wave (nouveau message)", App.Path & "\Message.wav"
            strSonMessage = App.Path & "\Message.wav"
        Else
            chkAvecSonMessage.Value = vbUnchecked
        End If
    End If
    SaveSetting AppSite, App.Title, "Son Wave (nouveau message) Avec", chkAvecSonMessage.Value
    iAvecSonMessage = chkAvecSonMessage.Value
    
    ' Son arrivée nouveau contact
    If Sys.FileExists(txtSonArrivée.Text) Then
        SaveSetting AppSite, App.Title, "Son Wave (nouveau contact)", txtSonArrivée.Text
        strSonArrivée = txtSonArrivée.Text
    Else
        If Sys.FileExists(App.Path & "\Arrivée.wav") Then
            SaveSetting AppSite, App.Title, "Son Wave (nouveau contact)", App.Path & "\Arrivée.wav"
            strSonArrivée = App.Path & "\Arrivée.wav"
        Else
            chkAvecSonArrivée.Value = vbUnchecked
        End If
    End If
    SaveSetting AppSite, App.Title, "Son Wave (nouveau contact) Avec", chkAvecSonArrivée.Value
    iAvecSonArrivée = chkAvecSonArrivée.Value
    
    ' paramètres Heure
    SaveSetting AppSite, App.Title, "Affichage Heure (couleur)", AffHeure.lCouleur
    SaveSetting AppSite, App.Title, "Affichage Heure (fonte)", AffHeure.strFonte
    SaveSetting AppSite, App.Title, "Affichage Heure (taille)", AffHeure.dTaille
    SaveSetting AppSite, App.Title, "Affichage Heure (gras)", AffHeure.bBold
    SaveSetting AppSite, App.Title, "Affichage Heure (italique)", AffHeure.bItalic
    ' paramètres Pseudo
    SaveSetting AppSite, App.Title, "Affichage Pseudo (couleur)", AffPseudo.lCouleur
    SaveSetting AppSite, App.Title, "Affichage Pseudo (fonte)", AffPseudo.strFonte
    SaveSetting AppSite, App.Title, "Affichage Pseudo (taille)", AffPseudo.dTaille
    SaveSetting AppSite, App.Title, "Affichage Pseudo (gras)", AffPseudo.bBold
    SaveSetting AppSite, App.Title, "Affichage Pseudo (italique)", AffPseudo.bItalic
    ' paramètres Système
    SaveSetting AppSite, App.Title, "Affichage System (couleur)", AffSystem.lCouleur
    SaveSetting AppSite, App.Title, "Affichage System (fonte)", AffSystem.strFonte
    SaveSetting AppSite, App.Title, "Affichage System (taille)", AffSystem.dTaille
    SaveSetting AppSite, App.Title, "Affichage System (gras)", AffSystem.bBold
    SaveSetting AppSite, App.Title, "Affichage System (italique)", AffSystem.bItalic
    ' paramètres Perso
    SaveSetting AppSite, App.Title, "Affichage Perso (couleur)", AffPerso.lCouleur
    SaveSetting AppSite, App.Title, "Affichage Perso (fonte)", AffPerso.strFonte
    SaveSetting AppSite, App.Title, "Affichage Perso (taille)", AffPerso.dTaille
    SaveSetting AppSite, App.Title, "Affichage Perso (gras)", AffPerso.bBold
    SaveSetting AppSite, App.Title, "Affichage Perso (italique)", AffPerso.bItalic
    ' paramètres Autres
    SaveSetting AppSite, App.Title, "Affichage Autre (couleur)", AffAutres.lCouleur
    SaveSetting AppSite, App.Title, "Affichage Autre (fonte)", AffAutres.strFonte
    SaveSetting AppSite, App.Title, "Affichage Autre (taille)", AffAutres.dTaille
    SaveSetting AppSite, App.Title, "Affichage Autre (gras)", AffAutres.bBold
    SaveSetting AppSite, App.Title, "Affichage Autre (italique)", AffAutres.bItalic
        
    ' Modif du Pseudo
    SaveSetting AppSite, App.Title, "Pseudo", txtPseudo.Text
    strNomUser = txtPseudo.Text
    
    ' Termine si le pseudo n'a pas changé ou s'il n'y a personne à avertir
    If strAncienPseudo = strNomUser Or _
        Forme.lstPostes.ListItems.Count = 0 Then GoTo Fin_Ok
    ' Le pseudo à changé : il faut informer tout le monde
    r = 2
    Do While r <= Forme.lstPostes.ListItems.Count
        i = 1
        Do While i <= 200
            If Forme.lstPostes.ListItems.Item(r) = Connexions(i).strIP Or _
               Forme.lstPostes.ListItems.Item(r) = Connexions(i).strPseudo Then
               Forme.wskConnexions(Connexions(i).iNoControl).SendData _
                    "Machine" & Chr(0) & strNomMachine & Chr(0) & strNomUser & Chr(0) & Chr(1)
            End If
            i = i + 1
        Loop
        r = r + 1
        DoEvents
    Loop
    
Fin_Ok:
    Screen.MousePointer = vbDefault
    Unload Me
 
 End Sub

Private Sub Form_Load()
    
    Dim Largeur As Integer, r As Integer
    Dim i As Integer
    
    ' Positionnement de la forme Options au milieu de la Main
    Me.Caption = App.Title & " - Options"
    r = Forme.Top + (Forme.Height / 2) - (Me.Height / 2)
    If r < 0 Then r = 0
    If r + Me.Height > Screen.Height Then r = Screen.Height - Me.Height - 100
    Me.Top = r
    r = Forme.Left + (Forme.Width / 2) - (Me.Width / 2)
    If r < 0 Then r = 0
    If r + Me.Width > Screen.Width Then r = Screen.Width - Me.Width - 100
    Me.Left = r

    ' Nom de la machine (non modifiable) et pseudo (modifiable)
    txtMachine.Text = strNomMachine
    txtPseudo.Text = strNomUser
    ' Mémorise le Pseudo pour le comparer lors du 'Ok'
    strAncienPseudo = strNomUser
        
    ' Nota sous la période de rafrachissement
    lblNota.Caption = _
        "Il n'est pas nécessaire d'entrer des périodes de rafraichissement trop courtes." & vbCr & _
        "En effet, lors du lancement de " & App.Title & " par un autre poste, c'est lui qui donnera signe de " & _
        "vie pour se connecter à vous." & vbCr & _
        "Ce rafraichissement n'est utile que dans le cas d'une perte de communication avec les autres"
    
    ' Son à jouer sur arrivée nouveau message
    btnTestSonMessage.Enabled = bCarteSon
    txtSonMessage.Enabled = bCarteSon
    chkAvecSonMessage.Enabled = bCarteSon
    txtSonMessage.Text = strSonMessage
    chkAvecSonMessage.Value = iAvecSonMessage
    
    ' Son à jouer sur arrivée nouvel contact
    btnTestSonArrivée.Enabled = bCarteSon
    txtSonArrivée.Enabled = bCarteSon
    chkAvecSonArrivée.Enabled = bCarteSon
    txtSonArrivée.Text = strSonArrivée
    chkAvecSonArrivée.Value = iAvecSonArrivée
    
    ' Exemple de fenêtre de message
    Call TestCouleurPolice
    
    txtLog_Purge.Text = Format(iPurgeLog, "0")
    txtRafraichissement.Text = Format(iPeriodeRafr, "0")

End Sub

Private Sub TestCouleurPolice()

    ' Exemple des couleurs utilisées dans l'affichage des messages
    With txtExemples
        ' Sélectionne tout le texte pour l'affacer
        .SelStart = 0
        .SelLength = Len(txtExemples.Text)
        
        ' Heure
        .SelBold = AffHeure.bBold
        .SelItalic = AffHeure.bItalic
        .SelFontName = AffHeure.strFonte
        .SelFontSize = AffHeure.dTaille
        .SelColor = AffHeure.lCouleur   ' &H8000&
            .SelText = Format(Now, "\[hh:nn:ss\]") & " "
        ' Message système
        .SelBold = AffSystem.bBold
        .SelItalic = AffSystem.bItalic
        .SelFontName = AffSystem.strFonte
        .SelFontSize = AffSystem.dTaille
        .SelColor = AffSystem.lCouleur  ' &H80000012
            .SelText = "Ceci est un message ""Système""" & vbCrLf

        ' Heure
        .SelBold = AffHeure.bBold
        .SelItalic = AffHeure.bItalic
        .SelFontName = AffHeure.strFonte
        .SelFontSize = AffHeure.dTaille
        .SelColor = AffHeure.lCouleur
            .SelText = Format(Now, "\[hh:nn:ss\]") & " "
        ' Message emis
        .SelBold = AffPseudo.bBold
        .SelItalic = AffPseudo.bItalic
        .SelFontName = AffPseudo.strFonte
        .SelFontSize = AffPseudo.dTaille
        .SelColor = AffPseudo.lCouleur   ' &H80FF&
            .SelText = "Jack > "
        .SelBold = AffPerso.bBold
        .SelItalic = AffPerso.bItalic
        .SelFontName = AffPerso.strFonte
        .SelFontSize = AffPerso.dTaille
        .SelColor = AffPerso.lCouleur   ' &H80FF&
            .SelText = "Ceci est un message que j'ai tapé moi-même." & vbCrLf
        
        ' Heure
        .SelBold = AffHeure.bBold
        .SelItalic = AffHeure.bItalic
        .SelFontName = AffHeure.strFonte
        .SelFontSize = AffHeure.dTaille
        .SelColor = AffHeure.lCouleur
            .SelText = Format(Now, "\[hh:nn:ss\]") & " "
        ' Message (extérieur)
        .SelBold = AffPseudo.bBold
        .SelItalic = AffPseudo.bItalic
        .SelFontName = AffPseudo.strFonte
        .SelFontSize = AffPseudo.dTaille
        .SelColor = AffPseudo.lCouleur   ' &HC00000
            .SelText = "Emile > "
        .SelBold = AffAutres.bBold
        .SelItalic = AffAutres.bItalic
        .SelFontName = AffAutres.strFonte
        .SelFontSize = AffAutres.dTaille
        .SelColor = AffAutres.lCouleur  ' vbBlack
            .SelText = "Ceci est un message tapé par un autre poste." & vbCrLf
    End With

End Sub

Private Sub txtSonArrivée_Click()

    With dialPoliceCouleur
        .DialogTitle = App.Title & " - Son à jouer lors de l'arrivée d'un nouveau Contact"
        .FileName = txtSonArrivée.Text
        .Filter = "Fichiers son (wav)|*.wav"
        .InitDir = Left(txtSonMessage.Text, InStrRev(txtSonArrivée.Text, "\") - 1)
        .ShowOpen
        txtSonArrivée.Text = .FileName
    End With

End Sub

Private Sub txtSonMessage_Click()

    With dialPoliceCouleur
        .DialogTitle = App.Title & " - Son à jouer lors de l'arrivée d'un nouveau Message"
        .FileName = txtSonMessage.Text
        .Filter = "Fichiers son (wav)|*.wav"
        .InitDir = Left(txtSonMessage.Text, InStrRev(txtSonMessage.Text, "\") - 1)
        .ShowOpen
        txtSonMessage.Text = .FileName
    End With

End Sub
