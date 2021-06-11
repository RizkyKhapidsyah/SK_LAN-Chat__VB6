VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "À propos de ..."
   ClientHeight    =   3990
   ClientLeft      =   8580
   ClientTop       =   8760
   ClientWidth     =   5655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Tag             =   "À propos de TransfertDoc"
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4215
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   3240
      Width           =   1260
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   2280
      Left            =   120
      Picture         =   "frmAbout.frx":164A
      ScaleHeight     =   2220
      ScaleMode       =   0  'User
      ScaleWidth      =   600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Licence accordée à :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Tag             =   "Licence accordée à"
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblLicenseTo 
      Caption         =   "Informations licence"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Tag             =   "Licence accordée à"
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   135
      X2              =   5495
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   135
      X2              =   5495
      Y1              =   2805
      Y2              =   2805
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Avertissement: ..."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   120
      TabIndex        =   5
      Tag             =   "Avertissement: ..."
      Top             =   2880
      Width           =   3960
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Version"
      Height          =   225
      Left            =   1020
      TabIndex        =   4
      Tag             =   "Version"
      Top             =   1110
      Width           =   4485
   End
   Begin VB.Label lblTitle 
      Caption         =   "Titre de l'application"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   900
      TabIndex        =   3
      Tag             =   "Titre de l'application"
      Top             =   660
      Width           =   4605
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "Description de l'application"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   900
      TabIndex        =   2
      Tag             =   "Description de l'application"
      Top             =   1455
      Width           =   4605
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
                           ByVal lpOperation As String, ByVal lpFile As String, _
                           ByVal lpParameters As String, ByVal lpDirectory As String, _
                           ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1

Private Sub Form_Load()

    Dim r As Double
    
    ' Positionnement de la forme About au milieu de la Main
    Me.Caption = "A propos de ... " & App.Title
    r = Forme.Top + (Forme.Height / 2) - (Me.Height / 2)
    If r < 0 Then r = 0
    Me.Top = r
    r = Forme.Left + (Forme.Width / 2) - (Me.Width / 2)
    If r < 0 Then r = 0
    Me.Left = r

    'lblLicenseTo.Caption = Licence_a & " - Service " & NomServeur
    lblVersion.Caption = "32 bits - Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title & " - " & App.FileDescription
    lblDescription.Caption = App.Comments
    lblDisclaimer.Caption = AppSite & vbCr & AppAdresse

End Sub

Private Sub cmdOK_Click()
        Unload Me
End Sub

Private Sub EnvoiMail(Optional Adresse As String, _
    Optional Sujet As String, Optional Contenu As String, _
    Optional CC As String, Optional CCC As String)

    Dim Temp As String
    
    ' Créé la chaîne de commande avec les paramètres fournis
    If Len(Sujet) Then Temp = "&Subject=" & Sujet
    If Len(Contenu) Then Temp = Temp & "&Body=" & Contenu
    If Len(CC) Then Temp = Temp & "&CC=" & CC
    If Len(CCC) Then Temp = Temp & "&BCC=" & CCC
    
    'Remplace le premier '&' (s'il existe) par un '?'
    If Mid(Temp, 1, 1) = "&" Then Mid(Temp, 1, 1) = "?"
    
    'Ajoute la commande 'mailto:' et l'adresse
    Temp = "mailto:" & Adresse & Temp
    
    'Execute la commande par l'API
    Call ShellExecute(Me.hwnd, "open", Temp, _
        vbNullString, vbNullString, SW_SHOWNORMAL)

End Sub

Private Sub lblMail_Click()

    Dim Temp As String
    
    Temp = "A propos de " & App.Title & ", version " & App.Major & "." & App.Minor & "." & App.Revision
    Call EnvoiMail(Temp, "Ce programme est formidable !")
        
End Sub
