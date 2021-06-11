VERSION 5.00
Begin VB.Form frmExtraNet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connexion ExtraNet"
   ClientHeight    =   5025
   ClientLeft      =   5100
   ClientTop       =   8535
   ClientWidth     =   5940
   ControlBox      =   0   'False
   Icon            =   "frmExtraNet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5940
   Begin VB.TextBox txtAdresseIP 
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   4185
      Width           =   2055
   End
   Begin VB.CommandButton btnAnnuler 
      Cancel          =   -1  'True
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblCommentaires 
      Caption         =   $"frmExtraNet.frx":164A
      Height          =   855
      Index           =   8
      Left            =   600
      TabIndex        =   12
      Top             =   3240
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCommentaires 
      Caption         =   $"frmExtraNet.frx":1760
      Height          =   855
      Index           =   7
      Left            =   600
      TabIndex        =   11
      Top             =   2400
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCommentaires 
      Caption         =   "Nota :"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCommentaires 
      Caption         =   $"frmExtraNet.frx":185B
      Height          =   495
      Index           =   5
      Left            =   600
      TabIndex        =   9
      Top             =   1920
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCommentaires 
      Caption         =   $"frmExtraNet.frx":18EA
      Height          =   495
      Index           =   4
      Left            =   600
      TabIndex        =   8
      Top             =   1440
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCommentaires 
      Caption         =   "Il faut donc connaître l'IP actuelle de votre interlocuteur. S'il peut vous la transmettre, pas de problème."
      Height          =   495
      Index           =   3
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCommentaires 
      Caption         =   $"frmExtraNet.frx":1979
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   480
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCommentaires 
      Caption         =   "Nota :"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCommentaires 
      Caption         =   "Entrez ici l'adresse IP d'une machine se trouvant hors de votre réseau"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Adresse IP à contacter"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
End
Attribute VB_Name = "frmExtraNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAnnuler_Click()

    Unload Me

End Sub

Private Sub btnOk_Click()

    Dim Temp As String, r As Integer
    
    If txtAdresseIP.Text = "" Then
        Beep
        GoTo Sortie
    End If
    Temp = Element(txtAdresseIP.Text, 1, ".")
    r = Val(Temp)
        If Temp = "" Or Temp Like "*[!0-9]*" Or r < 0 Or r > 255 Then GoTo PasGood
    Temp = Element(txtAdresseIP.Text, 2, ".")
    r = Val(Temp)
        If Temp = "" Or Temp Like "*[!0-9]*" Or r < 0 Or r > 255 Then GoTo PasGood
    Temp = Element(txtAdresseIP.Text, 3, ".")
    r = Val(Temp)
        If Temp = "" Or Temp Like "*[!0-9]*" Or r < 0 Or r > 255 Then GoTo PasGood
    Temp = Element(txtAdresseIP.Text, 4, ".")
    r = Val(Temp)
        If Temp = "" Or Temp Like "*[!0-9]*" Or r < 0 Or r > 255 Then GoTo PasGood
    
    strIPextraNet = txtAdresseIP.Text
    
Sortie:
    Unload Me
    Exit Sub
    
PasGood:
    Temp = "Cette entrée ne correspond pas à une adresse IP valable." & vbCr & vbCr
    Temp = Temp & "Elle doit avoir la forme :" & vbCr
    Temp = Temp & "12.58.128.91  = 4 chiffres de 0 à 255 séparés par des points"
    MsgBox Temp, vbCritical And vbOKOnly, App.Title & " - ExtraNet"

End Sub

Private Sub Form_Load()

    r = Forme.Top + (Forme.Height / 2) - (Me.Height / 2)
    If r < 0 Then r = 0
    Me.Top = r
    r = Forme.Left + (Forme.Width / 2) - (Me.Width / 2)
    If r < 0 Then r = 0
    Me.Left = r

    Me.Caption = App.Title & " - Adresse IP ExtraNet"
    
End Sub

