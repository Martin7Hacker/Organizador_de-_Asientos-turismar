VERSION 5.00
Begin VB.Form frmAcercade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de "
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1245
      Left            =   120
      Picture         =   "frmAcercade.frx":0000
      ScaleHeight     =   1185
      ScaleWidth      =   10155
      TabIndex        =   2
      Top             =   120
      Width           =   10215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Software desarollado por Martin Grasso."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   3450
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Organizador de Asientos turismar v1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   3450
   End
End
Attribute VB_Name = "frmAcercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
Me.Icon = frmPrograma.Icon
End Sub
