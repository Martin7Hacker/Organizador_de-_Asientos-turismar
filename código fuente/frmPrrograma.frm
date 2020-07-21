VERSION 5.00
Begin VB.Form frmPrograma 
   Caption         =   "Organizador de Asientos turismar v1.0"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10395
   Icon            =   "frmPrrograma.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10395
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPasillo 
      BackColor       =   &H8000000C&
      Caption         =   "Pasillo"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdVentanilla 
      BackColor       =   &H00FFFF80&
      Caption         =   "Ventanilla"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAcercaDe 
      BackColor       =   &H0080C0FF&
      Caption         =   "Acerca de"
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   8160
      ScaleHeight     =   315
      ScaleWidth      =   1635
      TabIndex        =   26
      Top             =   2400
      Width           =   1695
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENTANILLA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   45
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   8760
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   23
      Top             =   1800
      Width           =   495
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   50
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   8400
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   20
      Top             =   600
      Width           =   1215
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DERECHA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   45
         Width           =   900
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   8760
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   18
      Top             =   1200
      Width           =   495
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   50
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   2775
      Left            =   7320
      ScaleHeight     =   2715
      ScaleWidth      =   0
      TabIndex        =   14
      Top             =   70
      Width           =   53
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Columns         =   10
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   4440
      TabIndex        =   13
      Top             =   600
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4320
      Top             =   2400
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   3045
      Left            =   360
      Picture         =   "frmPrrograma.frx":57E2
      ScaleHeight     =   2985
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   840
         ScaleHeight     =   315
         ScaleWidth      =   1515
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Pasillo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   50
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FF8080&
         Height          =   375
         Left            =   1560
         ScaleHeight     =   315
         ScaleWidth      =   1515
         TabIndex        =   5
         Top             =   360
         Width           =   1575
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ventanilla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   45
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2130
      Left            =   360
      Picture         =   "frmPrrograma.frx":7368
      ScaleHeight     =   2070
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   3240
      Width           =   9855
      Begin VB.PictureBox picSelector 
         BackColor       =   &H0000FF00&
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1680
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   9
         Top             =   1560
         Width           =   495
         Begin VB.Label labaciento 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   50
            Width           =   735
         End
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Derecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Isquierda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Ventanilla O Pasillo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   25
      Top             =   2160
      Width           =   2715
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero Asiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8400
      TabIndex        =   22
      Top             =   1560
      Width           =   2715
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Fila"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8400
      TabIndex        =   17
      Top             =   960
      Width           =   2715
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Fila Isquierda / Derecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7920
      TabIndex        =   16
      Top             =   360
      Width           =   2715
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos del Asiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   15
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label lablugar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLANO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4800
      TabIndex        =   4
      Top             =   2880
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Asiento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   1020
   End
End
Attribute VB_Name = "frmPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*/**************************************************************/
'*/ ORGANIZADOR DE ASIENTOS DE OMNIBUS TURISMAR GRAFICO         */
'*/ EN VISUAL BASIC                                             */
'*/ EJ DE COMO CREAR UNA INTERFAZ DE UN HORGANIZADOR DE         */
'*/ ASIENTOS PARA EL S.O WINDOWS                                */
'*/ AUTOR: MARTIN GRASSO CASTRILLO.                             */
'*/**************************************************************/
'15 / 04 / 2017 | 14:10:57 HS.
'
'
'
Dim fila As String  ' guarda el valor de la fila de Asientos
Dim cont_as As Byte ' contador de asientos
'
'
Private Sub limpiar_Click()
limpiarDatos:  crearAsientos ' se limpia y se crea nuevamente los asientos
End Sub

Private Sub cmdresolver_Click()
 tablaAsientos Combo1.ListIndex + 5: labaciento.Caption = Combo1.Text
 ventanillaOPasillo Combo1.ListIndex + 5
 Label12.Caption = Combo1.Text
 Label10.Caption = fila
End Sub

Private Sub cmdAcercaDe_Click()
 frmAcercade.Show 1
End Sub

Private Sub cmdPasillo_Click()
cont_as = cont_as + 1
Select Case cont_as
       Case 1
       Combo1.ListIndex = 2
       Case 2
       Combo1.ListIndex = 6
       Case 3
       Combo1.ListIndex = 10
       Case 4
       Combo1.ListIndex = 14
       Case 5
       Combo1.ListIndex = 18
       Case 6
       Combo1.ListIndex = 22
       Case 7
       Combo1.ListIndex = 26
       Case 8
       Combo1.ListIndex = 30
       Case 9
       Combo1.ListIndex = 34
       Case 10
       Combo1.ListIndex = 38
       Case 11
       Combo1.ListIndex = 42
       Case 12
       Combo1.ListIndex = 1
       Case 13
       Combo1.ListIndex = 5
       Case 14
       Combo1.ListIndex = 9
       Case 15
       Combo1.ListIndex = 13
       Case 16
       Combo1.ListIndex = 17
       Case 17
       Combo1.ListIndex = 21
       Case 18
       Combo1.ListIndex = 25
       Case 19
       Combo1.ListIndex = 29
       Case 20
       Combo1.ListIndex = 33
       Case 21
       Combo1.ListIndex = 37
       Case 22
       Combo1.ListIndex = 41
 End Select
 If cont_as = 22 Then
    cont_as = 0
 End If

End Sub

Private Sub cmdVentanilla_Click()
cont_as = cont_as + 1
Select Case cont_as
       Case 1
       Combo1.ListIndex = 3
       Case 2
       Combo1.ListIndex = 7
       Case 3
       Combo1.ListIndex = 11
       Case 4
       Combo1.ListIndex = 15
       Case 5
       Combo1.ListIndex = 19
       Case 6
       Combo1.ListIndex = 23
       Case 7
       Combo1.ListIndex = 27
       Case 8
       Combo1.ListIndex = 31
       Case 9
       Combo1.ListIndex = 35
       Case 10
       Combo1.ListIndex = 39
       Case 11
       Combo1.ListIndex = 43
       Case 12
       Combo1.ListIndex = 0
       Case 13
       Combo1.ListIndex = 4
       Case 14
       Combo1.ListIndex = 8
       Case 15
       Combo1.ListIndex = 12
       Case 16
       Combo1.ListIndex = 16
       Case 17
       Combo1.ListIndex = 20
       Case 18
       Combo1.ListIndex = 24
       Case 19
       Combo1.ListIndex = 28
       Case 20
       Combo1.ListIndex = 32
       Case 21
       Combo1.ListIndex = 36
       Case 22
       Combo1.ListIndex = 40
 End Select
 If cont_as = 22 Then
    cont_as = 0
 End If

End Sub

Private Sub Combo1_Change()
 cmdresolver_Click
End Sub

Private Sub Combo1_Click()
 Combo1_Scroll
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
 Combo1_Scroll
End Sub

Private Sub Combo1_Scroll()
 Combo1_Change
 List1.ListIndex = Combo1.ListIndex
End Sub

'al cargar el programa
Private Sub Form_Load()
 crearAsientos
 With picSelector
      .Top = 120
      .Left = 1560
 End With
 eliminar_pasilloyVentanilla ' elimnia los pasillos y ventanillas
 Combo1.ListIndex = 0
End Sub

'crea los asientos en el combobox
Private Sub crearAsientos()
 Const maximo As Byte = 44
 Dim as_x As Byte
  limpiarDatos
   For as_x = 1 To maximo
    Combo1.AddItem as_x
    List1.AddItem as_x
   Next as_x
End Sub

'limpia los asientos en el combobox
Private Sub limpiarDatos()
Combo1.Clear
End Sub

'verifica si es ventana o pasillo
Private Sub EsVentanillaOPasillo(ByVal lugar As Byte)
 Select Case lugar
  Case 0
   Picture4.Visible = True  ' lado de pasillo verdadero
   Picture3.Visible = False ' lado de ventanilla false
  Case 1
   Picture4.Visible = False ' lado de pasillo falso
   Picture3.Visible = True  ' lado de ventanilla verdero
 End Select
End Sub

'resuelve el lugar de Asientos de todo el bondi
Private Sub asientoContratado()
Dim rec_terminado As Boolean
  If picSelector.Left = 8860 Or picSelector.Left = _
     8930 And picSelector.Top = 120 Then
     picSelector.Top = 480
     picSelector.Left = 900
  ElseIf picSelector.Left = 8930 And picSelector.Top = 480 Then
     picSelector.Top = 1200
     picSelector.Left = 990
  ElseIf picSelector.Left = 9020 And picSelector.Top = 1200 Then
     picSelector.Top = 1560
     picSelector.Left = 990
     
   ElseIf picSelector.Left = 9020 And picSelector.Top = 1560 Then
     picSelector.Top = 120
     picSelector.Left = 900
  End If
   picSelector.Left = picSelector.Left + 730
End Sub
  '
  ' tabla de valores de cada asiento
  '
  Private Sub tablaAsientos(ByVal asiento As Byte)
  Select Case asiento
   Case 5
  With picSelector
       .Top = 1560
       .Left = 1720
       fila = 1
  End With
   Case 6
  With picSelector
       .Top = 1200
       .Left = 1720
       fila = 1
  End With
   Case 7
  With picSelector
       .Top = 480
       .Left = 1630
       fila = 1
  End With
   Case 8
  With picSelector
       .Top = 120
       .Left = 1630
       fila = 1
  End With
   Case 9
  With picSelector
       .Top = 1560
       .Left = 2450
       fila = 2
  End With
   Case 10
  With picSelector
       .Top = 1200
       .Left = 2450
       fila = 2
  End With
   Case 11
  With picSelector
       .Top = 480
       .Left = 2360
       fila = 2
  End With
   Case 12
  With picSelector
       .Top = 120
       .Left = 2360
       fila = 2
  End With
   Case 13
  With picSelector
       .Top = 1560
       .Left = 3180
       fila = 3
  End With
   Case 14
  With picSelector
       .Top = 1200
       .Left = 3180
       fila = 3
  End With
   Case 15
  With picSelector
       .Top = 480
       .Left = 3090
       fila = 3
  End With
  Case 16
  With picSelector
       .Top = 120
       .Left = 3090
       fila = 3
  End With
   Case 17
  With picSelector
       .Top = 1560
       .Left = 3910
       fila = 4
  End With
   Case 18
  With picSelector
       .Top = 1200
       .Left = 3910
       fila = 4
  End With
   Case 19
  With picSelector
       .Top = 480
       .Left = 3820
       fila = 4
  End With
   Case 20
  With picSelector
       .Top = 120
       .Left = 3820
       fila = 4
  End With
   Case 21
  With picSelector
       .Top = 1560
       .Left = 4640
       fila = 5
  End With
   Case 22
  With picSelector
       .Top = 1200
       .Left = 4640
       fila = 5
  End With
   Case 23
  With picSelector
       .Top = 480
       .Left = 4550
       fila = 5
  End With
   Case 24
  With picSelector
       .Top = 120
       .Left = 4550
       fila = 5
  End With
   Case 25
  With picSelector
       .Top = 1560
       .Left = 5370
       fila = 6
  End With
   Case 26
  With picSelector
       .Top = 1200
       .Left = 5370
       fila = 6
  End With
   Case 27
  With picSelector
       .Top = 480
       .Left = 5280
       fila = 6
  End With
   Case 28
  With picSelector
       .Top = 120
       .Left = 5280
       fila = 6
  End With
   Case 29
  With picSelector
       .Top = 1560
       .Left = 6100
       fila = 7
  End With
   Case 30
  With picSelector
       .Top = 1200
       .Left = 6100
       fila = 7
  End With
   Case 31
  With picSelector
       .Top = 480
       .Left = 6010
       fila = 7
  End With
   Case 32
  With picSelector
       .Top = 120
       .Left = 6010
       fila = 7
  End With
   Case 33
  With picSelector
       .Top = 1560
       .Left = 6830
       fila = 8
  End With
   Case 34
  With picSelector
       .Top = 1200
       .Left = 6830
       fila = 8
  End With
   Case 35
  With picSelector
       .Top = 480
       .Left = 6740
       fila = 8
  End With
   Case 36
  With picSelector
       .Top = 120
       .Left = 6740
       fila = 8
  End With
   Case 37
  With picSelector
       .Top = 1560
       .Left = 7560
       fila = 9
  End With
   Case 38
  With picSelector
       .Top = 1200
       .Left = 7560
       fila = 9
  End With
   Case 39
  With picSelector
       .Top = 480
       .Left = 7470
       fila = 9
  End With
   Case 40
  With picSelector
       .Top = 120
       .Left = 7470
       fila = 9
  End With
   Case 41
  With picSelector
       .Top = 1560
       .Left = 8370
       fila = 10
  End With
   Case 42
  With picSelector
       .Top = 1200
       .Left = 8370
       fila = 10
  End With
    Case 43
  With picSelector
       .Top = 480
       .Left = 8200
       fila = 10
  End With
    Case 44
  With picSelector
       .Top = 120
       .Left = 8200
       fila = 10
  End With
   Case 45
  With picSelector
       .Top = 1560
       .Left = 9000
       fila = 11
  End With
   Case 46
  With picSelector
       .Top = 1200
       .Left = 9000
       fila = 11
  End With
   Case 47
  With picSelector
       .Top = 480
       .Left = 9000
       fila = 11
  End With
   Case 48
  With picSelector
       .Top = 120
       .Left = 9000
       fila = 11
  End With
  End Select
  End Sub
  
  Private Sub eliminar_pasilloyVentanilla()
  Picture3.Visible = False
  Picture4.Visible = False
  End Sub
  
Private Sub List1_Click()
Combo1.ListIndex = List1.ListIndex
End Sub

Private Sub Timer1_Timer()
If picSelector.BackColor = vbRed Then
   picSelector.BackColor = vbGreen
   labaciento.ForeColor = vbRed
   Picture4.BackColor = vbGreen
   Picture3.BackColor = vbGreen
   Label3.ForeColor = vbRed
   Label2.ForeColor = vbRed
   ElseIf picSelector.BackColor = vbGreen Then
   picSelector.BackColor = vbRed
   labaciento.ForeColor = vbGreen
   Label3.ForeColor = vbGreen
   Label2.ForeColor = vbGreen
   Picture4.BackColor = vbRed
   Picture3.BackColor = vbRed
End If
lablugar.Caption = "PLANO DE ÓMNIBUS: " & " " & Time
End Sub

Private Sub ventanillaOPasillo(ByVal asiento As Byte)
'optiene si es ventanilla o pasillo
If asiento = 5 Or asiento = 9 Or asiento = 13 Or asiento = 17 _
 Or asiento = 21 Or asiento = 25 Or asiento = 29 Or asiento = 33 Or asiento = 37 _
 Or asiento = 41 Or asiento = 45 Or asiento = 8 Or asiento = 12 Or asiento = 16 _
 Or asiento = 20 Or asiento = 24 Or asiento = 28 Or asiento = 32 Or asiento = 36 _
 Or asiento = 40 Or asiento = 44 Or asiento = 48 Then
    EsVentanillaOPasillo 1
    Label14.Caption = "VENTANILLA"
    Else
    EsVentanillaOPasillo 0
    Label14.Caption = "PASILLO"
End If

'optiene si es isquierda o derecha dentro del omnibus
 If asiento = 8 Or asiento = 7 Or asiento = 12 Or asiento = 11 Or asiento = 16 Or asiento = 15 _
  Or asiento = 19 Or asiento = 20 Or asiento = 24 Or asiento = 23 Or asiento = 28 Or asiento = 31 _
  Or asiento = 32 Or asiento = 27 Or asiento = 36 Or asiento = 35 Or asiento = 40 Or asiento = 39 _
  Or asiento = 44 Or asiento = 43 Or asiento = 48 Or asiento = 47 Then
     Label8.Caption = "DERECHA"
     Else
     Label8.Caption = "ISQUIERDA"
 End If
End Sub
