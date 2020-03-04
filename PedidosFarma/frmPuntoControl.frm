VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.8#0"; "AS-IFce1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPuntoControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Punto de Control de Cajas ..."
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
   Icon            =   "frmPuntoControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Objetivo"
      Height          =   855
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   7335
      Begin VB.Label Label1 
         Caption         =   $"frmPuntoControl.frx":030A
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6975
      End
   End
   Begin AIFCmp1.asxPowerButton cmdGrabar 
      Height          =   735
      Left            =   4680
      TabIndex        =   10
      Top             =   3480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      BackColor       =   8454016
      Picture         =   "frmPuntoControl.frx":039D
      Caption         =   "&Grabar"
      CaptionAlignment=   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   3
      PictureOffsetX  =   15
   End
   Begin VB.TextBox txtObser 
      Height          =   375
      Left            =   240
      MaxLength       =   100
      TabIndex        =   9
      Top             =   2760
      Width           =   9375
   End
   Begin VB.TextBox txtGrande 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   8
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txtMostrador 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txtChica 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   240
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1680
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   400
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   0
      Format          =   69795841
      CurrentDate     =   39674
   End
   Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx5 
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   423
      Caption         =   "Observaciones"
   End
   Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx2 
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   423
      Caption         =   "Importe Caja Chica"
   End
   Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx1 
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   423
      Caption         =   "Fecha"
   End
   Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx3 
      Height          =   240
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   423
      Caption         =   "Importe Caja Grande"
   End
   Begin AIFCmp1.asxLineHeaderEx asxLineHeaderEx4 
      Height          =   240
      Left            =   6840
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   423
      Caption         =   "Importe Caja Mostrador"
   End
   Begin AIFCmp1.asxPowerButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   7320
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      BackColor       =   12632319
      Picture         =   "frmPuntoControl.frx":1077
      Caption         =   "&Cancelar"
      CaptionAlignment=   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   3
      PictureOffsetX  =   15
   End
End
Attribute VB_Name = "frmPuntoControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsPunto As New ADODB.Recordset
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdGrabar_Click()
If Len(txtMostrador.Text) = 0 Then
    MsgBox "DEBE HABER UN IMPORTE EN CAJA MOSTRADOR PARA GRABAR LOS DATOS....!", vbCritical, "ATENCION !"
    txtMostrador.SetFocus
    Exit Sub
End If
rsPunto.Find "fecha =#" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#", , adSearchForward, 1
If rsPunto.EOF Then
    rsPunto.AddNew
End If
rsPunto!fecha = dtpFecha.Value
rsPunto!cajachica = txtChica.Text
rsPunto!cajagrande = txtGrande.Text
rsPunto!cajamostrador = txtMostrador.Text
rsPunto!observaciones = txtObser.Text
rsPunto.Update
MsgBox "EL PUNTO DE CONTROL FUE ALMACENADO CORRECTAMENTE ...", vbInformation, "FIN DE PROCESO..."
rsPunto.Close
Set rsPunto = Nothing
Unload Me
End Sub
Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyAscii = 0
    txtChica.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub
Private Sub Form_Load()
dtpFecha.Value = Date
rsPunto.Open "select * from puntocontrol order by fecha", cn, adOpenDynamic, adLockOptimistic, adCmdText
rsPunto.Find "fecha =#" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#", , adSearchForward, 1
If rsPunto.EOF Then
    txtChica.Text = 0
    txtGrande.Text = 0
    txtMostrador.Text = 0
    txtObser.Text = ""
Else
    txtChica.Text = rsPunto!cajachica
    txtGrande.Text = rsPunto!cajagrande
    txtMostrador.Text = rsPunto!cajamostrador
    txtObser.Text = rsPunto!observaciones
End If
End Sub
Private Sub Form_Unload(cancel As Integer)
If rsPunto.State = 1 Then
    rsPunto.Close
End If
End Sub
Private Sub txtChica_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtGrande.SetFocus
        SendKeys "{home}+{end}"
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtGrande_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtMostrador.SetFocus
        SendKeys "{home}+{end}"
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtMostrador_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 13
        KeyAscii = 0
        txtObser.SetFocus
    Case 44 'para que no acepte la entrada de la coma
        KeyAscii = 0
    Case 45
    Case 46
    Case 48 To 57
    Case Else: KeyAscii = 0
End Select
End Sub
Private Sub txtObser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGrabar.SetFocus
End If
End Sub
