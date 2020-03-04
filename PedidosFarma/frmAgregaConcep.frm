VERSION 5.00
Begin VB.Form frmAgregaConcep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agrega Concepto de Egresos..."
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   Icon            =   "frmAgregaConcep.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   3480
      Picture         =   "frmAgregaConcep.frx":0442
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtObs 
      Height          =   320
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox txtDes 
      Height          =   320
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion del Concepto:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2265
   End
End
Attribute VB_Name = "frmAgregaConcep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsConcep As New ADODB.Recordset
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdGrabar_Click()
SioNo = MsgBox("ESTA SEGURO DE GRABAR ESTE NUEVO CONCEPTO ?", vbInformation + vbYesNo, "ATENCION !")
If SioNo = vbNo Then
    Exit Sub
End If
If Len(txtDes.Text) = 0 Then
    MsgBox "DEBE INGRESAR ALGUNA DESCRIPCION DEL CONCEPTO DE EGRESO...", vbExclamation, "Atencion !!!"
    txtDes.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
rsConcep.Open "ConceptosEgresos", cn, adOpenDynamic, adLockOptimistic, adCmdTable
rsConcep.AddNew
rsConcep!descripcion = Trim(txtDes.Text)
rsConcep!observaciones = Trim(txtObs.Text)
rsConcep.Update
rsConcep.Close
Set rsConcep = Nothing
Unload Me
End Sub
Private Sub txtDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtObs.SetFocus
End If
End Sub
Private Sub txtObs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdGrabar.SetFocus
End If
End Sub
