VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmUtilidadesBaseDatosRestaurarCS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restaurar Copia de Seguridad"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   7770
   Icon            =   "FrmUtilidadesBaseDatosRestaurarCS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgLocalizarArchivo 
      Left            =   240
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRestaurar 
      Caption         =   "&Restaurar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   569
      Left            =   4560
      MouseIcon       =   "FrmUtilidadesBaseDatosRestaurarCS.frx":000C
      Picture         =   "FrmUtilidadesBaseDatosRestaurarCS.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1397
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   38
      TabIndex        =   4
      Top             =   0
      Width           =   7695
      Begin VB.Frame Frame2 
         Caption         =   "Copia de Seguridad"
         Height          =   1455
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   7215
         Begin VB.TextBox txtNombreArchivo 
            Height          =   285
            Left            =   240
            TabIndex        =   0
            Top             =   840
            Width           =   5655
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   6000
            TabIndex        =   1
            ToolTipText     =   "Buscar copia de seguridad"
            Top             =   800
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Seleccione el archivo que contiene la copia de seguridad con los datos que desea restaurar:"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   5535
         End
      End
      Begin VB.Label Label1 
         Caption         =   $"FrmUtilidadesBaseDatosRestaurarCS.frx":06A0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   7095
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   569
      Left            =   6000
      MouseIcon       =   "FrmUtilidadesBaseDatosRestaurarCS.frx":084F
      Picture         =   "FrmUtilidadesBaseDatosRestaurarCS.frx":0B59
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1397
   End
End
Attribute VB_Name = "FrmUtilidadesBaseDatosRestaurarCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Carpeta As String
Private Sub cmdBuscar_Click()
    With dlgLocalizarArchivo
        .Filter = "Copia de Seguridad (*.bck)|*.bck" ' Establece el filtro.
        .DialogTitle = "Seleccione la copia de seguridad a restaurar"
        .FileName = "Backup Farmacia.bck"
        .InitDir = App.path
        .DefaultExt = ".bck"
        .Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNNoChangeDir
        .ShowOpen
        If Len(.FileName) > 0 Then
            If .FileName <> "Backup Farmacia.bck" Then
                txtNombreArchivo = .FileName  ' Presentar el nombre del archivo seleccionado.
            End If
        End If
    End With
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdRestaurar_Click()
    On Error GoTo Solucion
    SioNo = MsgBox("Se va a restaurar la Copia de Seguridad contenida en el archivo" _
            & vbCr & txtNombreArchivo.Text & vbCr & vbCr & _
            "¿Procedemos a la restauración de los datos?", _
            vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación")
            
    If SioNo = vbNo Then Exit Sub
    
  
    'Si esta abierta la conexion cerrarla.
    If cn.State = adStateOpen Then
        cn.Close
    End If

    'pone el puntero en reloj de arena de espera
    Me.MousePointer = 11
    
    Set f = fso.GetFile(txtNombreArchivo.Text) 'Obtiene el archivo

    f.Copy App.path & "\Pedidos.mdb", True 'Copia el archivo nuevo a la carpeta de trabajo.
            
    
    Call AbrirBaseDeDatos
    
    'restaura la flecha del mouse
    Me.MousePointer = 1
    
    MsgBox "La Copia de Seguridad ha sido restaurada satisfactoriamente", _
            vbInformation, "Información"
    Unload Me
    Exit Sub
    
Solucion:
    If Not cn.State = adStateOpen Then
        Call AbrirBaseDeDatos
    End If
    MsgBox Err.Description, vbInformation, "Información"
    Exit Sub
End Sub
Private Sub Form_Load()
frmPideClave.Show vbModal
If TempNivel > 1 Then
    MsgBox "NO POSEE EL NIVEL DE AUTORIZACION PARA REALIZAR ESTA FUNCION...", vbExclamation, "Atencion !!!"
    Unload Me
End If
End Sub
Private Sub txtNombreArchivo_Change()
    If txtNombreArchivo.Text = "" Then
        cmdRestaurar.Enabled = False
    Else
        cmdRestaurar.Enabled = True
    End If
End Sub
Private Sub txtNombreArchivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then KeyCode = 0
End Sub
Private Sub txtNombreArchivo_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtNombreArchivo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtNombreArchivo.ToolTipText = txtNombreArchivo.Text & ""
End Sub
