VERSION 5.00
Begin VB.Form FrmUtilidadesBaseDatosRealizarCS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Realizar Copia de Seguridad"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   7890
   Icon            =   "FrmUtilidadesBaseDatosRealizarCS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   569
      Left            =   4800
      MouseIcon       =   "FrmUtilidadesBaseDatosRealizarCS.frx":058A
      Picture         =   "FrmUtilidadesBaseDatosRealizarCS.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Hacer la copia"
      Top             =   3125
      Width           =   1397
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   38
      TabIndex        =   6
      Top             =   0
      Width           =   7815
      Begin VB.Frame Frame2 
         Caption         =   "El nombre del Archivo"
         Height          =   1455
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   7335
         Begin VB.OptionButton optNombreArchivo 
            Caption         =   "&No contiene la fecha de hoy"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   3
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton optNombreArchivo 
            Caption         =   "&Contiene la fecha de hoy"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   360
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.Label lblNombreArchivo 
            AutoSize        =   -1  'True
            Caption         =   "Backup Farmacia.mdb"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   480
            TabIndex        =   9
            Top             =   840
            Width           =   2145
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Directorio de destino"
         Height          =   855
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   7335
         Begin VB.TextBox txtDestino 
            Height          =   285
            Left            =   360
            TabIndex        =   0
            Text            =   "A:\"
            Top             =   360
            Width           =   5655
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   6120
            TabIndex        =   1
            ToolTipText     =   "Buscar un directorio"
            Top             =   308
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   569
      Left            =   6240
      MouseIcon       =   "FrmUtilidadesBaseDatosRealizarCS.frx":0C1E
      Picture         =   "FrmUtilidadesBaseDatosRealizarCS.frx":0F28
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3125
      Width           =   1397
   End
End
Attribute VB_Name = "FrmUtilidadesBaseDatosRealizarCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnArchivoExistente As Boolean
Private OrigenArchivo As String
Private NombreArchivo As String
Private SpecIn As String

Private Sub cmdBuscar_Click()
    Dim bi As BROWSEINFO 'declara las variables necesarias
    Dim rtn&, pidl&, path$, pos%
    
    bi.hOwner = Me.hWnd 'centra el cuadro de dialogo en la pantalla
    bi.lpszTitle = "Seleccione el directorio para la copia de seguridad." 'asigna el texto de titulo
    bi.ulFlags = BIF_RETURNONLYFSDIRS 'el tipo de carpeta(s) para retornar
    
    pidl& = SHBrowseForFolder(bi) 'show the dialog box
    
    path = Space(512) 'asigna el maximo de caracteress
    
    Dim t As Variant
    t = SHGetPathFromIDList(ByVal pidl&, ByVal path) 'obtiene la ruta seleccionada

    pos% = InStr(path$, Chr$(0)) 'extrae la ruta desde la cadena
    SpecIn = Left(path$, pos - 1) 'asigna la ruta extraída a SpecIn
    

    If SpecIn = "" Then Exit Sub 'CESAR
    
    If DiscoApto(Left(SpecIn, 3)) = False Then
        txtDestino.Text = App.path
        Exit Sub
    End If
    txtDestino.Text = SpecIn
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo ErrProc
    If txtDestino.Text = "" Then Exit Sub
    If DiscoApto(txtDestino.Text) = False Then
        txtDestino.Text = App.path
        Exit Sub
    End If
    
    SioNo = MsgBox("¿Esta seguro de realizar la Copia de Seguridad?", _
               vbQuestion + vbYesNo, "Confirmación")
    
    If SioNo = vbNo Then Exit Sub
    Me.MousePointer = 11

    If blnArchivoExistente = True Then
        SioNo = MsgBox("Ya existe la copia de seguridad siguiente:" & _
                   vbCr & SpecIn & "\" & NombreArchivo & _
                   vbCr & vbCr & _
                   "¿Desea hacer la copia sobre el mismo archivo?", _
                   vbQuestion + vbYesNo + vbDefaultButton2, "Atención")
        If SioNo = vbNo Then Exit Sub
    End If
    Me.MousePointer = 11

    OrigenArchivo = App.path & "\Pedidos.mdb"
    

    'Si esta abierta la conexion cerrarla.
    If cn.State = adStateOpen Then
        cn.Close
    End If
    
    fso.CopyFile OrigenArchivo, SpecIn & "\" & NombreArchivo

    Call AbrirBaseDeDatos
    Me.MousePointer = 1
    
    MsgBox "La Copia de Seguridad a terminado satisfactoriamente.", vbInformation, "Información"
    Unload Me
    Exit Sub

ErrProc:
    If Not cn.State = adStateOpen Then
        Call AbrirBaseDeDatos
    End If
    MsgBox Err.Description, vbInformation, "Información"
    Exit Sub
End Sub

Private Sub txtDestino_Change()
    txtDestino.ToolTipText = txtDestino.Text & ""
    If txtDestino.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
End Sub

Private Sub txtDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            KeyCode = 0
    End Select
End Sub

Private Sub txtDestino_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    SpecIn = App.path
    txtDestino.Text = App.path
    
    txtDestino.ToolTipText = txtDestino.Text & ""
        
    NombreArchivo = "Backup Farmacia " & Format(Date$, "yyyy-mm-dd") & ".bck"
    lblNombreArchivo.Caption = NombreArchivo
End Sub

Private Sub optNombreArchivo_Click(Index As Integer)
    Select Case Index
        Case 0
            NombreArchivo = "Backup Farmacia " & Format(Date$, "yyyy-mm-dd") & ".bck"
            lblNombreArchivo.Caption = NombreArchivo
        Case 1
            NombreArchivo = "Backup Farmacia.bck"
            lblNombreArchivo.Caption = NombreArchivo
    End Select
End Sub

Function DiscoApto(ByVal drvpath As String) As Boolean
    Dim t As String
    Dim d As Drive
    
    Set d = fso.GetDrive(fso.GetDriveName(Trim(drvpath)))
    
    Select Case d.DriveType
        Case 1 'Disco 3 1/2
            If d.IsReady = False Then
                MsgBox "Disco 3 1/2  no disponible.", vbExclamation, "Atención"
                DiscoApto = False
                Exit Function
            End If
            DiscoApto = True
            'Verificar si esta habilitada la disquetera
        Case 4 'Disco CD-ROM
            MsgBox "Disco no valido"
            DiscoApto = False
            Exit Function
        Case Else
            DiscoApto = True
            Exit Function
    End Select
End Function
