Attribute VB_Name = "Module1"
Public cbd As String 'guarda la clave de la base de datos
Public cn As New Connection
Public SioNo As String
Public VarPro As String
Public vidPro As Integer
Public strSQL As String
Public vAgrega As Boolean
Public vProducto As String
Public VarTroquel As String
Public rptGeneral As Object
Public vIdCliente As Integer
Public vIdFactura As Integer
Public TempNivel As Integer
Public vUsu As String 'usuario autorizado
Public vFecFac As String
Public vFecVto As Date
Public rtaLargoPlazo As Boolean

'------ PARA REALIZAR COPIAS DE SEGURIDAD DE DATOS -------
'Public fso As New FileSystemObject ' Trabajar con archivos.
Public fso As New FileSystemObject
Public d As Drive
Public fo As Folder
Public f As File
'--------------------------------------------------
' COPIA SE SEGURIDAD (Seleccionar carpeta)
'---------------------------------------------------
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Type SHITEMID
     cb As Long
     abID As Byte
End Type
Type ITEMIDLIST
     mkid As SHITEMID
End Type
Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type
Public Const NOERROR = 0
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
'----------------------------------------------------
Public Sub CompactarBaseDatos()
    On Error GoTo Solucion
    
    SioNo = MsgBox("Atención: se va a reducir el tamaño de la base de datos eliminando el " & _
        vbCr & "espacio que dejan los registros borrados." & vbCr & _
        vbCr & "Este proceso puede durar algunos segundos y no debe ser interrumpido " & _
        vbCr & "para evitar que la información se degrade." & vbCr & _
        vbCr & "Antes de realizar esta operación es recomendable realizar una copia de " & _
        vbCr & "seguridad de los datos.", vbOKCancel + vbExclamation, _
        "Compactar Base de Datos")
    If SioNo = vbCancel Then
        Exit Sub
    End If

    Call CerrarBaseDeDatos
    
    ' Compactar la base de datos con ADO.
    Dim je As New JRO.JetEngine
    je.CompactDatabase "Data Source=" & App.path & "\Pedidos.mdb" & _
                       ";Jet OLEDB:Database Password=", _
                       "Data Source=" & App.path & "\xxx_221133.mdb" & _
                       ";Jet OLEDB:Database Password=julian"
    ' Eliminar la base de datos original.
    fso.DeleteFile App.path & "\Pedidos.mdb", True
    ' Renombrar la base con el nombre original.
    Set f = fso.GetFile(App.path & "\xxx_221133.mdb")
    f.Name = "Pedidos.mdb"
    
    Call AbrirBaseDeDatos
    
    MsgBox "La Base de Datos ha sido compactada satisfactoriamente.", _
            vbInformation, "Información"
        
    Exit Sub
    
Solucion:
    Call AbrirBaseDeDatos
    ' Mostrar el mensaje de error.
    MsgBox "Error al compactar la base de datos: " & vbCrLf & _
            Err.Number & " - " & Err.Description, _
            vbInformation, "Información"
    Err.Clear
End Sub
Public Sub CerrarBaseDeDatos()
    If cn.State = adStateClosed Then Exit Sub
    cn.Close
    Set cn = Nothing
End Sub
Public Sub AbrirBaseDeDatos()
    If cn.State = adStateOpen Then Exit Sub
    cn.CursorLocation = adUseClient
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; " & _
            "Data Source= " & App.path & "\Pedidos.mdb;" & _
            "Mode=ReadWrite;Jet OLEDB:Database Password=" & cbd
End Sub
Public Function DeCrypt(strEncrypted)
'para desencriptar claves
    g_Key = "O_[:S&&]44AK;;^&*R?ZN^9_7LL),VG;;$=QY,JMM1*2*KW<^@I@T,3YY6V0]$2DA)+T0RZIOC`;>:FA[)6P)#=13N&"
    Dim strChar, iKeyChar, iStringChar, i
    For i = 1 To Len(strEncrypted)
       iKeyChar = (Asc(Mid(g_Key, i, 1)))
       iStringChar = Asc(Mid(strEncrypted, i, 1))
       iDeCryptChar = iStringChar - iKeyChar
       strDecrypted = strDecrypted & Chr(iDeCryptChar)
    Next
    DeCrypt = strDecrypted
End Function

Public Function FinDeMes(Vfin)
'identifica el ultimo dia del mes actual en los campos fechas para filtros

If Month(Date) = 1 Or Month(Date) = 3 Or Month(Date) = 5 Or Month(Date) = 7 Or Month(Date) = 8 Or Month(Date) = 10 Or Month(Date) = 12 Then
    Vfin = (Format("31/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy"))
End If
If Month(Date) = 4 Or Month(Date) = 6 Or Month(Date) = 9 Or Month(Date) = 11 Then
    Vfin = (Format("30/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy"))
End If
If Month(Date) = 2 Then
    Vfin = (Format("28/" & Month(Date) & "/" & Year(Date), "dd/mm/yyyy"))
End If
End Function

Public Sub ValidarDigitos(DatosActuales As String, Caracter As Integer)
  ' Salimos si se ha pulsado la tecla de Retroceso
  If Caracter = 8 Then Exit Sub
  ' Salimos si es de 0 a 9
  If InStr("0123456789", Chr$(Caracter)) Then Exit Sub
  ' Si es punto y no está en el contenido salimos
  If Caracter = 46 And InStr(DatosActuales, ".") = 0 Then Exit Sub
  ' Borramos el Caracter introducido
  Caracter = 0
End Sub
