Attribute VB_Name = "LOT_ErrorHandling"
' *============================================================================*
' *
' *     Fichero    : LOT_ErrorHandling
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sáb, 23/05/2009 23:42
' *     Versión    : 1.0
' *     Propósito  :
' *
' *
' *============================================================================*

Option Explicit
Option Base 0
' *============================================================================*
' *         Constantes
' *============================================================================*
Private Const DEBUG_MODE = False
Public Path_Log As String
' *============================================================================*
' *         Funcion:        HandleException
' *
' *         Parametros:     ErrorNumber             (I)
' *                         ErrorDescription        (I)
' *                         [ErrorSource]           (I)
' *                         [OriginalSource]        (I)
' *============================================================================*
Public Function HandleException(ByVal ErrorNumber As Long, _
                                ByVal ErrorDescription As String, _
                                Optional ByVal ErrorSource As String, _
                                Optional OriginalSource As String)
    
    ' si hay error Soltar Mensaje al usuario
    Dim strMsg As String
 On Error Resume Next
    strMsg = "=>" & vbTab & "Error Number = #" & ErrorNumber & " (Hex = " & Hex(ErrorNumber) & ")" & vbCrLf
    strMsg = strMsg & vbTab & vbTab & "Error Description = " & ErrorDescription & vbCrLf
    strMsg = strMsg & vbTab & vbTab & "Error Source = " & ErrorSource & vbCrLf
    strMsg = strMsg & vbTab & vbTab & "Error Original Source = " & OriginalSource & vbCrLf

    If DEBUG_MODE Then
        Debug.Print strMsg
    Else
        Call Trace(strMsg)
    End If
    '
    '   TODO: Almacenar los errores hasta que una llamada indique que guarde el error
    '
End Function
' *============================================================================*
' *         Funcion:        Trace
' *
' *         Parametros:     strRegistroLog          (I)
' *============================================================================*
Public Function Trace(ByVal strRegistroLog As String)
        Static nArchivo         As Integer
        Static nmFile           As String
        Dim strFechaHora        As String
        
On Error Resume Next
        '
        '   Define la fecha y hora
        strFechaHora = Date$ & " " & Time$
        '
        '   Si está activado el Debug_mode la salida es por Debug
        If DEBUG_MODE Then
            Debug.Print strFechaHora & " " & strRegistroLog
            Exit Function
        End If
        '
        ' Obtiene la ruta de log
        If Trim(Path_Log) = "" Then
            Path_Log = ThisWorkbook.Path
        End If
        '
        '   Si no está definido el log establece el nombre del fichero
        If (Trim(nmFile) = "") Then
            nmFile = Path_Log & "\" & ThisWorkbook.Name & ".log"
        End If
        '
        '   Cierra el fichero Log
        If (Trim(strRegistroLog) = "CERRAR") Then
            Close #nArchivo
            nmFile = ""
            nArchivo = 0
            Exit Function
        End If
        '
        '   Si no tenemos archivo asignado se abre
        If (nArchivo = 0) Or (nArchivo >= FreeFile) Then
            nArchivo = FreeFile
            Open nmFile For Append Shared As nArchivo
        End If
        '
        '   Grabamos en el log
        Write #nArchivo, strFechaHora & strRegistroLog
    
End Function
