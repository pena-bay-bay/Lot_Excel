VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Participante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Participante.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 28/ago/2019 11:34:25
' *     Versión    : 1.0
' *     Propósito  : Clase que representa a un participante del Sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mEntidad As EntidadNegocio
Private mUsuario As Usuario
'--- Constantes ---------------------------------------------------------------*
Private mErrorValidacion                    As Long        ' Control de errores de validación
Private Const ERR_NOMBREOBLIGADO            As Long = 1    ' Nombre de la persona obligatorio
Private Const ERR_EMAILOBLIGADO             As Long = 2    ' Correo Obligatorio
Private Const ERR_USEROBLIGADO              As Long = 4    ' Usuario obligatorio
Private Const ERR_USERDUPLI                 As Long = 8    ' Nombre de usuario duplicado
Private Const ERR_EMAILDUPLI                As Long = 16   ' Correo duplicado
'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_ERRORESVALIDACION As String = "Los datos del participante no cumple " & _
                                            "las siguientes validaciones:" & vbCrLf
Private Const MSG_NOMBREOBLIGADO As String = "* El nombre es obligatorio."
Private Const MSG_EMAILOBLIGADO As String = "* El correo electrónico es obligatorio."
Private Const MSG_USEROBLIGADO As String = "* El usuario es obligatorio."
Private Const MSG_USERDUPLI As String = "* El nombre de usuario está ya registrado."
Private Const MSG_EMAILDUPLI As String = "* El correo electrónico no es único, ya existe."
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public Property Get Id() As Integer
    Id = mEntidad.Id
End Property

Public Property Let Id(ByVal vNewValue As Integer)
    mEntidad.Id = vNewValue
End Property

Public Property Get Usuario() As String
    Usuario = mUsuario.IdUsuario
End Property

Public Property Let Usuario(ByVal vNewValue As String)
    mUsuario.IdUsuario = vNewValue
End Property

Public Property Get CorreoElectronico() As String
    CorreoElectronico = mUsuario.Email
End Property

Public Property Let CorreoElectronico(ByVal vNewValue As String)
    mUsuario.Email = vNewValue
End Property

Public Property Get Nombre() As String
    Nombre = mUsuario.Nombre
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    mUsuario.Nombre = vNewValue
End Property

Public Property Get Apellido1() As String
    Apellido1 = mUsuario.Persona.Apellido1
End Property

Public Property Let Apellido1(ByVal vNewValue As String)
    mUsuario.Persona.Apellido1 = vNewValue
End Property

Public Property Get Apellido2() As String
    Apellido2 = mUsuario.Persona.Apellido2
End Property

Public Property Let Apellido2(ByVal vNewValue As String)
    mUsuario.Persona.Apellido2 = vNewValue
End Property

Public Property Get Rol() As Integer
    Rol = mUsuario.Rol
End Property

Public Property Let Rol(ByVal vNewValue As Integer)
    mUsuario.Rol = vNewValue
End Property

Public Property Get NombreRol() As String
    Dim mVar As Variant
    mVar = Split(NOMBRES_ROLES, ";")
    NombreRol = mVar(Rol)
End Property
'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mUsuario = New Usuario
    Set mEntidad = New EntidadNegocio
End Sub

Private Sub Class_Terminate()
    Set mUsuario = Nothing
    Set mEntidad = Nothing
End Sub
'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : ma., 24/sep/2019 19:57:52
' Propósito      : Espresar las propiedades del participante en formato texto
' Retorno        : Información de la clase en formato texto
'------------------------------------------------------------------------------*
Public Function ToString() As String
    Dim mStr As String
    
  On Error GoTo ToString_Error
    '
    '
    '
    mStr = "#$1, Persona: $2 $3 $4, Correo: $5, Usuario: $6, Rol: $7"
    '
    '
    '
    mStr = Replace(mStr, "$1", CStr(Id))
    mStr = Replace(mStr, "$2", Nombre)
    mStr = Replace(mStr, "$3", Apellido1)
    mStr = Replace(mStr, "$4", Apellido2)
    mStr = Replace(mStr, "$5", CorreoElectronico)
    mStr = Replace(mStr, "$6", Usuario)
    mStr = Replace(mStr, "$7", NombreRol)
    '
    '
    '
    ToString = mStr
    
  On Error GoTo 0
ToString__CleanExit:
    Exit Function
           
ToString_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Participante.ToString", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Participante.ToString", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : IsValid
' Fecha          : lu., 07/oct/2019 14:06:06
' Propósito      : Validar las propiedades del objeto
' Retorno        : Devuelve False si hay al menos una propiedad erronea
'------------------------------------------------------------------------------*
Public Function IsValid() As Boolean
  On Error GoTo IsValid_Error
    '
    mErrorValidacion = 0
    '
    '
    If Len(Me.Nombre) = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_NOMBREOBLIGADO
    End If
    '
    '
    If Len(Me.CorreoElectronico) = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_EMAILOBLIGADO
    End If
    '
    '
    If Len(Me.Usuario) = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_USEROBLIGADO
    End If
    '
    '  TODO: Implementar control de user dupli en ParticipanteEngine
    '
    '  mUser = mEng.
    ' if mEntidad.IsNew and Me.Usuario =
    '  TODO: Implementar control de correo dupli en ParticipanteEngine
    '
    '   Devolvemos el resultado de la validación
    '
    
    If mErrorValidacion = 0 Then
        IsValid = True
    Else
        IsValid = False
    End If
  
  On Error GoTo 0
IsValid__CleanExit:
    Exit Function
            
IsValid_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Participante.IsValid", ErrSource)
    Err.Raise ErrNumber, "Participante.IsValid", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetMensaje
' Fecha          : lu., 07/oct/2019 14:17:28
' Propósito      : Devuelve los errores de consistencia del objeto
' Retorno        : Cadena de texto con todos los errores
'------------------------------------------------------------------------------*
Public Function GetMensaje() As String
    Dim sResult             As String
  
  On Error GoTo GetMensaje_Error
    '
    '   Si no hay error inicializamos el mensaje
    '
    If mErrorValidacion = 0 Then
        sResult = ""
    Else
        '
        '   Si hay algún error inicializamos la cabecera del error
        '
        sResult = MSG_ERRORESVALIDACION
    End If
    If (mErrorValidacion And ERR_NOMBREOBLIGADO) Then
        sResult = sResult & vbTab & MSG_NOMBREOBLIGADO & vbCrLf
    End If
    If (mErrorValidacion And ERR_EMAILOBLIGADO) Then
        sResult = sResult & vbTab & MSG_EMAILOBLIGADO & vbCrLf
    End If
    If (mErrorValidacion And ERR_USEROBLIGADO) Then
        sResult = sResult & vbTab & MSG_USEROBLIGADO & vbCrLf
    End If
    If (mErrorValidacion And ERR_USERDUPLI) Then
        sResult = sResult & vbTab & MSG_USERDUPLI & vbCrLf
    End If
    If (mErrorValidacion And ERR_EMAILDUPLI) Then
        sResult = sResult & vbTab & MSG_EMAILDUPLI & vbCrLf
    End If
    '
    '   Devolvemos el mensaje
    '
    GetMensaje = sResult
  
  On Error GoTo 0
GetMensaje__CleanExit:
    Exit Function
    
GetMensaje_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Participante.GetMensaje", ErrSource)
    Err.Raise ErrNumber, "Participante.GetMensaje", ErrDescription
End Function

'' *===========(EOF): Participante.cls


