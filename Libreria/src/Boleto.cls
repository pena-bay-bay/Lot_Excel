VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Boleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Boleto.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ju., 08/ago/2019 20:04:47
' *     Versión    : 1.0
' *     Propósito  : Documento de participación en los concursos de la ONLAE
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mEntidad As EntidadNegocio
Private mJuego As Juego
Private mFechaSorteo As Date
Private mFechaValidez As Date
Private mCadencia As Integer
Private mApuestas As Apuestas
Private mEsMultiple As Boolean
Private mIdParticipante As Integer
Private mNumeroApuestas As Integer
Private mJoker As Long
Private mSituacion As Integer
Private mMillon As String
Private mReintegroClave As Integer
Private mCoste As Currency
Private mImportePremios As Currency
Private mDesglosePremios As String
Private mComentarios As String


'--- Constantes ---------------------------------------------------------------*
Private mErrorValidacion                    As Long        ' Control de errores de validación
Private Const ERR_JUEGOOBLIGATORIO          As Long = 1    ' Juego Obligatorio
Private Const ERR_FECHAOBLIGATORIA          As Long = 2    ' Fecha Obligatoria
Private Const ERR_MULTANDMASAPUESTAS        As Long = 4    ' Boleto multiple solo una apuesta
Private Const ERR_APUESTASOBLIGATORIAS      As Long = 8    ' Apuestas obligatorias
Private Const ERR_APUESTASNOVALIDAS         As Long = 16   ' Apuesta incorrecta
Private Const ERR_PARTICIPANTEOBLIGATORIO   As Long = 32   ' Participante obligatorio si el boleto es validado
Private Const ERR_REINTEGROOBLIGATORIO      As Long = 64   ' Reintegro Obligatorio
Private Const ERR_CLAVEOBLIGATORIO          As Long = 128  ' Clave Obligatoria
Private Const ERR_JOKEROBLIGATORIO          As Long = 256  ' Joker Obligatoria
Private Const ERR_MILLONOBLIGATORIO         As Long = 512  ' Millon Obligatoria
Private Const ERR_DIASEMINCORRECTO          As Long = 1024 ' Dia Semana Incorrecto
Private Const ERR_BONOLOTOMINDOS            As Long = 2048 ' Boleto bonoloto minimo 2 apuestas
Private Const ERR_COSTENOTCERO              As Long = 4096 ' Boleto validado con coste 0
Private Const ERR_FECHAVENCIMIENTOERR       As Long = 8192 ' Fecha de vencimiento erronea
Private Const ERR_MILLONMALCONSTRUIDO       As Long = 16384 ' Millon mal construido

'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_ERRORESVALIDACION As String = "Los datos del boleto no cumple " & _
                                            "las siguientes validaciones:" & vbCrLf
Private Const MSG_JUEGOOBLIGATORIO As String = "* El juego es un dato obligatorio."
Private Const MSG_FECHAOBLIGATORIA As String = "* La fecha del sorteo es obligatoria."
Private Const MSG_MULTANDMASAPUESTAS As String = "* Boleto múltiple con sólo una apuesta."
Private Const MSG_APUESTASOBLIGATORIAS As String = "* Apuestas obligatorias."
Private Const MSG_APUESTASNOVALIDAS As String = "* Apuestas no válidas."
Private Const MSG_PARTICIPANTEOBLIGATORIO As String = "* El participante del boleto es obligatorio."
Private Const MSG_REINTEGROOBLIGATORIO As String = "* El número de reintegro del boleto es obligatorio."
Private Const MSG_CLAVEOBLIGATORIO As String = "* La clave del boleto es obligatoria."
Private Const MSG_JOKEROBLIGATORIO As String = "* El Joker no es numérico."
Private Const MSG_MILLONOBLIGATORIO As String = "* El millon es obligatorio. "
Private Const MSG_DIASEMINCORRECTO As String = "* El día de la semana del sorteo no es correcto."
Private Const MSG_BONOLOTOMINDOS As String = "* El boleto debe contener un mínimo de 2 apuestas."
Private Const MSG_COSTENOTCERO As String = "* El coste es obligatorio si el boleto está validado."
Private Const MSG_FECHAVENCIMIENTOERR As String = "* La fecha de vencimiento es erronea para este juego."
Private Const MSG_MILLONMALCONSTRUIDO As String = "* Codigo del millon mal construido 'AAAnnnnn'."


'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'
'   Identificador único del Boleto
'
Public Property Get Id() As Integer
    Id = mEntidad.Id
End Property

Public Property Let Id(ByVal vNewValue As Integer)
    mEntidad.Id = vNewValue
End Property
'
'   Juego al que pertenece la apuesta
'
Public Property Get Juego() As Juego
    Juego = mJuego
End Property

Public Property Let Juego(ByVal vNewValue As Juego)
    mJuego = vNewValue
End Property
'
'   Fecha de celebración del sorteo
'
Public Property Get FechaSorteo() As Date
    FechaSorteo = mFechaSorteo
End Property

Public Property Let FechaSorteo(ByVal vNewValue As Date)
    mFechaSorteo = vNewValue
End Property
'
'   Fecha hasta la que el boleto es válido para múltisorteo
'
Public Property Get FechaValidez() As Date
    FechaValidez = mFechaValidez
End Property

Public Property Let FechaValidez(ByVal vNewValue As Date)
    mFechaValidez = vNewValue
End Property
'
'   Periodicidad del boleto: 1.-Diaria, 2.-Semanal, 3.-bisemanal
'
Public Property Get Cadencia() As Integer
    Cadencia = mCadencia
End Property

Public Property Let Cadencia(ByVal vNewValue As Integer)
    mCadencia = vNewValue
End Property
'
'   Colección de apuestas del boleto.
'
Public Property Get Apuestas() As Apuestas
    Set Apuestas = mApuestas
End Property

Public Property Set Apuestas(ByVal vNewValue As Apuestas)
    Set mApuestas = vNewValue
End Property
'
'   Indicador de apuesta múltiple.
'
Public Property Get EsMultiple() As Boolean
    EsMultiple = mEsMultiple
End Property

Public Property Let EsMultiple(ByVal vNewValue As Boolean)
    mEsMultiple = vNewValue
End Property
'
'   Identificador del participante que valida el boleto
'
Public Property Get IdParticipante() As Integer
    IdParticipante = mIdParticipante
End Property

Public Property Let IdParticipante(ByVal vNewValue As Integer)
    mIdParticipante = vNewValue
End Property
'
'   Número total de apuestas
'
Public Property Get NumeroApuestas() As Integer
    Dim mApt As Apuesta
    For Each mApt In Apuestas.Items
        mNumeroApuestas = mNumeroApuestas + mApt.TotalApuestas
    Next mApt
    NumeroApuestas = mNumeroApuestas
End Property

'Public Property Let NumeroApuestas(ByVal vNewValue As Integer)
'    mNumeroApuestas = vNewValue
'End Property
'
'   Numero Joker del juego de la primitiva
'
Public Property Get Joker() As Long
    Joker = mJoker
End Property

Public Property Let Joker(ByVal vNewValue As Long)
    mJoker = vNewValue
End Property
'
'   Las situaciones del boleto pueden ser: 0 - Borrador,
'        1- validado o activo, 2- premiado y 3- sorteado
'
Public Property Get Situacion() As Integer
    Situacion = mSituacion
End Property

Public Property Let Situacion(ByVal vNewValue As Integer)
    mSituacion = vNewValue
End Property
'
'   Situación del boleto en formato texto
'
Public Property Get SituacionText() As String
    Dim mVar As Variant
    mVar = Split(NOMBRES_SITBOLETO, ";")
    SituacionText = mVar(mSituacion)
End Property
'
'   Código del millón del juego Euromillones
'
Public Property Get Millon() As String
    Millon = mMillon
End Property

Public Property Let Millon(ByVal vNewValue As String)
    mMillon = vNewValue
End Property
'
'   Número del reintegro para bonoloto y primitiva y clave para el gordo
'
Public Property Get ReintegroClave() As Integer
    ReintegroClave = mReintegroClave
End Property

Public Property Let ReintegroClave(ByVal vNewValue As Integer)
    mReintegroClave = vNewValue
End Property
'
'   Coste de la apuesta para este juego
'
Public Property Get Coste() As Currency
    Coste = mCoste
End Property

Public Property Let Coste(ByVal vNewValue As Currency)
    mCoste = vNewValue
End Property
'
'   Importe de la apuesta obtenido en premios
'
Public Property Get ImportePremios() As Currency
    ImportePremios = mImportePremios
End Property

Public Property Let ImportePremios(ByVal vNewValue As Currency)
    mImportePremios = vNewValue
End Property
'
'   Resumen de los premios obtenidos por el boleto
'
Public Property Get DesglosePremios() As String
    DesglosePremios = mDesglosePremios
End Property

Public Property Let DesglosePremios(ByVal vNewValue As String)
    mDesglosePremios = vNewValue
End Property
'
'   Comentarios del usuario en el boleto
'
Public Property Get Comentarios() As String
    Comentarios = mComentarios
End Property

Public Property Let Comentarios(ByVal vNewValue As String)
    mComentarios = vNewValue
End Property

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mEntidad = New EntidadNegocio
    Set mApuestas = New Apuestas
    mJuego = SinJuego
    mCadencia = 0
    mEsMultiple = False
    mIdParticipante = 0
    mNumeroApuestas = 0
    mJoker = 0
    mSituacion = blBorrador
    mMillon = Empty
    mReintegroClave = -1
    mCoste = 0
    mImportePremios = 0
    mDesglosePremios = Empty
    mComentarios = Empty
End Sub

'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : EsValido
' Fecha          : vi., 04/oct/2019 22:41:41
' Propósito      : Valida el contenido de la apuesta, y si es válido para el
'                  juego devuelve True
' Retorno        : True si el boleto es válido, False si hay errores
'------------------------------------------------------------------------------*
Public Function EsValido() As Boolean
    Dim mApt As Apuesta
    Dim mDia As Integer
    Dim mDiaV As Integer
    
  On Error GoTo EsValido_Error
    '
    mErrorValidacion = 0
    '
    '   Juego Obligatorio
    If mJuego = SinJuego Then
        mErrorValidacion = mErrorValidacion + ERR_JUEGOOBLIGATORIO
    End If
    '
    '   Fecha de sorteo Obligatoria
    If mFechaSorteo = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_FECHAOBLIGATORIA
    End If
    '
    '   Boleto múltiple con más de una apuesta simple
    If mEsMultiple And (mApuestas.Count > 1) Then
        mErrorValidacion = mErrorValidacion + ERR_MULTANDMASAPUESTAS
    End If
    '
    '   Boleto sin apuestas o apuestas no válidas
    If mApuestas.Count = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_APUESTASOBLIGATORIAS
    Else
        For Each mApt In mApuestas.Items
            If Not (mApt.EsValido) Then
                mErrorValidacion = mErrorValidacion + ERR_APUESTASNOVALIDAS
                Exit For
            End If
        Next mApt
    End If
    '
    '   Boleto validado sin el participante
    If mSituacion >= blValidado And mIdParticipante = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_PARTICIPANTEOBLIGATORIO
    End If
    '
    '   Calculamos el dia del sorteo 1= Lunes
    '
    mDia = Weekday(mFechaSorteo, vbMonday)
    If Cadencia >= 2 Then
        mDiaV = Weekday(mFechaValidez, vbMonday)
    End If
    '
    '
    '   Analisis según juego
    Select Case mJuego
        Case LoteriaPrimitiva:
            '
            '   Reintegro Obligatorio
            If mSituacion >= blValidado And (ReintegroClave < 0 Or ReintegroClave > 9) Then
                mErrorValidacion = mErrorValidacion + ERR_REINTEGROOBLIGATORIO
            End If
            '
            '   Joker Obligatorio
            If mSituacion >= blValidado And Joker = 0 Then
                mErrorValidacion = mErrorValidacion + ERR_JOKEROBLIGATORIO
            End If
            '
            '   Dia incorrecto no es jueves o sabado
            If Not (mDia = 4 Or mDia = 6) Then
                mErrorValidacion = mErrorValidacion + ERR_DIASEMINCORRECTO
            End If
            '
            '   Fecha de vencimiento no correcta
            If (Not (mDiaV = 4 Or mDiaV = 6)) And Cadencia >= 2 Then
                mErrorValidacion = mErrorValidacion + ERR_FECHAVENCIMIENTOERR
            End If
            
        Case bonoloto:
            '
            '   Reintegro Obligatorio
            If mSituacion >= blValidado And (ReintegroClave < 0 Or ReintegroClave > 9) Then
                mErrorValidacion = mErrorValidacion + ERR_REINTEGROOBLIGATORIO
            End If
            '
            '   Dia incorrecto Es domingo
            If mDia = 7 Then
                mErrorValidacion = mErrorValidacion + ERR_DIASEMINCORRECTO
            End If
            '
            '   Minimo dos apuestas
            '
            If NumeroApuestas < 2 Then
                 mErrorValidacion = mErrorValidacion + ERR_BONOLOTOMINDOS
            End If
            '
            '   Fecha de vencimiento no correcta
            If (mDiaV = 7) And Cadencia >= 2 Then
                mErrorValidacion = mErrorValidacion + ERR_FECHAVENCIMIENTOERR
            End If
            
        Case Euromillones:
            '
            '   Millon Obligatorio
            If mSituacion >= blValidado And Len(Millon) = 0 Then
                mErrorValidacion = mErrorValidacion + ERR_MILLONOBLIGATORIO
            End If
            '
            '   Dia incorrecto no es martes ni viernes
            If Not (mDia = 2 Or mDiaV = 5) Then
                mErrorValidacion = mErrorValidacion + ERR_DIASEMINCORRECTO
            End If
            '
            '   Fecha de vencimiento no correcta
            If Not (mDiaV = 2 Or mDiaV = 5) And Cadencia >= 2 Then
                mErrorValidacion = mErrorValidacion + ERR_FECHAVENCIMIENTOERR
            End If
            '
            '   Comprobar Euromillon
            '
            If mSituacion >= blValidado And _
            (IsNumeric(mId(Millon, 1, 1)) Or _
             IsNumeric(mId(Millon, 2, 1)) Or _
             IsNumeric(mId(Millon, 3, 1))) And _
            (Not IsNumeric(mId(Millon, 4, 5))) Then
                mErrorValidacion = mErrorValidacion + ERR_MILLONMALCONSTRUIDO
            End If
        
        Case gordoPrimitiva:
            '
            '   Reintegro Obligatorio
            If mSituacion >= blValidado And (ReintegroClave < 0 Or ReintegroClave > 9) Then
                mErrorValidacion = mErrorValidacion + ERR_REINTEGROOBLIGATORIO
            End If
            '
            '   Dia incorrecto no es domingo
            If Not (mDia = 7) Then
                mErrorValidacion = mErrorValidacion + ERR_DIASEMINCORRECTO
            End If
            '
            '   Fecha de vencimiento no correcta
            If Not (mDiaV = 7) And Cadencia >= 2 Then
                mErrorValidacion = mErrorValidacion + ERR_FECHAVENCIMIENTOERR
            End If
        
    End Select
    '
    '   Boleto validado con coste 0
    If mSituacion >= blValidado And Coste = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_COSTENOTCERO
    End If
    '
    '   Devolvemos el resultado de la validación
    '
    If mErrorValidacion = 0 Then
        EsValido = True
    Else
        EsValido = False
    End If
    
  On Error GoTo 0
EsValido__CleanExit:
    Exit Function
            
EsValido_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Boleto.EsValido", ErrSource)
    Err.Raise ErrNumber, "Boleto.EsValido", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetMensaje
' Fecha          : dd/MMM/yyyy
' Propósito      : Devuelve los errores de la apuesta, si no es válido
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
    If (mErrorValidacion And ERR_JUEGOOBLIGATORIO) Then
        sResult = sResult & vbTab & MSG_JUEGOOBLIGATORIO & vbCrLf
    End If
    If (mErrorValidacion And ERR_FECHAOBLIGATORIA) Then
        sResult = sResult & vbTab & MSG_FECHAOBLIGATORIA & vbCrLf
    End If
    If (mErrorValidacion And ERR_MULTANDMASAPUESTAS) Then
        sResult = sResult & vbTab & MSG_MULTANDMASAPUESTAS & vbCrLf
    End If
    If (mErrorValidacion And ERR_APUESTASOBLIGATORIAS) Then
        sResult = sResult & vbTab & MSG_APUESTASOBLIGATORIAS & vbCrLf
    End If
    If (mErrorValidacion And ERR_APUESTASNOVALIDAS) Then
        sResult = sResult & vbTab & MSG_APUESTASNOVALIDAS & vbCrLf
    End If
    If (mErrorValidacion And ERR_PARTICIPANTEOBLIGATORIO) Then
        sResult = sResult & vbTab & MSG_PARTICIPANTEOBLIGATORIO & vbCrLf
    End If
    If (mErrorValidacion And ERR_REINTEGROOBLIGATORIO) Then
        sResult = sResult & vbTab & MSG_REINTEGROOBLIGATORIO & vbCrLf
    End If
    If (mErrorValidacion And ERR_CLAVEOBLIGATORIO) Then
        sResult = sResult & vbTab & MSG_CLAVEOBLIGATORIO & vbCrLf
    End If
    If (mErrorValidacion And ERR_JOKEROBLIGATORIO) Then
        sResult = sResult & vbTab & MSG_JOKEROBLIGATORIO & vbCrLf
    End If
    If (mErrorValidacion And ERR_MILLONOBLIGATORIO) Then
        sResult = sResult & vbTab & MSG_MILLONOBLIGATORIO & vbCrLf
    End If
    If (mErrorValidacion And ERR_DIASEMINCORRECTO) Then
        sResult = sResult & vbTab & MSG_DIASEMINCORRECTO & vbCrLf
    End If
    If (mErrorValidacion And ERR_BONOLOTOMINDOS) Then
        sResult = sResult & vbTab & MSG_BONOLOTOMINDOS & vbCrLf
    End If
    If (mErrorValidacion And ERR_COSTENOTCERO) Then
        sResult = sResult & vbTab & MSG_COSTENOTCERO & vbCrLf
    End If
    If (mErrorValidacion And ERR_FECHAVENCIMIENTOERR) Then
        sResult = sResult & vbTab & MSG_FECHAVENCIMIENTOERR & vbCrLf
    End If
    If (mErrorValidacion And ERR_MILLONMALCONSTRUIDO) Then
        sResult = sResult & vbTab & MSG_MILLONMALCONSTRUIDO & vbCrLf
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
    Call HandleException(ErrNumber, ErrDescription, "Boleto.GetMensaje", ErrSource)
    Err.Raise ErrNumber, "Boleto.GetMensaje", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : mi., 28/ago/2019 12:36:07
' Propósito      : Visualiza las propiedades de un Boleto
' Retorno        : Cadena de texto con las propiedades del boleto
'------------------------------------------------------------------------------*
Public Function ToString() As String
    Dim mTemplate As String
    Dim mTmp As String
    
  On Error GoTo ToString_Error
    '
    ' Id Juego Situacion Cadencia FechaSorteo FechaValidez NumeroApuestas EsMultiple
    ' Joker Millon Participante ReintegroClave Coste   ImportePremios  DesglosePremios Comentarios
    '
    Select Case mJuego
        Case Euromillones:
            mTemplate = "#$1, $2, Fecha:$3, Sit:$4, Apuestas:$5, Multiple:$6," & _
                        " Millon:$7, Participante:$8, Coste:$9 , Premios:$10"
        Case LoteriaPrimitiva:
            mTemplate = "#$1, $2, Fecha:$3, Sit:$4, Apuestas:$5, Multiple:$6," & _
                         " Joker:$11, Reintegro:$12, Participante:$8, Coste:$9, Premios:$10"
        Case gordoPrimitiva:
            mTemplate = "#$1, $2, Fecha:$3, Sit:$4, Apuestas:$5, Multiple:$6," & _
                         " Clave:$12, Participante:$8, Coste:$9, Premios:$10"
    End Select

    mTmp = Replace(mTemplate, "$1", CStr(Id))
    mTmp = Replace(mTmp, "$3", Format(mFechaSorteo, "dd/mm/yyyy"))
    mTmp = Replace(mTmp, "$4", Me.SituacionText)
    'mTmp = Replace(mTmp, "$5", Cstr(mTotalApuestas)
    mTmp = Replace(mTmp, "$6", IIf(Me.EsMultiple, "Es Multiple", "Diario"))
    'mTmp = Replace(mTmp, "$8", Participante
    mTmp = Replace(mTmp, "$9", Format(mCoste, FMT_IMPORTE))
    mTmp = Replace(mTmp, "$10", Format(mImportePremios, FMT_IMPORTE))
        
    Select Case mJuego
        Case bonoloto:
            mTmp = Replace(mTmp, "$2", "BL")
            mTmp = Replace(mTmp, "$12", mReintegroClave)
            
        Case LoteriaPrimitiva:
            mTmp = Replace(mTmp, "$2", "LP")
        
        Case gordoPrimitiva:
            mTmp = Replace(mTmp, "$2", "GP")
            
        Case Euromillones:
            mTmp = Replace(mTmp, "$2", "EU")
            mTmp = Replace(mTmp, "$7", mMillon)
            
    End Select
      
    ToString = mTmp
    
  On Error GoTo 0
ToString__CleanExit:
    Exit Function
            
ToString_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Boleto.ToString", ErrSource)
    Err.Raise ErrNumber, "Boleto.ToString", ErrDescription
End Function

'' *===========(EOF): Boleto.cls
