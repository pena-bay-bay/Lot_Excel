VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Apuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Apuesta.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sá., 10/ago/2019 20:06:26
' *     Versión    : 1.0
' *     Propósito  :
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'   TODO: Al validar la apuesta, se puede calcular el coste utilizando
'         Tarifas y TotalApuestas
'
'
'--- Variables Privadas -------------------------------------------------------*
Private mIdBoleto As Integer
Private mJuego As Juego
Private mFechaSorteo As Date
Private mEsMultiple As Boolean
Private mCoste As Currency
Private mImportePremios As Currency
Private mNumeros As Combinacion
Private mEstrellas As Combinacion

'--- Constantes ---------------------------------------------------------------*
Private mErrorValidacion                    As Long       ' Control de errores de validación
Private Const ERR_JUEGOOBLIGATORIO          As Long = 1   ' Juego Obligatorio
Private Const ERR_FECHAOBLIGATORIA          As Long = 2   ' Fecha Obligatoria
Private Const ERR_NUMEROSOBLIGATORIO        As Long = 4   ' Numeros de combinación obligatorios
Private Const ERR_ESTRLLASOBLIGATORIO       As Long = 8   ' Estrellas Obligatorias
Private Const ERR_COSTEOBLIGATORIO          As Long = 16  ' Coste Obligatorio
Private Const ERR_APTSIMPLECOMBINACIONMAL   As Long = 32  ' Boleto diario => numeros en combinaciones 6
Private Const ERR_APTMULTICOMBINACIONMAL    As Long = 64  ' Boleto multiple => depende del juego
Private Const ERR_COMBINACIONNOVALIDA       As Long = 128 ' Combinación no válida
Private Const ERR_ESTRELLASNOVALIDA         As Long = 256 ' Estrellas no válida
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
Private Const MSG_ERRORESVALIDACION         As String = "Los datos de la apuesta no cumple las siguientes validaciones:" & vbCrLf
Private Const MSG_JUEGOOBLIGATORIO          As String = "* El juego es un dato obligatorio."
Private Const MSG_FECHAOBLIGATORIA          As String = "* La fecha de la apuesta es obligatoria."
Private Const MSG_NUMEROSOBLIGATORIO        As String = "* Los numeros de la combinación son obligatorios."
Private Const MSG_ESTRLLASOBLIGATORIO       As String = "* Las estrellas de la apuesta son obligatorios."
Private Const MSG_COSTEOBLIGATORIO          As String = "* El coste de la apuesta es obligatorio."
Private Const MSG_APTSIMPLECOMBINACIONMAL   As String = "* La combinación es erronea para apuestas simples."
Private Const MSG_APTMULTICOMBINACIONMAL    As String = "* La combinación es erronea para apuestas multiples."
Private Const MSG_COMBINACIONNOVALIDA       As String = "* La combinación no es valida:"
Private Const MSG_ESTRELLASNOVALIDA         As String = "* Las estrellas no son validas:"

'--- Propiedades --------------------------------------------------------------*
Public EntidadNegocio As EntidadNegocio

Public Property Get Id() As Integer
    Id = EntidadNegocio.Id
End Property

Public Property Let Id(ByVal vNewValue As Integer)
    EntidadNegocio.Id = vNewValue
End Property
Public Property Get IdBoleto() As Integer
    IdBoleto = mIdBoleto
End Property

Public Property Let IdBoleto(ByVal vNewValue As Integer)
    mIdBoleto = vNewValue
End Property

Public Property Get Juego() As Juego
    Juego = mJuego
End Property

Public Property Let Juego(ByVal vNewValue As Juego)
    mJuego = vNewValue
End Property

Public Property Get FechaSorteo() As Date
    FechaSorteo = mFechaSorteo
End Property

Public Property Let FechaSorteo(ByVal vNewValue As Date)
    mFechaSorteo = vNewValue
End Property

Public Property Get EsMultiple() As Boolean
    EsMultiple = mEsMultiple
End Property

Public Property Let EsMultiple(ByVal vNewValue As Boolean)
    mEsMultiple = vNewValue
End Property

Public Property Get Coste() As Currency
    Coste = mCoste
End Property

Public Property Let Coste(ByVal vNewValue As Currency)
    mCoste = vNewValue
End Property

Public Property Get ImportePremios() As Currency
    ImportePremios = mImportePremios
End Property

Public Property Let ImportePremios(ByVal vNewValue As Currency)
    mImportePremios = vNewValue
End Property

Public Property Get Numeros() As Combinacion
    Set Numeros = mNumeros
End Property

Public Property Set Numeros(ByVal vNewValue As Combinacion)
    Set mNumeros = vNewValue
End Property

Public Property Get Estrellas() As Combinacion
    Set Estrellas = mEstrellas
End Property

Public Property Set Estrellas(ByVal vNewValue As Combinacion)
    Set mEstrellas = vNewValue
End Property

Public Property Get TotalApuestas() As Integer
    TotalApuestas = 0
    Select Case mJuego
        Case bonoloto, LoteriaPrimitiva:
            Select Case mNumeros.Count
                Case 5: TotalApuestas = 44
                Case 6: TotalApuestas = 1
                Case 7: TotalApuestas = 7
                Case 8: TotalApuestas = 28
                Case 9: TotalApuestas = 84
                Case 10: TotalApuestas = 210
                Case 11: TotalApuestas = 462
                Case Else: TotalApuestas = 0
            End Select
        
        Case Euromillones
            Select Case mEstrellas.Count
                Case 2:
                    Select Case mNumeros.Count
                        Case 5: TotalApuestas = 1
                        Case 6: TotalApuestas = 6
                        Case 7: TotalApuestas = 21
                        Case 8: TotalApuestas = 56
                        Case 9: TotalApuestas = 126
                        Case 10: TotalApuestas = 252
                    End Select
                    
                Case 3:
                    Select Case mNumeros.Count
                        Case 5: TotalApuestas = 3
                        Case 6: TotalApuestas = 18
                        Case 7: TotalApuestas = 63
                        Case 8: TotalApuestas = 168
                        Case 9: TotalApuestas = 378
                        Case 10: TotalApuestas = 756
                    End Select
                    
                Case 4:
                    Select Case mNumeros.Count
                        Case 5: TotalApuestas = 6
                        Case 6: TotalApuestas = 36
                        Case 7: TotalApuestas = 126
                        Case 8: TotalApuestas = 336
                        Case 9: TotalApuestas = 756
                        Case 10: TotalApuestas = 1512
                    End Select
                
                Case 5:
                    Select Case mNumeros.Count
                        Case 5: TotalApuestas = 10
                        Case 6: TotalApuestas = 60
                        Case 7: TotalApuestas = 210
                        Case 8: TotalApuestas = 560
                        Case 9: TotalApuestas = 1260
                        Case 10: TotalApuestas = 2520
                    End Select
                    
                Case Else: TotalApuestas = 0
            End Select
                
        Case gordoPrimitiva
            Select Case mNumeros.Count
                Case 6: TotalApuestas = 6
                Case 7: TotalApuestas = 21
                Case 8: TotalApuestas = 56
                Case 9: TotalApuestas = 126
                Case 10: TotalApuestas = 252
                Case 11: TotalApuestas = 462
                Case Else: TotalApuestas = 0
            End Select
    End Select

End Property
'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set EntidadNegocio = New EntidadNegocio
    mJuego = SinJuego
    mEsMultiple = False
    mCoste = 0
    mImportePremios = 0
    Set mNumeros = New Combinacion
    Set mEstrellas = New Combinacion
End Sub

'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : EsValido
' Fecha          : lu., 19/ago/2019 20:30:36
' Propósito      : Comprobar la validación de la apuesta
' Retorno        : True si la apuesta es valida
'------------------------------------------------------------------------------*
Public Function EsValido() As Boolean
  On Error GoTo EsValido_Error
    '
    mErrorValidacion = 0
    '
    '
    If mJuego = SinJuego Then
        mErrorValidacion = mErrorValidacion + ERR_JUEGOOBLIGATORIO
    End If
    '
    '
    If mFechaSorteo = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_FECHAOBLIGATORIA
    End If
    '
    '
    If mNumeros.Count = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_NUMEROSOBLIGATORIO
    End If
    '
    '
    If mEstrellas.Count = 0 And mJuego = Euromillones Then
        mErrorValidacion = mErrorValidacion + ERR_ESTRLLASOBLIGATORIO
    End If
    '
    '
    If mCoste = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_COSTEOBLIGATORIO
    End If
    '
    '
    If Not mEsMultiple Then
        Select Case mJuego
            Case bonoloto, LoteriaPrimitiva:
                If mNumeros.Count <> 6 Then
                    mErrorValidacion = mErrorValidacion + ERR_APTSIMPLECOMBINACIONMAL
                End If
                
            Case gordoPrimitiva:
                If mNumeros.Count <> 5 Then
                    mErrorValidacion = mErrorValidacion + ERR_APTSIMPLECOMBINACIONMAL
                End If
                
            Case Euromillones:
                If mNumeros.Count <> 5 And mEstrellas.Count <> 2 Then
                    mErrorValidacion = mErrorValidacion + ERR_APTSIMPLECOMBINACIONMAL
                End If
        End Select
    Else
        Select Case mJuego
            Case bonoloto, LoteriaPrimitiva:
                If (mNumeros.Count < 5) _
                Or (mNumeros.Count = 6) _
                Or (mNumeros.Count > 11) Then
                    mErrorValidacion = mErrorValidacion + ERR_APTMULTICOMBINACIONMAL
                End If
            Case gordoPrimitiva:
                If (mNumeros.Count < 5) _
                Or (mNumeros.Count = 5) _
                Or (mNumeros.Count > 11) Then
                    mErrorValidacion = mErrorValidacion + ERR_APTMULTICOMBINACIONMAL
                End If
                
            Case Euromillones:
                If (mNumeros.Count <= 5) _
                Or (mNumeros.Count > 10) _
                Or (mEstrellas.Count <= 2) _
                Or (mEstrellas.Count > 5) Then
                    mErrorValidacion = mErrorValidacion + ERR_APTSIMPLECOMBINACIONMAL
                End If
        End Select
    End If
    '
    '
    If Not mNumeros.EsValido(mJuego, True) Then
            mErrorValidacion = mErrorValidacion + ERR_COMBINACIONNOVALIDA
    End If
    '
    '
    If mEstrellas.EsValido(mJuego, True) And mJuego = Euromillones Then
            mErrorValidacion = mErrorValidacion + ERR_ESTRELLASNOVALIDA
    End If
    '
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
    Call HandleException(ErrNumber, ErrDescription, "Apuesta.EsValido", ErrSource)
    Err.Raise ErrNumber, "Apuesta.EsValido", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetMensaje
' Fecha          : lu., 19/ago/2019 20:32:43
' Propósito      : Devolver el error de validación de la Apuesta
' Parámetros     :
' Retorno        :
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

    If (mErrorValidacion And ERR_NUMEROSOBLIGATORIO) Then
        sResult = sResult & vbTab & MSG_NUMEROSOBLIGATORIO & vbCrLf
    End If

    If (mErrorValidacion And ERR_ESTRLLASOBLIGATORIO) Then
        sResult = sResult & vbTab & MSG_ESTRLLASOBLIGATORIO & vbCrLf
    End If

    If (mErrorValidacion And ERR_COSTEOBLIGATORIO) Then
        sResult = sResult & vbTab & MSG_COSTEOBLIGATORIO & vbCrLf
    End If

    If (mErrorValidacion And ERR_APTSIMPLECOMBINACIONMAL) Then
        sResult = sResult & vbTab & MSG_APTSIMPLECOMBINACIONMAL & vbCrLf
    End If

    If (mErrorValidacion And ERR_APTMULTICOMBINACIONMAL) Then
        sResult = sResult & vbTab & MSG_APTMULTICOMBINACIONMAL & vbCrLf
    End If
    
    If (mErrorValidacion And ERR_COMBINACIONNOVALIDA) Then
        sResult = sResult & vbTab & MSG_COMBINACIONNOVALIDA & mNumeros.GetMensaje & vbCrLf
    End If
    
    If (mErrorValidacion And ERR_ESTRELLASNOVALIDA) Then
        sResult = sResult & vbTab & MSG_ESTRELLASNOVALIDA & mNumeros.GetMensaje & vbCrLf
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
    Call HandleException(ErrNumber, ErrDescription, "Apuesta.GetMensaje", ErrSource)
    Err.Raise ErrNumber, "Apuesta.GetMensaje", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : lu., 19/ago/2019 18:09:44
' Propósito      : Visualizar las propiedades de una apuesta
' Retorno        : Cadena de texto con la apuesta
'------------------------------------------------------------------------------*
Public Function ToString() As String
    Dim mTemplate As String
    Dim mTmp As String
    
  On Error GoTo ToString_Error
    '
    '  Id IdBoleto Juego FechaSorteo EsMultiple N1 N2 N3 N4 N5 N6 N7 N8 N9 N10 N11 E1 E2 E3 E4 E5 Coste ImportePremios
    '
    If mJuego = Euromillones Then
        mTemplate = "#$1, Boleto:#$2, $3, $4, $5, Combinación: $6, Estrellas:$9 Coste:$7, Premios:$8"
    Else
        mTemplate = "#$1, Boleto:#$2, $3, $4, $5, Combinación: $6, Coste:$7, Premios:$8"
    End If
    mTmp = Replace(mTemplate, "$1", CStr(Id))
    mTmp = Replace(mTmp, "$2", CStr(IdBoleto))
    mTmp = Replace(mTmp, "$4", Format(mFechaSorteo, "dd/mm/yyyy"))
    mTmp = Replace(mTmp, "$5", IIf(Me.EsMultiple, "Es Multiple", "Diario"))
    mTmp = Replace(mTmp, "$6", mNumeros.ToString(True))
    mTmp = Replace(mTmp, "$7", Format(mCoste, FMT_IMPORTE))
    mTmp = Replace(mTmp, "$8", Format(mImportePremios, FMT_IMPORTE))
        
    Select Case mJuego
        Case bonoloto:
            mTmp = Replace(mTmp, "$3", "BL")
            
        Case LoteriaPrimitiva:
            mTmp = Replace(mTmp, "$3", "LP")
        
        Case gordoPrimitiva:
            mTmp = Replace(mTmp, "$3", "GP")
            
        Case Euromillones:
            mTmp = Replace(mTmp, "$3", "EU")
            mTmp = Replace(mTmp, "$9", mEstrellas.ToString(True))
            
    End Select
      
    ToString = mTmp
    
  On Error GoTo 0
ToString__CleanExit:
    Exit Function
            
ToString_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Apuesta.ToString", ErrSource)
    Err.Raise ErrNumber, "Apuesta.ToString", ErrDescription
End Function
'' *===========(EOF): <<nombre fichero>>


