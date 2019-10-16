VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sorteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Clase que representa un Sorteo del Juego"
' *============================================================================*
' *
' *     Fichero    : Sorteo.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : 22-ago-2007 08:23
' *     Versión    : 1.0
' *     Propósito  : Clase que representa un sorteo de cada uno de los juegos
' *
' *============================================================================*
Option Explicit
Option Base 0
'--- Variables Publicas -------------------------------------------------------*
Public EntidadNegocio               As EntidadNegocio
Public Premios                      As Premios
'--- Constantes ---------------------------------------------------------------*
Private Const ERR_SORTEOEMPTY                   As Long = 1
Private Const ERR_SORTEOFALTAJUEGO              As Long = 2
Private Const ERR_SORTEOSINFECHA                As Long = 4
Private Const ERR_SORTEOSINCOMBINACION          As Long = 8
Private Const ERR_SORTEOCOMBIMALFALTAN          As Long = 16
Private Const ERR_SORTEOCOMBIMALSOBRAN          As Long = 32
Private Const ERR_SORTEONUMMAL6                 As Long = 64
Private Const ERR_SORTEOFALTAREINTEGRO          As Long = 128
Private Const ERR_SORTEOFALTACOMPLEMENTARIO     As Long = 256
Private Const ERR_SORTEOFALTAESTRELLAS          As Long = 512
Private Const ERR_SORTEOFALTACLAVE              As Long = 1024
Private Const ERR_SORTEOFALTANUMSORTEO          As Long = 2028
      
'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_ERRORESVALIDACION  As String = "#Sorteo, Hay errores de validación:"
Private Const MSG_SORTEOEMPTY        As String = "El sorteo no es válido para el juego que" & _
                                     " se define. Faltan datos de fecha, combin" & _
                                     "ación, reintegro, estrellas, etc."
Private Const MSG_SORTEOFALTAJUEGO           As String = "No se ha definido un juego para este sorteo."
Private Const MSG_SORTEOSINFECHA             As String = "No se ha definido la fecha del sorteo"
Private Const MSG_SORTEOSINCOMBINACION       As String = "No se ha definido la combinación del sorteo."
Private Const MSG_SORTEOCOMBIMALFALTAN       As String = "La combinación del sorteo está incompleta faltan números"
Private Const MSG_SORTEOCOMBIMALSOBRAN       As String = "La combinación del sorteo es errónea, sobran números"
Private Const MSG_SORTEONUMMAL6              As String = "Hay uno o varios números que no son correctos para el juego del sorteo"
Private Const MSG_SORTEOFALTAREINTEGRO       As String = "Falta el reintegro del sorteo."
Private Const MSG_SORTEOFALTACOMPLEMENTARIO  As String = "Falta el número complementario."
Private Const MSG_SORTEOFALTAESTRELLAS       As String = "Faltan las estrellas del sorteo."
Private Const MSG_SORTEOFALTACLAVE           As String = "Falta la clave del sorteo."
Private Const MSG_SORTEOFALTANUMSORTEO       As String = "Falta el número de serie del sorteo."

Private Const LIT_JUEGO7_49 As String = "#$1 Sorteo:$2 Fecha:$3 (Sem:$4, Dia:$5) Combinación:$6 C-$7 R-$8 Orden:$9"
Private Const LIT_JUEGO5_50 As String = "#$1 Sorteo:$2 Fecha:$3 (Sem:$4, Dia:$5) Combinación:$6 E-$7 Orden:$8"
Private Const LIT_JUEGO5_54 As String = "#$1 Sorteo:$2 Fecha:$3 (Sem:$4, Dia:$5) Combinación:$6 C-$7 Orden:$8"

Private Const LIT_ENORDEN = "En Orden de Salida"
Private Const LIT_NOTENORDEN = "Orden Ascendente"

'--- Variables Privadas -------------------------------------------------------*
Private mErrorValidacion            As Long         ' Control de errores de validación
Private m_eJuego                    As Juego        ' Juego al que pertenece el sorteo
Private m_dtFecha                   As Date         ' Fecha del sorteo
Private m_sDia                      As String       ' Dia de la semana
Private m_iComplementario           As Integer      ' Complementario
Private m_iReintegro                As Integer      ' Reintegro
Private m_objCombinacion            As Combinacion  ' Combinación ganadora
Private m_objEstrellas              As Combinacion  ' Estrellas
Private m_iSemana                   As Integer      ' Semana del año
Private m_bOrdenado                 As Boolean      ' Combinación en orden de aparición
Private m_NumSorteo                 As String       ' Numero de sorteo en el año

'--- Metodos Privados ---------------------------------------------------------*
'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_objCombinacion = New Combinacion
    Set EntidadNegocio = New EntidadNegocio
    Set m_objEstrellas = New Combinacion
    Set Premios = New Premios
End Sub

Private Sub Class_Terminate()
    Set m_objCombinacion = Nothing
    Set m_objEstrellas = Nothing
    Set EntidadNegocio = Nothing
    Set Premios = Nothing
End Sub


'--- Propiedades --------------------------------------------------------------*

'---------------------------------------------------------------------------------------
' Procedure : Juego
' Author    : CHARLY
' Date      : 09/05/2014
' Purpose   : Juego al que pertenece el sorteo
'---------------------------------------------------------------------------------------
'
Public Property Get Juego() As Juego
    Juego = m_eJuego
End Property

Public Property Let Juego(ByVal eJuego As Juego)
    m_eJuego = eJuego
End Property
'---------------------------------------------------------------------------------------
' Procedure : Año
' Author    : CHARLY
' Date      : 10/07/2018
' Purpose   : Año del sorteo
'---------------------------------------------------------------------------------------
'
Public Property Get Año() As Integer
    Año = CInt(Format(m_dtFecha, "yyyy"))
End Property
'---------------------------------------------------------------------------------------
' Procedure : NumSorteo
' Author    : CHARLY
' Date      : 10/07/2018
' Purpose   : Numero de serie del sorteo del juego
'---------------------------------------------------------------------------------------
'
Public Property Get NumSorteo() As String
    NumSorteo = m_NumSorteo
End Property

Public Property Let NumSorteo(vNewValue As String)
    m_NumSorteo = vNewValue
End Property

'---------------------------------------------------------------------------------------
' Procedure : ID
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   : Identificador único del sorteo
'---------------------------------------------------------------------------------------
'
Public Property Get Id() As Integer
    Id = Me.EntidadNegocio.Id
End Property

Public Property Let Id(vNewValue As Integer)
    Me.EntidadNegocio.Id = vNewValue
End Property

'---------------------------------------------------------------------------------------
' Procedure : Fecha
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   : Fecha de Sorteo
'---------------------------------------------------------------------------------------
'
Public Property Get Fecha() As Date
    Fecha = m_dtFecha
End Property

Public Property Let Fecha(ByVal dtFecha As Date)
    Dim strDiaSemana As String
    m_dtFecha = dtFecha
    strDiaSemana = Format(dtFecha, "dddd")

    Me.Dia = UCase(Left(strDiaSemana, 1))
    If Me.Dia = "M" And Len(strDiaSemana) > 6 Then
        Me.Dia = "X"
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : Texto
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   : Expresión de la combinación en formato texto
'---------------------------------------------------------------------------------------
'
Public Property Get Texto() As String
    Texto = m_objCombinacion.Texto
End Property

Public Property Let Texto(vNewValue As String)
    m_objCombinacion.Texto = CStr(vNewValue)
End Property

'---------------------------------------------------------------------------------------
' Procedure : Dia
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   : Abreviatura del día de la semana del sorteo
'---------------------------------------------------------------------------------------
'
Public Property Get Dia() As String
    Dia = m_sDia
End Property

Public Property Let Dia(ByVal sDia As String)
    m_sDia = sDia
End Property
'---------------------------------------------------------------------------------------
' Procedure : CombinacionGanadora
' Author    : CHARLY
' Date      : vie, 20/01/2012 23:33
' Purpose   : Conjunto de números que forman la combinación ganadora del sorteo
'---------------------------------------------------------------------------------------
'
Public Property Get CombinacionGanadora() As Combinacion
    Set CombinacionGanadora = m_objCombinacion
End Property

Public Property Set CombinacionGanadora(vNewValue As Combinacion)
    Set m_objCombinacion = vNewValue
End Property

'---------------------------------------------------------------------------------------
' Procedure : Complementario
' Author    : Charly
' Date      : 17/11/2013
' Purpose   : Número complementario si el juego es Bonoloto o Primitiva
'---------------------------------------------------------------------------------------
'
Public Property Get Complementario() As Integer
   Complementario = m_iComplementario
End Property
'
Public Property Let Complementario(vNewValue As Integer)
    m_iComplementario = vNewValue
End Property

'---------------------------------------------------------------------------------------
' Procedure : Reintegro
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   : Número del reintegro si el juego es Bonoloto o Primitiva
'---------------------------------------------------------------------------------------
'
Public Property Get Reintegro() As Integer
    Reintegro = m_iReintegro
End Property

Public Property Let Reintegro(ByVal iReintegro As Integer)
    m_iReintegro = iReintegro
End Property
'---------------------------------------------------------------------------------------
' Procedure : Clave
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   : Número de la clave si el juego es el Gordo de la Primitiva
'---------------------------------------------------------------------------------------
'
Public Property Get Clave() As Integer
    Clave = m_iReintegro
End Property

Public Property Let Clave(ByVal vNewValue As Integer)
    m_iReintegro = vNewValue
End Property

'---------------------------------------------------------------------------------------
' Procedure : Semana
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   : Numero de la semana dentro del año del sorteo.
'---------------------------------------------------------------------------------------
'
Public Property Get Semana() As Integer

   On Error GoTo Semana_Error

    m_iSemana = DatePart("ww", m_dtFecha, vbMonday, vbFirstFourDays)
    Semana = m_iSemana

   On Error GoTo 0
   Exit Property

Semana_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.Semana")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Ordenado
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   : Indicador de si los números de la combinación están
'             por orden de salida del bombo
'---------------------------------------------------------------------------------------
'
Public Property Get Ordenado() As Boolean
    Ordenado = m_bOrdenado
End Property

Public Property Let Ordenado(ByVal bOrdenado As Boolean)
    m_bOrdenado = bOrdenado
End Property

'---------------------------------------------------------------------------------------
' Procedure : Estrellas
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   : Conjunto de números que componen las estrellas para el juego Euromillones
'---------------------------------------------------------------------------------------
'
Public Property Get Estrellas() As Combinacion
    Set Estrellas = m_objEstrellas
End Property

Public Property Set Estrellas(objEstrellas As Combinacion)
    Set m_objEstrellas = objEstrellas
End Property

'---------------------------------------------------------------------------------------
' Procedure : EstrellaUno
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   : Número correspondiente a la primera estrella si el juego es Euromillones
'---------------------------------------------------------------------------------------
'
Public Property Get EstrellaUno() As Numero
    If m_objEstrellas.Count > 0 Then
        Set EstrellaUno = m_objEstrellas.Numeros.Item(1)
    Else
        Set EstrellaUno = New Numero
    End If
End Property

Public Property Set EstrellaUno(objEstrellaUno As Numero)
    Dim mNum As Numero
    Set mNum = New Numero
    Set mNum = objEstrellaUno
    If m_objEstrellas.Count = 0 Then
        m_objEstrellas.Add mNum
    Else
        Set m_objEstrellas.Numeros.Item(1) = mNum
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : EstrellaDos
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   : Número correspondiente a la segunda estrella si el juego es Euromillones
'---------------------------------------------------------------------------------------
'
Public Property Get EstrellaDos() As Numero
    If m_objEstrellas.Count > 1 Then
        Set EstrellaDos = m_objEstrellas.Numeros.Item(2)
    Else
        Set EstrellaDos = New Numero
    End If
End Property

Public Property Set EstrellaDos(objEstrellaDos As Numero)
    Dim mNum As Numero
    Set mNum = New Numero
    Set mNum = objEstrellaDos
    If m_objEstrellas.Count <= 1 Then
        m_objEstrellas.Add mNum
    Else
        Set m_objEstrellas.Numeros.Item(2) = mNum
    End If
End Property
'--- Metodos Publicos ---------------------------------------------------------*
'---------------------------------------------------------------------------------------
' Procedure : EsValido
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   : Valida el contenido del sorteo, y si es válido para el juego devuelve True
'---------------------------------------------------------------------------------------
'
Public Function EsValido() As Boolean
 On Error GoTo EsValido_Error
    '
    '   Inicializamos variables de control
    '
    mErrorValidacion = 0
    '
    '
    '
    If m_dtFecha = 0 _
    And m_objCombinacion.Count = 0 _
    And m_iReintegro = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOEMPTY
    End If
    '
    '
    If m_eJuego = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOFALTAJUEGO
    End If
    '
    '
    If m_dtFecha = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOSINFECHA
    End If
    '
    '
    If m_objCombinacion.Count = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOSINCOMBINACION
    End If
    '
    '
    If m_objCombinacion.Count < 5 And m_eJuego = Euromillones Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOCOMBIMALFALTAN
    End If
    If m_objCombinacion.Count < 6 And m_eJuego = bonoloto Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOCOMBIMALFALTAN
    End If
    If m_objCombinacion.Count < 6 And m_eJuego = LoteriaPrimitiva Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOCOMBIMALFALTAN
    End If
    If m_objCombinacion.Count < 6 And m_eJuego = gordoPrimitiva Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOCOMBIMALFALTAN
    End If
    '
    '
    If m_objCombinacion.Count > 5 And m_eJuego = Euromillones Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOCOMBIMALSOBRAN
    End If
    If m_objCombinacion.Count > 7 And m_eJuego = bonoloto Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOCOMBIMALSOBRAN
    End If
    If m_objCombinacion.Count > 7 And m_eJuego = LoteriaPrimitiva Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOCOMBIMALSOBRAN
    End If
    If m_objCombinacion.Count > 6 And m_eJuego = gordoPrimitiva Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOCOMBIMALSOBRAN
    End If
    '
    'ERR_SORTEONUMMAL TODO: Falta método EsValido en combinación (juego)
'    If m_objCombinacion.EsValida Then
'        mErrorValidacion = mErrorValidacion + ERR_SORTEONUMMAL
'    End If
    '
    '
    If m_iReintegro = -1 Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOFALTAREINTEGRO
    End If
    '
    '
    If m_iComplementario = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOFALTACOMPLEMENTARIO
    End If
    '
    '
    If m_objEstrellas.Count < 2 And m_eJuego = Euromillones Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOFALTAESTRELLAS
    End If
    '
    '
    If m_iReintegro = -1 And m_eJuego = gordoPrimitiva Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOFALTACLAVE
    End If
    '
    '
    If Len(m_NumSorteo) = 0 Then
        mErrorValidacion = mErrorValidacion + ERR_SORTEOFALTANUMSORTEO
    End If
    '
    '   Evaluamos el código de error
    '
    If mErrorValidacion = 0 Then
        EsValido = True
    Else
        EsValido = False
    End If
 On Error GoTo 0
   Exit Function

EsValido_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.EsValido")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetMensaje
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   : Devuelve los errores del sorteo, si no es válido
'---------------------------------------------------------------------------------------
'
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
        '   Si hay algun error inicializamos la cabecera del error
        '
        sResult = MSG_ERRORESVALIDACION
    End If
    '
    If (mErrorValidacion And ERR_SORTEOEMPTY) Then
        sResult = sResult & vbTab & MSG_SORTEOEMPTY & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_SORTEOFALTAJUEGO) Then
        sResult = sResult & vbTab & MSG_SORTEOFALTAJUEGO & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_SORTEOSINFECHA) Then
        sResult = sResult & vbTab & MSG_SORTEOSINFECHA & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_SORTEOSINCOMBINACION) Then
        sResult = sResult & vbTab & MSG_SORTEOSINCOMBINACION & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_SORTEOCOMBIMALFALTAN) Then
        sResult = sResult & vbTab & MSG_SORTEOCOMBIMALFALTAN & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_SORTEOCOMBIMALSOBRAN) Then
        sResult = sResult & vbTab & MSG_SORTEOCOMBIMALSOBRAN & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_SORTEONUMMAL6) Then
        sResult = sResult & vbTab & MSG_SORTEONUMMAL6 & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_SORTEOFALTAREINTEGRO) Then
        sResult = sResult & vbTab & MSG_SORTEOFALTAREINTEGRO & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_SORTEOFALTACOMPLEMENTARIO) Then
        sResult = sResult & vbTab & MSG_SORTEOFALTACOMPLEMENTARIO & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_SORTEOFALTAESTRELLAS) Then
        sResult = sResult & vbTab & MSG_SORTEOFALTAESTRELLAS & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_SORTEOFALTACLAVE) Then
        sResult = sResult & vbTab & MSG_SORTEOFALTACLAVE & vbCrLf
    End If
    '
    If (mErrorValidacion And ERR_SORTEOFALTANUMSORTEO) Then
        sResult = sResult & vbTab & MSG_SORTEOFALTANUMSORTEO & vbCrLf
    End If
    '
    '   Devolvemos el mensaje
    '
    GetMensaje = sResult

 On Error GoTo 0
   Exit Function

GetMensaje_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.GetMensaje")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : ToString
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   : Expresión del Sorteo en formato texto
'---------------------------------------------------------------------------------------
'
Public Function ToString() As String
    Dim sResult As String
    

   On Error GoTo ToString_Error
    '
    '   TODO: Incluir el nombre del juego o letras
    '
    Select Case m_eJuego
    Case bonoloto, LoteriaPrimitiva:
        sResult = Replace(LIT_JUEGO7_49, "$1", Me.Id)
        sResult = Replace(sResult, "$2", Me.NumSorteo)
        sResult = Replace(sResult, "$3", Me.Fecha)
        sResult = Replace(sResult, "$4", Me.Semana)
        sResult = Replace(sResult, "$5", Me.Dia)
        sResult = Replace(sResult, "$6", Me.CombinacionGanadora.ToString(True))
        sResult = Replace(sResult, "$7", Me.Complementario)
        sResult = Replace(sResult, "$8", Me.Reintegro)
        sResult = Replace(sResult, "$9", IIf(Me.Ordenado, LIT_ENORDEN, LIT_NOTENORDEN))
       
    Case Euromillones
        sResult = Replace(LIT_JUEGO5_50, "$1", Me.Id)
        sResult = Replace(sResult, "$2", Me.NumSorteo)
        sResult = Replace(sResult, "$3", Me.Fecha)
        sResult = Replace(sResult, "$4", Me.Semana)
        sResult = Replace(sResult, "$5", Me.Dia)
        sResult = Replace(sResult, "$6", Me.CombinacionGanadora.ToString(True))
        sResult = Replace(sResult, "$7", Me.Estrellas.ToString(True))
        sResult = Replace(sResult, "$8", IIf(Me.Ordenado, LIT_ENORDEN, LIT_NOTENORDEN))
        
    Case gordoPrimitiva
        sResult = Replace(LIT_JUEGO5_54, "$1", Me.Id)
        sResult = Replace(sResult, "$2", Me.NumSorteo)
        sResult = Replace(sResult, "$3", Me.Fecha)
        sResult = Replace(sResult, "$4", Me.Semana)
        sResult = Replace(sResult, "$5", Me.Dia)
        sResult = Replace(sResult, "$6", Me.CombinacionGanadora.ToString(True))
        sResult = Replace(sResult, "$7", Me.Clave)
        sResult = Replace(sResult, "$8", IIf(Me.Ordenado, LIT_ENORDEN, LIT_NOTENORDEN))
    
    Case Else
        sResult = MSG_SORTEOFALTAJUEGO
    End Select

    ToString = sResult

   On Error GoTo 0
   Exit Function

ToString_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.ToString")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : Constructor
' Fecha          : vi., 03/may/2019 23:08:29
' Propósito      : Construir la clase sorteo a partir de un rango de datos
' Parámetros     : Rango de información de la tabla sorteo
'------------------------------------------------------------------------------*
Public Sub Constructor(mRange As Range)
    Dim i       As Integer
    Dim oNum    As Numero
  On Error GoTo Constructor_Error
    '
    '   Si no tenemos definido el juego Error
    '
    If Me.Juego = SinJuego Then
        Err.Raise ERR_SORTEOFALTAJUEGO, "Sorteo.Constructor", MSG_SORTEOFALTAJUEGO
    End If
    '
    '   Cargamos datos comunes a todos los juegos
    '
    Me.EntidadNegocio.ClassStorage = True
    Me.Id = mRange.Cells(1, 1).Value                    ' ID
    Me.NumSorteo = mRange.Cells(1, 2).Value             ' Sorteo
    Me.Fecha = mRange.Cells(1, 3).Value                 ' Fecha
    Me.Ordenado = IIf(mRange.Cells(1, 6).Value = "Si", True, False)
    '
    '   TODO: Definir premios por defecto
    '
    '
    '   Datos especificos de cada juego
    '
    Select Case Juego
        'N   Sorteo  Fecha   Dia Sem Orden   N1  N2  N3  N4  N5  N6  C   R
        Case bonoloto:
            For i = 1 To 6
                Set oNum = New Numero
                oNum.Valor = mRange.Cells(1, 6 + i).Value    ' Numero
                oNum.Orden = i
                Me.CombinacionGanadora.Add oNum
            Next i
            Me.Complementario = mRange.Cells(1, 13).Value    ' Complementario
            Me.Reintegro = mRange.Cells(1, 14).Value         ' Reintegro
                    
        
        'N   Sorteo  Fecha   Dia Sem Orden   N1  N2  N3  N4  N5  N6  C   R
        Case LoteriaPrimitiva:
            For i = 1 To 6
                Set oNum = New Numero
                oNum.Valor = mRange.Cells(1, 6 + i).Value    ' Numero
                oNum.Orden = i
                Me.CombinacionGanadora.Add oNum
            Next i
            Me.Complementario = mRange.Cells(1, 13).Value    ' Complementario
            Me.Reintegro = mRange.Cells(1, 14).Value         ' Reintegro
                    
        'N   Sorteo  Fecha   Dia Sem Orden   N1  N2  N3  N4  N5  E1  E2
        Case Euromillones:
            For i = 1 To 5
                Set oNum = New Numero
                oNum.Valor = mRange.Cells(1, 6 + i).Value    ' Numero
                oNum.Orden = i
                Me.CombinacionGanadora.Add oNum
            Next i
            Set oNum = New Numero
            oNum.Valor = mRange.Cells(1, 12).Value
            oNum.Orden = 1
            Set Me.EstrellaUno = oNum                       ' Estrella uno
            Set oNum = New Numero
            oNum.Valor = mRange.Cells(1, 13).Value
            oNum.Orden = 2
            Set Me.EstrellaDos = oNum                       ' Estrella dos
                    
                
        'N   Sorteo  Fecha   Dia Sem Orden   N1  N2  N3  N4  N5  R
        Case gordoPrimitiva:
            For i = 1 To 5
                Set oNum = New Numero
                oNum.Valor = mRange.Cells(1, 6 + i).Value   ' Numero
                oNum.Orden = i
                Me.CombinacionGanadora.Add oNum
            Next i
            Me.Clave = mRange.Cells(1, 12).Value            ' Clave
    End Select
    '
    '   Parametros de entidad
    '
    Me.EntidadNegocio.IsNew = False
    Me.EntidadNegocio.ClassStorage = False
  
  On Error GoTo 0
Constructor__CleanExit:
    Exit Sub
            
Constructor_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Sorteo.Constructor", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Sorteo.Constructor", ErrDescription
End Sub
''---------------------------------------------------------------------------------------
'' Procedure : Constructor
'' Author    : CHARLY
'' Date      : 17/04/2014
'' Purpose   :
''---------------------------------------------------------------------------------------
''
Public Sub ConstructorV1_0(Fila As Range)
    Dim mCol    As Collection
    Dim i       As Integer
    Dim m_num   As Numero
    Dim mStrKey As String
    
   On Error GoTo Constructor_Error

    Set mCol = New Collection
    With Fila
        Me.Id = .Cells(1, 1).Value
        Me.Fecha = CDate(.Cells(1, 3).Value)
        For i = 0 To 6
            If IsNumeric(.Cells(1, 6 + i).Value) _
            And Not IsEmpty(.Cells(1, 6 + i).Value) Then
                If (.Cells(1, 6 + i).Value) < 50 Then
                    Set m_num = New Numero
                    m_num.Valor = .Cells(1, 6 + i).Value
                    mStrKey = CStr(m_num.Valor)
                    If (mCol.Count <> 0) Then
                        mCol.Add m_num, mStrKey, , mCol.Count
                    Else
                        mCol.Add m_num, mStrKey
                    End If
                End If
            End If
        Next i
        m_iComplementario = .Cells(1, 12).Value
        m_iReintegro = .Cells(1, 13).Value
        Set m_objCombinacion.Numeros = mCol
'        m_eTipoJuego = LP_LB_6_49
        m_eJuego = JUEGO_DEFECTO
    End With

   On Error GoTo 0
   Exit Sub

Constructor_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.Constructor")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
'' *===========(EOF): Sorteo.cls
