VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BolaV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Bola.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creaci�n   : dom, 23/Abr/2017 09:32:00
' *     Versi�n    : 2.0
' *     Prop�sito  : Clase estructura que contiene los datos estadisticos
' *                  de un n�mero
' *
' *============================================================================*
Option Explicit
Option Base 0
'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'
'
'--- Variables Internas -------------------------------------------------------*
Private mApariciones As Integer             ' Numero de apariciones del Numero
                                            ' en la muestra
Private mAusencias As Integer               ' Numero de ausencias desde el
                                            ' �ltimo sorteo de aparicion
Private mProbabilidad As Double             ' Probabilidad de aparici�n
Private mFrecuencia As Double               ' Probabilidad de la frecuencia
Private mTiempoMedio As Double              ' Probabilidad del tiempo medio en
                                            ' aparecer
Private mRegistroAnalisis As Integer        ' Numero del registros de an�lisis
                                            ' para calcular las ausencias
Private mNumero As Numero                   ' Numero que representa
Private mFechasAparicion  As Collection     ' Colecci�n de apariciones
Private mFrecuencias As Collection          ' Colecci�n de frecuencias
Private mTiempos  As Collection             ' Colecci�n de tiempos de aparici�n
                                            ' Rango de frecuencias para PrimiyBonoloto
Private Const mRgFrecuenciasBonoloto As String = "7,14,28,35,42,50,100,150,999"
                                            ' Rango de frecuencias para Euromillon
Private Const mRgFrecuenciasEuromillon As String = "5,10,15,20,25,30,50,100,999"
Private Const mRgFrecuenciasEstrellas As String = "2,4,6,8,10,12,14,16,18,20,40,999"
'
'-------------------------------------------+
Public Juego As Juego                       ' Juego al que hace referencia la
                                            ' Bola
Public FechaAnalisis As Date                ' Fecha de analisis de la muestra
Public RegistroAnalisis As Integer          ' Registro de la fecha de analisis
Public TotalNumeros As Integer              ' Total Numeros de la muestra
Public ColorFrecuencia As Integer           ' Color de la bola en la muestra por
                                            ' la probabilidad de frecuencia
Public ColorProbabilidad As Integer         ' Color de la bola en la muestra por
                                            ' la probabilidad
Public ColorTiempoMedio As Integer          ' Color de la bola en la muestra por
                                            ' la probabilidad del tiempo medio
Public TipoBola As Integer                  ' Tipo de bola del sorteo: 1=Numero,
                                            ' 2-Estrella, 3-Numero Clave
                                            
                                            
Public DesviacionTiempoMedio As Double      ' Sigma para el tiempo medio
Public FechaUltimaAparicion As Date         ' Fecha de �ltima aparici�n del n�mero
Public Frecuencias As Collection            ' Colecci�n de tiempos de aparicion
                                            ' clasificados por rangos
Public MaximoTiempo As Integer              ' Tiempo m�ximo en aparecer
Public Mediana As Double                    '
Public MinimoTiempo As Integer
Public Moda As Integer
Public ProbabilidadFrecuencia As Double
Public ProbabilidadTiempo As Double
Public ProximaFechaAparicion As Date
Public RegistroAparicion As Integer
Public Tendencia As String
Public TiempoMedio As Double
Public TipoAusencia As String
Public UltimoRegistro As Integer
'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Propiedad      : Numero
' Fecha          : 23/Abr/2017
' Prop�sito      : Numero del que se obtiene la estadistica
'------------------------------------------------------------------------------*
'
Public Property Get Numero() As Numero
    Set Numero = mNumero
End Property

Public Property Set Numero(ByVal vNewValue As Numero)
    Set mNumero = vNewValue
End Property
'------------------------------------------------------------------------------*
' Propiedad      : Apariciones
' Fecha          : 23/Abr/2017
' Prop�sito      : Numero de apariciones del n�mero en la muestra
'------------------------------------------------------------------------------*
'
Public Property Get Apariciones() As Integer
    If mApariciones = -1 Then
        mApariciones = mFechasAparicion.Count
    End If
    Apariciones = mApariciones
End Property

'------------------------------------------------------------------------------*
' Propiedad      : Ausencias
' Fecha          : 23/Abr/2017
' Prop�sito      : Numero del que se obtiene la estadistica
'------------------------------------------------------------------------------*
'
Public Property Get Ausencias() As Integer
    Dim mTupla As New TuplaAparicion
    
    If mAusencias = 0 And Apariciones > 0 Then
        '
        '   Obtenemos la �ltima aparici�n
        '
        Set mTupla = mFechasAparicion.Item(Apariciones)
        '
        ' calculamos el Numero de bolas entre esta y el an�lisis
        '
        mAusencias = getTiempoIntervalo(mTupla.NumeroRegistro, mTupla.OrdenAparicion, RegistroAnalisis, 0)
    End If
    Ausencias = mAusencias
End Property
Public Property Let Ausencias(ByVal vNewValue As Integer)
    '
    '   Si no hay apariciones podemos guardar las ausencias de una muestra
    '   anual
    '
    If mAusencias = 0 And mFechasAparicion.Count = 0 Then
        mAusencias = vNewValue
    End If
End Property

'------------------------------------------------------------------------------*
' Propiedad      : FechasAparicion
' Fecha          : 23/Abr/2017
' Prop�sito      : Colecci�n de fechas en la que aparece el n�mero
'------------------------------------------------------------------------------*
'
Public Property Get FechasAparicion() As Collection
    Set FechasAparicion = mFechasAparicion
End Property

Public Property Set FechasAparicion(ByVal vNewValue As Collection)
    Set mFechasAparicion = vNewValue
End Property
'------------------------------------------------------------------------------*
' Propiedad      : Probabilidad
' Fecha          : 02/May/2017
' Prop�sito      : Probabilidad del n�mero en la muestra
'------------------------------------------------------------------------------*
'
Public Property Get Probabilidad() As Double
    If mProbabilidad = -1 Then
        '
        '   Calculamos la probabilidad de aparicion
        '
        mProbabilidad = Apariciones / TotalNumeros
    End If
    Probabilidad = mProbabilidad
End Property


'--- Metodos ------------------------------------------------------------------*
'
'
Private Sub Class_Initialize()
    mApariciones = -1
    mProbabilidad = -1
    mFrecuencia = -1
    mTiempoMedio = -1
    mRegistroAnalisis = 0
    Set mNumero = New Numero
    Set mFechasAparicion = New Collection
    Set mTiempos = New Collection
    Set Frecuencias = New Collection
End Sub
'
'
'
Private Sub Class_Terminate()
    Set mNumero = Nothing
    Set mFechasAparicion = Nothing
    Set mTiempos = Nothing
    Set Frecuencias = Nothing
End Sub

'------------------------------------------------------------------------------*
' Propiedad      : Add
' Fecha          : 23/Abr/2017
' Prop�sito      : Agrega una aparicion del Numero a la colecci�n de fechas
'------------------------------------------------------------------------------*
'
Public Sub Add(ByVal vNewValue As TuplaAparicion)
  On Error GoTo Add_Error
    mFechasAparicion.Add vNewValue
  
  On Error GoTo 0
Add__CleanExit:
  Exit Sub
            
Add_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Bola.Add", ErrSource)
        '   Lanza el Error
    Err.Raise ErrNumber, "Bola.Add", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Actualizar
' Fecha          : 01/May/2017
' Prop�sito      : Proceso de c�lculo de las estadisticas
'------------------------------------------------------------------------------*
'
Public Sub Actualizar()
    Dim mOrdenSup As Integer            ' Orden de la aparici�n actual
    Dim mOrdenInf As Integer            ' Orden de la aparici�n anterior
    Dim mRegSup As Integer              ' Registro Actual
    Dim mRegInf As Integer              ' Registro Anterior
    Dim mTmAparicion As Integer         ' Tiempo de aparici�n
    Dim mTupla As TuplaAparicion        ' Registro de aparici�n de un n�mero
    Dim mSumTiempos As Integer          ' Sumatorio de tiempos de aparicion
    Dim mTm As Variant                  ' Tiempo de aparicion
    
  On Error GoTo Actualizar_Error
    '
    '   Inicializamos la colecci�n de tiempos de aparici�n
    '
    mOrdenSup = 0: mRegSup = 0
    mOrdenInf = 0: mRegInf = 0: mTmAparicion = 0
    Set mTiempos = New Collection
    '
    '   Si no hay aparciones salimos de la rutina
    '
    If mFechasAparicion.Count = 0 Then
        Exit Sub
    End If
    '
    ' Para cada tupla en la colecci�n
    '
    For Each mTupla In mFechasAparicion
        '
        ' Si es la primera tupla no tenemos datos
        ' anteriores se inicializan datos anteriores
        '
        If mOrdenInf = 0 And mOrdenSup = 0 Then
            mOrdenInf = mTupla.OrdenAparicion
            mRegInf = mTupla.NumeroRegistro
        Else
            '
            ' Obtenemos datos del registro actual
            '
            mOrdenSup = mTupla.OrdenAparicion
            mRegSup = mTupla.NumeroRegistro
            '
            ' Calculamos tiempo de aparicion
            '
            mTmAparicion = getTiempoIntervalo(mRegInf, mOrdenInf, mRegSup, mOrdenSup)
            '
            ' Agregamos tiempo aparici�n a la coleccion
            '
            mTiempos.Add mTmAparicion
            '
            ' trasladamos registro actual a registro anterior
            '
            mOrdenInf = mOrdenSup
            mRegInf = mRegSup
        End If
    Next mTupla
    '
    ' Calculamos la media de apariciones, m�ximo y Minimo
    '
    TiempoMedio = 0:   MaximoTiempo = 0:    MinimoTiempo = 9999
    mSumTiempos = 0
    '
    ' Para cada aparici�n en e
    '
    For Each mTm In mTiempos
        '
        '  Acumulamos los tiempos de aparici�n
        '
        mSumTiempos = mSumTiempos + mTm
        '
        '  Comprobamos que es mayor que el m�ximo
        '
        If mTm > MaximoTiempo Then
            MaximoTiempo = mTm
        End If
        '
        '  Comprobamos que es inferior al minimo
        '
        If mTm < MinimoTiempo Then
            MinimoTiempo = mTm
        End If
    Next mTm
    '
    '   Calculamos la media
    '
    TiempoMedio = mSumTiempos / mTiempos.Count
    '
    '   Calculamos las frecuencias
    '
    
    '
    '
    ' Err.Raise ERR_TODO, "Bola.Actualizar", MSG_TODO
    '
  On Error GoTo 0
Actualizar__CleanExit:
    Exit Sub
            
Actualizar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Bola.Actualizar", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Bola.Actualizar", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Funcion        : getTiempoIntervalo
' Fecha          : 02/May/2017
' Prop�sito      : Obtener el intervalo de apariciones entre dos apariciones
'                  de un n�mero
' Par�metros     : Registro del sorteo inicial,
'                  Orden en la combinaci�n del n�mero Inicial
'                  Registro del sorteo Final,
'                  Orden en la combinaci�n del n�mero Final
'
'------------------------------------------------------------------------------*
'
Private Function getTiempoIntervalo(vRegInf As Integer, vOrdInf As Integer, _
                                    vRegSup As Integer, vOrdSup As Integer) As Integer
    Dim mTmAparicion As Integer
    
  On Error GoTo getTiempoIntervalo_Error
    '
    '   Inicializamos el tiempo de aparici�n
    '
    mTmAparicion = 0
    '
    ' Calculamos tiempo de aparicion
    '
    Select Case Juego
      Case LoteriaPrimitiva, Bonoloto
           mTmAparicion = (vRegSup - vRegInf - 1) * 7 ' distancia entre registros * 7 Numeros
           mTmAparicion = mTmAparicion + (7 - vOrdInf) + vOrdSup
           
      Case Euromillones
           '
           ' Si la bola es un Numero o es una Estrella
           '
           If TipoBola = 1 Then
                mTmAparicion = (vRegSup - vRegInf - 1) * 5 ' distancia entre registros * 5 Numeros
                mTmAparicion = mTmAparicion + (5 - vOrdInf) + vOrdSup
           Else
                mTmAparicion = (vRegSup - vRegInf - 1) * 2 ' distancia entre registros * 2 Numeros
                mTmAparicion = mTmAparicion + (2 - vOrdInf) + vOrdSup
           End If
           
      Case GordoPrimitiva
           '
           ' Si la bola es un Numero o es Clave
           '
           If TipoBola = 1 Then
                mTmAparicion = (vRegSup - vRegInf - 1) * 5 ' distancia entre registros * 5 Numeros
                mTmAparicion = mTmAparicion + (5 - vOrdInf) + vOrdSup
           Else
                mTmAparicion = (vRegSup - vRegInf - 1) ' distancia entre registros
           End If
    End Select
    '
    ' Devolvemos el tiempo entre apariciones
    '
    getTiempoIntervalo = mTmAparicion
    
  On Error GoTo 0
getTiempoIntervalo__CleanExit:
    Exit Function
            
getTiempoIntervalo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Bola.getTiempoIntervalo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Bola.getTiempoIntervalo", ErrDescription
End Function

' *===========(EOF): Bola.cls

