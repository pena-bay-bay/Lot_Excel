VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParametrosMuestra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : ParametrosMuestra
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : jue, 16/04/2009 23:51
' *     Versión    : 1.0
' *     Propósito  : Preparar la información de partida para la obtención de la
' *                  Muestra
' *
' *============================================================================*
Option Explicit
'
'
Private mFechaAnalisis          As Date         ' Fecha de analisis de la muestra
Private mDiasAnalisis           As Integer      ' Rango de dias para la muestra
Private mPeriodo                As Periodo      ' Periodo de fechas para el Rango
Private mInfoSorteo             As InfoSorteo   ' Objeto para comprobar las fechas del juego
Private m_iResgistroInicial     As Integer      ' Numero de registro inicial de la muestra
Private m_iResgistroFinal       As Integer      ' Numero de registro final de la muestra
Private m_iResgistroAnalisis    As Integer      ' Numero de registro de la fecha de analisis
Private m_datos                 As New BdDatos  ' Base de datos
Private mJuego                  As Juego        ' Juego de la muestra
Private mNumeroSorteos          As Integer      ' Numero de sorteos a analizar
Private mError                  As Integer      ' Código de Error
Private mTipoMuestra            As Boolean      ' Tipo de muestra: True - Por Numero sorteos
                                                '                  False - Por dias desde fecha análisis



' *============================================================================*
' *     Procedure  : Class_Initialize
' *     Fichero    : ParametrosMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : jue, 06/08/2009
' *     Asunto     : Inicializa las variables internas de la clase
' *============================================================================*
'
Private Sub Class_Initialize()
    Set mPeriodo = New Periodo
    Set mInfoSorteo = New InfoSorteo
    mInfoSorteo.Constructor JUEGO_DEFECTO
    m_iResgistroAnalisis = 0
    m_iResgistroFinal = 0
    m_iResgistroInicial = 0
    mFechaAnalisis = 0
End Sub


' *============================================================================*
' *     Procedure  : FechaAnalisis
' *     Fichero    : ParametrosMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : jue, 06/08/2009
' *     Asunto     :
' *============================================================================*
'
Public Property Get FechaAnalisis() As Date
    FechaAnalisis = mFechaAnalisis
End Property

Public Property Let FechaAnalisis(ByVal dtFechaAnalisis As Date)
    '
    '   Comprueba que la fecha sea de sorteo
    '
    If mInfoSorteo.EsFechaSorteo(dtFechaAnalisis) Then
        mFechaAnalisis = dtFechaAnalisis
    Else
        mFechaAnalisis = mInfoSorteo.GetProximoSorteo(dtFechaAnalisis)
    End If
    '
    ' Si no hay registro para la fecha de análisis ¿cuantos sorteos deberían haber?
    '
    m_iResgistroAnalisis = m_datos.GetRegistroFecha(mFechaAnalisis)
    If m_iResgistroAnalisis = 0 And FechaFinal <> 0 Then
        m_iResgistroAnalisis = m_iResgistroFinal + mInfoSorteo.GetSorteosEntreFechas(FechaFinal, mFechaAnalisis)
    End If
End Property

' *============================================================================*
' *     Procedure  : FechaInicial
' *     Fichero    : ParametrosMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : jue, 06/08/2009
' *     Asunto     :
' *============================================================================*
'
Public Property Get FechaInicial() As Date
    FechaInicial = mPeriodo.FechaInicial
End Property

Public Property Let FechaInicial(ByVal dtFechaInicial As Date)
    mPeriodo.FechaInicial = dtFechaInicial
    m_iResgistroInicial = m_datos.GetRegistroFecha(dtFechaInicial)
End Property

' *============================================================================*
' *     Procedure  : FechaFinal
' *     Fichero    : ParametrosMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : jue, 06/08/2009
' *     Asunto     :
' *============================================================================*
'
Public Property Get FechaFinal() As Date
    FechaFinal = mPeriodo.FechaFinal
End Property

Public Property Let FechaFinal(ByVal dtFechaFinal As Date)
    mPeriodo.FechaFinal = dtFechaFinal
    m_iResgistroFinal = m_datos.GetRegistroFecha(dtFechaFinal)
    '
    ' Si no hay registro para la fecha de análisis ¿cuantos sorteos deberían haber?
    '
    m_iResgistroAnalisis = m_datos.GetRegistroFecha(mFechaAnalisis)
    If m_iResgistroAnalisis = 0 Then
        m_iResgistroAnalisis = m_iResgistroFinal + mInfoSorteo.GetSorteosEntreFechas(dtFechaFinal, mFechaAnalisis)
    End If
End Property

' *============================================================================*
' *     Procedure  : DiasAnalisis
' *     Fichero    : ParametrosMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : jue, 06/08/2009
' *     Asunto     :
' *============================================================================*
'
Public Property Get DiasAnalisis() As Integer
    DiasAnalisis = mDiasAnalisis
End Property
Public Property Let DiasAnalisis(NewValue As Integer)
    mDiasAnalisis = NewValue
    '
    '   Si tenemos fecha de analisis se calcula la fecha final como fecha anterior sorteo
    '
End Property


' *============================================================================*
' *     Procedure  : ResgistroInicial
' *     Fichero    : ParametrosMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 21/01/2012 22:11
' *     Asunto     :
' *============================================================================*
'
Public Property Get ResgistroInicial() As Integer
    ResgistroInicial = m_iResgistroInicial
End Property


' *============================================================================*
' *     Procedure  : ResgistroFinal
' *     Fichero    : ParametrosMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 21/01/2012 22:11
' *     Asunto     :
' *============================================================================*
'
Public Property Get ResgistroFinal() As Integer
    ResgistroFinal = m_iResgistroFinal
End Property


' *============================================================================*
' *     Procedure  : ResgistroAnalisis
' *     Fichero    : ParametrosMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 21/01/2012 22:11
' *     Asunto     :
' *============================================================================*
'
Public Property Get ResgistroAnalisis() As Integer
    ResgistroAnalisis = m_iResgistroAnalisis
End Property


' *============================================================================*
' *     Property   : NumeroSorteos
' *     Fichero    : ParametrosMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 01/04/2017 19:21
' *     Asunto     : Devuelve el número de sorteos entre el sorteo inicial
' *                  y el final
' *============================================================================*
'
Public Property Get NumeroSorteos() As Integer

    If mNumeroSorteos = 0 Then
        mNumeroSorteos = m_iResgistroFinal - m_iResgistroInicial + 1
    End If
    NumeroSorteos = mNumeroSorteos

End Property

Public Property Let NumeroSorteos(vNewValue As Integer)
    mNumeroSorteos = vNewValue
    '
    '   Se obtiene la fecha inicial del sorteo si tenemos una fecha inicial
    '
    If mPeriodo.FechaFinal <> 0 Then
        FechaInicial = m_datos.GetFecha(mPeriodo.FechaFinal, mNumeroSorteos)
    End If
End Property



'------------------------------------------------------------------------------*
' Procedimiento  : Juego
' Fecha          : vi., 08/may/2020 19:26:10
' Propósito      : Concurso de los parametros de la muestra
'------------------------------------------------------------------------------*
Public Property Get Juego() As Juego
    Juego = mJuego
End Property

Public Property Let Juego(ByVal vNewValue As Juego)
    mJuego = vNewValue
    mInfoSorteo.Constructor mJuego
End Property


'------------------------------------------------------------------------------*
' Procedimiento  : TipoMuestra
' Fecha          : vi., 08/may/2020 19:26:10
' Propósito      : Tipo de muestra; { True: Numeros de sorteo, False: Dias }
'------------------------------------------------------------------------------*
Public Property Get TipoMuestra() As Boolean
    TipoMuestra = mTipoMuestra
End Property

Public Property Let TipoMuestra(NewValue As Boolean)
    mTipoMuestra = NewValue
End Property


' *============================================================================*
' *     Function   : Validar
' *     Fichero    : ParametrosMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 01/04/2017 19:42
' *     Asunto     : Valida la consistencia de los datos
' *
' *============================================================================*
'
Public Function Validar() As Boolean
        
   On Error GoTo Validar_Error
  
        mError = 0
        If FechaAnalisis = 0 Then            ' Fecha de analisis distinta de 0
            mError = 1
        End If
        If FechaFinal = 0 Then               ' Fecha final distinta de 0
            mError = mError + 2
        End If
        If FechaInicial = 0 And _
        NumeroSorteos = 0 Then               ' O fecha inicial <> 0 o Sorteos <> 0
            mError = mError + 4
        End If
        If FechaAnalisis <= FechaFinal Then  ' Validar Fecha Analisis menor fecha final
            mError = mError + 8
        End If
        If FechaFinal <= FechaInicial Then   ' Validar Fecha final menor fecha inicial
            mError = mError + 16
        End If
                                             '
                                             ' Validar que la fecha de análisis
                                             ' pertenece al sorteo
        If Not mInfoSorteo.EsFechaSorteo(FechaAnalisis) Then
            mError = mError + 32
        End If
                                             ' Validar que la fecha final
                                             ' pertenece al sorteo
        If Not mInfoSorteo.EsFechaSorteo(FechaFinal) Then
            mError = mError + 64
        End If
                                             ' Validar que la fecha inicial
                                             ' pertenece al sorteo
        If Not mInfoSorteo.EsFechaSorteo(FechaInicial) Then
            mError = mError + 128
        End If
                
        If mError = 0 Then
            Validar = True
        Else
            Validar = False
        End If
    On Error GoTo 0
Validar_CleanExit:
       Exit Function

Validar_Error:
       Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
       ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
       Call HandleException(ErrNumber, ErrDescription, "ParametrosMuestra.Validar", ErrSource)
       Err.Raise ErrNumber, "ParametrosMuestra.Validar", ErrDescription
End Function

' *============================================================================*
' *     Function   : GetMensaje
' *     Fichero    : ParametrosMuestra
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 15/04/2017 18:09
' *     Asunto     : Devuelve el error de inconsistencia de la clase
' *
' *============================================================================*
'
Public Function GetMensaje() As String
    Dim m_Mensaje           As String
    
   On Error GoTo GetMensaje_Error
   
       If mError = 0 Then
            m_Mensaje = MSG_NOERROR
       Else
            m_Mensaje = MSG_HAYERRORES & vbCrLf
       End If
       If mError And 1 Then
            m_Mensaje = m_Mensaje & MSG_FECHAANALISCERO & vbCrLf
       End If
       If mError And 2 Then
            m_Mensaje = m_Mensaje & MSG_FECHAFINALCERO & vbCrLf
       End If
       If mError And 4 Then
            m_Mensaje = m_Mensaje & MSG_FECHAINICIALCERO & vbCrLf
       End If
       If mError And 4 Then
            m_Mensaje = m_Mensaje & MSG_NUMSORTEOSCERO & vbCrLf
       End If
       If mError And 8 Then
            m_Mensaje = m_Mensaje & MSG_FECHAANALISMENOR & vbCrLf
       End If
       If mError And 16 Then
            m_Mensaje = m_Mensaje & MSG_FECHAFINALMENOR & vbCrLf
       End If
       If mError And 32 Then
            m_Mensaje = m_Mensaje & MSG_FECHAANALISNOJUEGO & vbCrLf
       End If
       If mError And 64 Then
            m_Mensaje = m_Mensaje & MSG_FECHAFINALNOJUEGO & vbCrLf
       End If
       If mError And 128 Then
            m_Mensaje = m_Mensaje & MSG_FECHAINICIALNOJUEGO & vbCrLf
       End If
       '
       '
       GetMensaje = m_Mensaje
       
    On Error GoTo 0
GetMensaje_CleanExit:
       Exit Function

GetMensaje_Error:
       Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
       ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
       Call HandleException(ErrNumber, ErrDescription, "ParametrosMuestra.GetMensaje", ErrSource)
       Err.Raise ErrNumber, "ParametrosMuestra.GetMensaje", ErrDescription
End Function


'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          :
' Propósito      : Visualizar la clase en formato Texto
'------------------------------------------------------------------------------*
Public Function ToString() As String
    Dim mStr As String
  On Error GoTo ToString_Error
    
    ' Muestra: (FechaAnalisis) [FechaDesde..FechaHasta] (n Sorteos)
    mStr = "Parametros Muestra: ($1) $2 $3 Sorteos"
    mStr = Replace(mStr, "$1", Format(mFechaAnalisis, "dd/mm/yyyy"))
    mStr = Replace(mStr, "$2", mPeriodo.ToString())
    mStr = Replace(mStr, "$3", Format(mNumeroSorteos, "#0"))
    ToString = mStr
  
  On Error GoTo 0
ToString_CleanExit:
    Exit Function
ToString_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ParametrosMuestra.ToString", ErrSource)
    Err.Raise ErrNumber, "ParametrosMuestra.ToString", ErrDescription
End Function

'   TODO: hacer una rutina SetDatos que analice la información suministrada al objeto y deduzca el resto de parametros
'   con fechaAnalisis y dias se puede calcular periodo y tipo de muestra
'   con fechaAnalisis y numero de registros se puede calcular registros y tipo de muestra
'   Con FechaAnalisis y fecha inicial y final se pueden calcular dias y tipo de muestra
'   Si tipo de muestra esta establecido podemos calcular el
'' *===========(EOF): ParametrosMuestra.cls
