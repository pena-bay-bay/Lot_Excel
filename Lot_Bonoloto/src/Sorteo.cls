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
'---------------------------------------------------------------------------------------
' Modulo    : Sorteo
' Creado    : 22-ago-2007 08:23
' Autor     : Carlos Almela Baeza
' Version   : 1.0.0 22-ago-2007 08:23
' Objeto    : Clase que representa un sorteo de cada uno de los juegos
'---------------------------------------------------------------------------------------
Option Explicit
'
'
'  Variables Publicas
'
Public EntidadNegocio               As EntidadNegocio
Private m_eJuego                    As Juego
Public Premios                      As Collection
'
'  Variables privadas
'
Private m_eTipoJuego                As ModalidadJuego
Private m_sNumeroSorteo             As String
Private m_dtFecha                   As Date         ' Fecha del sorteo
Private m_sDia                      As String       ' Dia de la semana
Private m_iComplementario           As Integer      ' Complementario
Private m_iReintegro                As Integer      ' Reintegro
Private m_objCombinacion            As New Combinacion  ' Combinación ganadora
Private m_iRegistro                 As Integer      ' Numero de registro
Private m_iSemana                   As Integer
Private m_objEstrellaUno            As Numero
Private m_objEstrellaDos            As Numero
Private m_cImporteVenta             As Currency
Private m_cImporteBote              As Currency
Private m_bOrdenado                 As Boolean

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
'
' TODO: crear la colección de premios por defecto
   Set Premios = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_objCombinacion = Nothing
    Set m_objEstrellaUno = Nothing
    Set m_objEstrellaDos = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Constructor
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Constructor(Fila As Range)
    Dim mCol As Collection
    Dim i As Integer
    Dim m_num As Numero
    Dim mStrKey As String
   On Error GoTo Constructor_Error

    Set mCol = New Collection
    
    Select Case JUEGO_DEFECTO
        Case Bonoloto, LoteriaPrimitiva:
            With Fila
                Me.Registro = .Cells(1, 1).Value
                Me.NumeroSorteo = .Cells(1, 2).Value
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
                m_bOrdenado = IIf(.Cells(1, 14).Value = "Si", True, False)
                Set m_objCombinacion.Numeros = mCol
                m_eTipoJuego = LP_LB_6_49
                m_eJuego = JUEGO_DEFECTO
            End With
        
        Case GordoPrimitiva:
            With Fila
                Me.Registro = .Cells(1, 1).Value
                Me.NumeroSorteo = .Cells(1, 2).Value
                Me.Fecha = CDate(.Cells(1, 3).Value)
                m_bOrdenado = IIf(.Cells(1, 6).Value = "Si", True, False)
                For i = 0 To 4
                    If IsNumeric(.Cells(1, 7 + i).Value) _
                    And Not IsEmpty(.Cells(1, 7 + i).Value) Then
                        If (.Cells(1, 7 + i).Value) <= 54 Then
                            Set m_num = New Numero
                            m_num.Valor = .Cells(1, 7 + i).Value
                            mStrKey = CStr(m_num.Valor)
                            If (mCol.Count <> 0) Then
                                mCol.Add m_num, mStrKey, , mCol.Count
                            Else
                                mCol.Add m_num, mStrKey
                            End If
                        End If
                    End If
                Next i
                m_iReintegro = .Cells(1, 12).Value
                Set m_objCombinacion.Numeros = mCol
                m_eTipoJuego = GP_5_54
                m_eJuego = JUEGO_DEFECTO
            End With
       End Select


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

'---------------------------------------------------------------------------------------
' Procedure : Numero de Sorteo
' Author    : CHARLY
' Date      : 12/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get NumeroSorteo() As String
    NumeroSorteo = m_sNumeroSorteo
End Property
Public Property Let NumeroSorteo(NewValue As String)
    m_sNumeroSorteo = NewValue
End Property

'---------------------------------------------------------------------------------------
' Procedure : Registro
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Registro() As Integer
    Registro = Me.EntidadNegocio.Id
End Property

Public Property Let Registro(ByVal iRegistro As Integer)
    Me.EntidadNegocio.Id = iRegistro
End Property

'---------------------------------------------------------------------------------------
' Procedure : Fecha
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
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
    Me.TipoJuego = LP_LB_6_49
    If Me.Dia = "J" Or Me.Dia = "S" Then
        Juego = LoteriaPrimitiva
    Else
        Juego = Bonoloto
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : Total_Pronosticos
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Total_Pronosticos() As Integer
    Total_Pronosticos = m_objCombinacion.Count
End Property

'---------------------------------------------------------------------------------------
' Procedure : Total_Apuestas
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get TotalApuestas() As Integer
    Select Case m_objCombinacion.Count
        Case 5: TotalApuestas = 44
        Case 6: TotalApuestas = 1
        Case 7: TotalApuestas = 7
        Case 8: TotalApuestas = 28
        Case 9: TotalApuestas = 84
        Case 10: TotalApuestas = 210
        Case 11: TotalApuestas = 462
        Case Else: TotalApuestas = 0
    End Select
End Property

'---------------------------------------------------------------------------------------
' Procedure : Texto
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Texto() As Variant
    Texto = m_objCombinacion.Texto
End Property

Public Property Let Texto(datTexto As Variant)
    m_objCombinacion.Texto = CStr(datTexto)
End Property


'---------------------------------------------------------------------------------------
' Procedure : Dia
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Dia() As String
    Dia = m_sDia
End Property

Public Property Let Dia(ByVal sDia As String)
    m_sDia = sDia
End Property

' *============================================================================*
' *     Procedure  : Combinacion
' *     Fichero    : Sorteo
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : vie, 20/01/2012 23:33
' *     Asunto     :
' *============================================================================*
'
Public Property Get Combinacion() As Combinacion

  On Error GoTo Combinacion_Error

    Set Combinacion = m_objCombinacion

Combinacion_CleanExit:
   On Error GoTo 0
    Exit Property

Combinacion_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Sorteo.Combinacion")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

Public Property Set Combinacion(objCombinacion As Combinacion)

  On Error GoTo Combinacion_Error

    Set m_objCombinacion = objCombinacion

Combinacion_CleanExit:
   On Error GoTo 0
    Exit Property

Combinacion_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Sorteo.Combinacion")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Complementario
' Author    : Charly
' Date      : 17/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Complementario() As Integer
    Dim oNum As Numero
    Dim i As Integer
  On Error GoTo Complementario_Error

    If m_iComplementario = 0 Then
        i = Me.Combinacion.Count
        Set oNum = Me.Combinacion.Numeros.Item(i)
        m_iComplementario = oNum.Valor
    End If
    Complementario = m_iComplementario

   On Error GoTo 0
   Exit Property

Complementario_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.Complementario")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Property

'---------------------------------------------------------------------------------------
' Procedure : TipoJuego
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get TipoJuego() As ModalidadJuego
    TipoJuego = m_eTipoJuego
End Property

Public Property Let TipoJuego(ByVal eTipoJuego As ModalidadJuego)
    m_eTipoJuego = eTipoJuego
End Property

'---------------------------------------------------------------------------------------
' Procedure : Reintegro
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Reintegro() As Integer
    Reintegro = m_iReintegro
End Property

Public Property Let Reintegro(ByVal iReintegro As Integer)
    m_iReintegro = iReintegro
End Property

'---------------------------------------------------------------------------------------
' Procedure : Semana
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
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
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Ordenado() As Boolean

    Ordenado = m_bOrdenado

End Property

Public Property Let Ordenado(ByVal bOrdenado As Boolean)

    m_bOrdenado = bOrdenado

End Property

'---------------------------------------------------------------------------------------
' Procedure : ImporteVenta
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ImporteVenta() As Currency

   On Error GoTo ImporteVenta_Error

    ImporteVenta = m_cImporteVenta

   On Error GoTo 0
   Exit Property

ImporteVenta_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.ImporteVenta")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : ImporteVenta
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ImporteVenta(ByVal cImporteVenta As Currency)

   On Error GoTo ImporteVenta_Error

    m_cImporteVenta = cImporteVenta

   On Error GoTo 0
   Exit Property

ImporteVenta_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.ImporteVenta")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : EstrellaUno
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get EstrellaUno() As Numero

   On Error GoTo EstrellaUno_Error

    Set EstrellaUno = m_objEstrellaUno

   On Error GoTo 0
   Exit Property

EstrellaUno_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.EstrellaUno")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : EstrellaUno
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Set EstrellaUno(objEstrellaUno As Numero)

   On Error GoTo EstrellaUno_Error

    Set m_objEstrellaUno = objEstrellaUno

   On Error GoTo 0
   Exit Property

EstrellaUno_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.EstrellaUno")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : EstrellaDos
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get EstrellaDos() As Numero

   On Error GoTo EstrellaDos_Error

    Set EstrellaDos = m_objEstrellaDos

   On Error GoTo 0
   Exit Property

EstrellaDos_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.EstrellaDos")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : EstrellaDos
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Set EstrellaDos(objEstrellaDos As Numero)

   On Error GoTo EstrellaDos_Error

    Set m_objEstrellaDos = objEstrellaDos

   On Error GoTo 0
   Exit Property

EstrellaDos_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.EstrellaDos")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : ImporteBote
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ImporteBote() As Currency

   On Error GoTo ImporteBote_Error

    ImporteBote = m_cImporteBote

   On Error GoTo 0
   Exit Property

ImporteBote_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.ImporteBote")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : ImporteBote
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ImporteBote(ByVal cImporteBote As Currency)

   On Error GoTo ImporteBote_Error

    m_cImporteBote = cImporteBote

   On Error GoTo 0
   Exit Property

ImporteBote_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.ImporteBote")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Existe
' DateTime  : 12/ago/2007 18:56
' Author    : Carlos Almela Baeza
' Purpose   : Determina si un número está en la apuesta
'---------------------------------------------------------------------------------------
'
Public Function Existe(datNumero As Integer) As Boolean
    Existe = m_objCombinacion.Contiene(datNumero)
    If Not Existe Then
        If m_iComplementario = datNumero Then
            Existe = True
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : EsValido
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function EsValido() As Boolean

    Dim bResult As Boolean


   On Error GoTo EsValido_Error
' TODO: comprobar tipo de sorteo
    EsValido = bResult

   On Error GoTo 0
   Exit Function

EsValido_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.EsValido")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : EsWarming
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function EsWarming() As Boolean

    Dim bResult As Boolean


   On Error GoTo EsWarming_Error
' TODO: comprobar tipo de sorteo

    EsWarming = bResult

   On Error GoTo 0
   Exit Function

EsWarming_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.EsWarming")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetMensaje
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetMensaje() As String

    Dim sResult As String


   On Error GoTo GetMensaje_Error
' TODO: comprobar tipo de sorteo

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
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function ToString() As String

    Dim sResult As String


   On Error GoTo ToString_Error
' TODO: Serizalizar los datos en una cadena de texto

    ToString = sResult

   On Error GoTo 0
   Exit Function

ToString_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.ToString")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function


'---------------------------------------------------------------------------------------
' Procedure : Juego
' Author    : CHARLY
' Date      : 09/05/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Juego() As Juego

   On Error GoTo Juego_Error

    Juego = m_eJuego
   
   On Error GoTo 0
   Exit Property

Juego_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.Juego")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Juego
' Author    : CHARLY
' Date      : 09/05/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Juego(ByVal eJuego As Juego)

   On Error GoTo Juego_Error

    m_eJuego = eJuego
    
    Select Case m_eJuego
        Case LoteriaPrimitiva: m_eTipoJuego = LP_LB_6_49
        Case Bonoloto: m_eTipoJuego = LP_LB_6_49
        Case GordoPrimitiva: m_eTipoJuego = GP_5_54
        Case Euromillones:: m_eTipoJuego = EU_5_50
        Case PrimitivaBonoloto:: m_eTipoJuego = LP_LB_6_49
    End Select
'
' TODO: Cargar los premios por defecto
'
    SetPremiosDefecto

   On Error GoTo 0
   Exit Property

Juego_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.Juego")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : SetPremiosDefecto
' Author    : CHARLY
' Date      : 09/05/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SetPremiosDefecto()

   On Error GoTo SetPremiosDefecto_Error

    

   On Error GoTo 0
   Exit Sub

SetPremiosDefecto_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.SetPremiosDefecto")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
