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

'
'
'  Variables Publicas
'
Public EntidadNegocio               As EntidadNegocio
Private m_eJuego                    As Juego
Public Premios                      As Premios
'
'  Variables privadas
'
Private m_eTipoJuego                As ModalidadJuego
Private m_sNumeroSorteo             As String
Private m_dtFecha                   As Date         ' Fecha del sorteo
Private m_sDia                      As String       ' Dia de la semana
Private m_iComplementario           As Integer      ' Complementario
Private m_iReintegro                As Integer      ' Reintegro
Private m_objCombinacion            As Combinacion  ' Estrellas
Private m_objEstrellas              As Combinacion  ' Estrellas
Private m_iRegistro                 As Integer      ' Numero de registro
Private m_iSemana                   As Integer
Private m_objEstrellaUno            As Numero
Private m_objEstrellaDos            As Numero

Private m_cImporteVenta             As Currency
Private m_cImporteBote              As Currency
Private m_bOrdenado                 As Boolean
Private m_sElMillon                 As String
Private m_lJoker                    As Long
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
    Set m_objEstrellaUno = New Numero
    Set m_objEstrellaDos = New Numero
End Sub

Private Sub Class_Terminate()
    Set EntidadNegocio = Nothing
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
                For i = 0 To 5
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
                '
                '   Premios
                '
                If Len(.Cells(1, 15).Value) > 0 Then
                    Set Premios = New Premios
                    Premios.Parse .Cells(1, 15).Value
                End If
            End With
        
        Case gordoPrimitiva:
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
                '
                '   Premios
                '
                If Len(.Cells(1, 13).Value) > 0 Then
                    Set Premios = New Premios
                    Premios.Parse (.Cells(1, 13).Value)
                End If
            End With
        
        Case Euromillones:
            With Fila
                Me.Registro = .Cells(1, 1).Value
                Me.NumeroSorteo = .Cells(1, 2).Value
                Me.Fecha = CDate(.Cells(1, 3).Value)
                m_bOrdenado = IIf(.Cells(1, 6).Value = "Si", True, False)
                For i = 0 To 4
                    If IsNumeric(.Cells(1, 7 + i).Value) _
                    And Not IsEmpty(.Cells(1, 7 + i).Value) Then
                        If (.Cells(1, 7 + i).Value) <= 50 Then
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
                Set m_objCombinacion.Numeros = mCol
                
                If IsNumeric(.Cells(1, 11).Value) _
                And Not IsEmpty(.Cells(1, 11).Value) Then
                    If (.Cells(1, 11).Value) <= 12 Then
                        Set m_num = New Numero
                        m_num.Valor = .Cells(1, 11).Value
                        Set m_objEstrellaUno = m_num
                    End If
                End If
                If IsNumeric(.Cells(1, 12).Value) _
                And Not IsEmpty(.Cells(1, 12).Value) Then
                    If (.Cells(1, 11).Value) <= 12 Then
                        Set m_num = New Numero
                        m_num.Valor = .Cells(1, 12).Value
                        Set m_objEstrellaDos = m_num
                    End If
                End If
                m_eTipoJuego = EU_5_50
                m_eJuego = JUEGO_DEFECTO
                '
                '   Premios
                '
                If Len(.Cells(1, 13).Value) > 0 Then
                    Set Premios = New Premios
                    Premios.Parse (.Cells(1, 13).Value)
                End If
            End With
        
       End Select


   On Error GoTo 0
   Exit Sub

Constructor_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.Constructor")
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
'    strDiaSemana = Format(dtFecha, "dddd")
'
'    Me.Dia = UCase(Left(strDiaSemana, 1))
'    If Me.Dia = "M" And Len(strDiaSemana) > 6 Then
'        Me.Dia = "X"
'    End If
'    Me.TipoJuego = LP_LB_6_49
'    If Me.Dia = "J" Or Me.Dia = "S" Then
'        Juego = LoteriaPrimitiva
'    Else
'        Juego = Bonoloto
'    End If
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

'---------------------------------------------------------------------------------------
' Procedure : ElMillon
' Author    : CHARLY
' Date      :
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ElMillon() As String
    ElMillon = m_sElMillon
End Property
Public Property Let ElMillon(vNewValue As String)
    m_sElMillon = vNewValue
End Property



'---------------------------------------------------------------------------------------
' Procedure : ElMillon
' Author    : CHARLY
' Date      :
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Joker() As Long
    Joker = m_lJoker
End Property

Public Property Let Joker(vNewValue As Long)
    m_lJoker = vNewValue
End Property


'---------------------------------------------------------------------------------------
' Procedure : Total_Pronosticos
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Property Get Total_Pronosticos() As Integer
'    Total_Pronosticos = m_objCombinacion.Count
'End Property

'---------------------------------------------------------------------------------------
' Procedure : Total_Apuestas
' Author    : CHARLY
' Date      : 17/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Property Get TotalApuestas() As Integer
'    '
'    '   #TODO: SEGUN JUEGO
'    '
'
'    Select Case m_objCombinacion.Count
'        Case 5: TotalApuestas = 44
'        Case 6: TotalApuestas = 1
'        Case 7: TotalApuestas = 7
'        Case 8: TotalApuestas = 28
'        Case 9: TotalApuestas = 84
'        Case 10: TotalApuestas = 210
'        Case 11: TotalApuestas = 462
'        Case Else: TotalApuestas = 0
'    End Select
'End Property
'
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
    Call HandleException(ErrNumber, ErrDescription, "Sorteo.Combinacion")
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

''---------------------------------------------------------------------------------------
'' Procedure : Complementario
'' Author    : Charly
'' Date      : 17/11/2013
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Get Complementario() As Integer
'    Dim oNum As Numero
'    Dim i As Integer
'  On Error GoTo Complementario_Error
'
'    If m_iComplementario = 0 Then
'        i = Me.Combinacion.Count
'        Set oNum = Me.Combinacion.Numeros.Item(i)
'        m_iComplementario = oNum.Valor
'    End If
'    Complementario = m_iComplementario
'
'   On Error GoTo 0
'   Exit Property
'
'Complementario_Error:
'   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
'   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
'   Call HandleException(ErrNumber, ErrDescription, "Sorteo.Complementario")
'   Err.Raise ErrNumber, ErrSource, ErrDescription
'End Property

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
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.ImporteVenta")
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
'Public Property Get EstrellaUno() As Numero
'
'   On Error GoTo EstrellaUno_Error
'
'    Set EstrellaUno = m_objEstrellaUno
'
'   On Error GoTo 0
'   Exit Property
'
'EstrellaUno_Error:
'   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
'   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
'   Call HandleException(ErrNumber, ErrDescription, "Sorteo.EstrellaUno")
'   Err.Raise ErrNumber, ErrSource, ErrDescription
'End Property

'---------------------------------------------------------------------------------------
' Procedure : EstrellaUno
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Property Set EstrellaUno(objEstrellaUno As Numero)
'
'   On Error GoTo EstrellaUno_Error
'
'    Set m_objEstrellaUno = objEstrellaUno
'
'   On Error GoTo 0
'   Exit Property
'
'EstrellaUno_Error:
'   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
'   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
'   Call HandleException(ErrNumber, ErrDescription, "Sorteo.EstrellaUno")
'   Err.Raise ErrNumber, ErrSource, ErrDescription
'End Property

'---------------------------------------------------------------------------------------
' Procedure : EstrellaDos
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Property Get EstrellaDos() As Numero
'
'   On Error GoTo EstrellaDos_Error
'
'    Set EstrellaDos = m_objEstrellaDos
'
'   On Error GoTo 0
'   Exit Property
'
'EstrellaDos_Error:
'   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
'   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
'   Call HandleException(ErrNumber, ErrDescription, "Sorteo.EstrellaDos")
'   Err.Raise ErrNumber, ErrSource, ErrDescription
'End Property

'---------------------------------------------------------------------------------------
' Procedure : EstrellaDos
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Property Set EstrellaDos(objEstrellaDos As Numero)
'
'   On Error GoTo EstrellaDos_Error
'
'    Set m_objEstrellaDos = objEstrellaDos
'
'   On Error GoTo 0
'   Exit Property
'
'EstrellaDos_Error:
'   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
'   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
'   Call HandleException(ErrNumber, ErrDescription, "Sorteo.EstrellaDos")
'   Err.Raise ErrNumber, ErrSource, ErrDescription
'End Property
'
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
' #TODO: comprobar tipo de sorteo
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
' #TODO: comprobar tipo de sorteo

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
' #TODO: comprobar tipo de sorteo

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
    '
    '   #TODO: Incluir el nombre del juego o letras
    '
    Select Case m_eJuego
    Case Bonoloto, LoteriaPrimitiva:
        sResult = Replace(LIT_JUEGO7_49, "$1", Me.Id)
        sResult = Replace(sResult, "$2", Me.NumeroSorteo)
        sResult = Replace(sResult, "$3", Me.Fecha)
        sResult = Replace(sResult, "$4", Me.Semana)
        sResult = Replace(sResult, "$5", Me.Dia)
        sResult = Replace(sResult, "$6", Me.Combinacion.ToString(True))
        sResult = Replace(sResult, "$7", Me.Complementario)
        sResult = Replace(sResult, "$8", Me.Reintegro)
        sResult = Replace(sResult, "$9", IIf(Me.Ordenado, LIT_ENORDEN, LIT_NOTENORDEN))
       
    Case Euromillones
        sResult = Replace(LIT_JUEGO5_50, "$1", Me.Id)
        sResult = Replace(sResult, "$2", Me.NumeroSorteo)
        sResult = Replace(sResult, "$3", Me.Fecha)
        sResult = Replace(sResult, "$4", Me.Semana)
        sResult = Replace(sResult, "$5", Me.Dia)
        sResult = Replace(sResult, "$6", Me.Combinacion.ToString(True))
        sResult = Replace(sResult, "$7", Me.Estrellas.ToString(True))
        sResult = Replace(sResult, "$8", IIf(Me.Ordenado, LIT_ENORDEN, LIT_NOTENORDEN))
        
    Case gordoPrimitiva
        sResult = Replace(LIT_JUEGO5_54, "$1", Me.Id)
        sResult = Replace(sResult, "$2", Me.NumeroSorteo)
        sResult = Replace(sResult, "$3", Me.Fecha)
        sResult = Replace(sResult, "$4", Me.Semana)
        sResult = Replace(sResult, "$5", Me.Dia)
        sResult = Replace(sResult, "$6", Me.Combinacion.ToString(True))
        sResult = Replace(sResult, "$7", Me.Reintegro)
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
        Case gordoPrimitiva: m_eTipoJuego = GP_5_54
        Case Euromillones:: m_eTipoJuego = EU_5_50
        Case PrimitivaBonoloto:: m_eTipoJuego = LP_LB_6_49
    End Select
    '
    '   Establece premios por defecto
    '
    Premios.Juego = eJuego

   On Error GoTo 0
   Exit Property

Juego_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Sorteo.Juego")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

