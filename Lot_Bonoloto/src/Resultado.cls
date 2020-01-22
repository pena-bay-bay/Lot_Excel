VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Resultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Modulo    : Resultado
' Creado    : 01-nov-2006 20:29
' Autor     : Carlos Almela Baeza
' Version   : 1.0.0 Fecha
' Objeto    : Objeto Sorteo
'---------------------------------------------------------------------------------------
Private m_dtFecha                   As Date         ' Fecha del sorteo
Private m_sDia                      As String       ' Dia de la semana
Private m_iSem                      As Integer      ' Numero de sorteo o semana dentro del año
Private m_vNumeros(5)               As Integer      ' Combinación ganadora
Private m_iComplementario           As Integer      ' Complementario
Private m_iReintegro                As Integer      ' Reintegro
Private m_iNumRegistro              As Integer      ' Numero de registro

'---------------------------------------------------------------------------------------
' Procedure : Constructor
' DateTime  : 12/ago/2007 18:30
' Author    : Carlos Almela Baeza
' Purpose   : Procedimiento de carga de Sorteos celebrados
' Parametros: Fila de tipo rango de Excel que contiene la siguiente estructura:
'           Fecha del sorteo
'           Dia de la semana
'           Semana del año
'           Combinacion ganadora ( 6 Numeros )
'           Complementario
'           Reintegro
'---------------------------------------------------------------------------------------
'
Public Sub Constructor(Fila As Range)
    'N   Sorteo  Fecha   Dia Sem N1  N2  N3  N4  N5  N6  C   R
    '1   2       3       4   5   6   7   8   9   10  11  12  13
   On Error GoTo Constructor_Error

    With Fila
        m_iNumRegistro = .Cells(1, 1).Value
        m_dtFecha = CDate(.Cells(1, 3).Value)
        m_sDia = .Cells(1, 4).Value
        m_iSem = .Cells(1, 5).Value
        m_vNumeros(0) = .Cells(1, 6).Value
        m_vNumeros(1) = .Cells(1, 7).Value
        m_vNumeros(2) = .Cells(1, 8).Value
        m_vNumeros(3) = .Cells(1, 9).Value
        m_vNumeros(4) = .Cells(1, 10).Value
        m_vNumeros(5) = .Cells(1, 11).Value
        m_iComplementario = .Cells(1, 12).Value
        m_iReintegro = .Cells(1, 13).Value
    End With
   On Error GoTo 0
   Exit Sub

Constructor_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure Constructor of Módulo de clase Resultado"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Fecha
' DateTime  : 12/ago/2007 18:46
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Fecha() As Date
    Fecha = m_dtFecha
End Property

Public Property Let Fecha(ByVal dtFecha As Date)
    m_dtFecha = dtFecha
End Property

'---------------------------------------------------------------------------------------
' Procedure : dia
' DateTime  : 12/ago/2007 18:47
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Dia() As String
    Dia = m_sDia
End Property

Public Property Let Dia(ByVal sDia As String)
    m_sDia = sDia
End Property

'---------------------------------------------------------------------------------------
' Procedure : Sem
' DateTime  : 12/ago/2007 18:47
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Sem() As Integer
    Sem = m_iSem
End Property

Public Property Let Sem(ByVal iSem As Integer)
    m_iSem = iSem
End Property

'---------------------------------------------------------------------------------------
' Procedure : Numeros
' DateTime  : 12/ago/2007 18:47
' Author    : Carlos Almela Baeza
' Purpose   : Array con los Numeros de la combinación ganadora
'---------------------------------------------------------------------------------------
'
Public Property Get Numeros() As Variant

   On Error GoTo Numeros_Error

    Numeros = m_vNumeros

   On Error GoTo 0
   Exit Property

Numeros_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & _
        ") in procedure Numeros of Módulo de clase Resultado"

End Property

Public Property Let Numeros(ByVal vNumeros As Variant)

   On Error GoTo Numeros_Error
    If IsArray(vNumeros) Then
        m_vNumeros(0) = vNumeros(0)
        m_vNumeros(1) = vNumeros(1)
        m_vNumeros(2) = vNumeros(2)
        m_vNumeros(3) = vNumeros(3)
        m_vNumeros(4) = vNumeros(6)
        m_vNumeros(5) = vNumeros(5)
    End If
   On Error GoTo 0
   Exit Property

Numeros_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Numeros of Módulo de clase Resultado"

End Property

'---------------------------------------------------------------------------------------
' Procedure : Complementario
' DateTime  : 12/ago/2007 18:48
' Author    : Carlos Almela Baeza
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Complementario() As Integer
    Complementario = m_iComplementario
End Property

Public Property Let Complementario(ByVal iComplementario As Integer)
    m_iComplementario = iComplementario
End Property

'---------------------------------------------------------------------------------------
' Procedure : Reintegro
' DateTime  : 12/ago/2007 18:48
' Author    : Carlos Almela Baeza
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
' Procedure : Existe
' DateTime  : 12/ago/2007 18:49
' Author    : Carlos Almela Baeza
' Purpose   : Determina si el Numero pasado por parámetros está en la combinación
'             ganadora o en el complementario
'---------------------------------------------------------------------------------------
'
Public Function Existe(Numero As Integer, Optional Mas_Complementario As Boolean = True) As Boolean
    Existe = False
    For i = 0 To 5
        If m_vNumeros(i) = Numero Then
            Existe = True
            Exit For
        End If
    Next i
    If m_iComplementario = Numero And Mas_Complementario Then
        Existe = True
    End If
End Function

Public Property Get TipoJuego() As Juego
    Select Case Me.Dia
        Case "D": TipoJuego = gordoPrimitiva
        Case "J", "S": TipoJuego = LoteriaPrimitiva
        Case Else: TipoJuego = Bonoloto
    End Select
End Property

' *============================================================================*
' *     Procedure  : NumRegistro
' *     Fichero    : Resultado
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 14/01/2012 20:44
' *     Asunto     :
' *============================================================================*
'
Public Property Get NumRegistro() As Integer

  On Error GoTo NumRegistro_Error

    NumRegistro = m_iNumRegistro

NumRegistro_CleanExit:
   On Error GoTo 0
    Exit Property

NumRegistro_Error:
    

    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Resultado.NumRegistro")
    '   Lanza el Error
    'Err.Raise ErrNumber, ErrSource, ErrDescription
    '   Informa del error
    'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
    'Call Trace("CERRAR")


End Property
