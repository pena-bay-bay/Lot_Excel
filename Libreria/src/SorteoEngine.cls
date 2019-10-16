VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SorteoEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : SorteoEngine.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sá., 06/abr/2019 00:41:00
' *     Versión    : 1.0
' *     Propósito  : Proporcionar y almacenar Sorteos de un juego
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mDBase      As DataBaseExcel
Private mCol        As Sorteos
Private mObj        As Sorteo
Private oNum        As Numero
Private oPrms       As Premios
Private mRange      As Range
Private mFila       As Range
Private m_eJuego    As Juego
Private mId         As Integer
Private mTabla      As NombreTabla


'--- Constantes ---------------------------------------------------------------*
Private Const ERR_SORTEONOTFOUND = 2001
Private Const ERR_SORTEOIDNOTFOUND = 2002
Private Const ERR_FECHASORTEONOTFOUND = 2003
Private Const ERR_SORTEODUPLICADO = 2004
'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_SORTEONOTFOUND As String = "Sorteo no encontrando"
Private Const MSG_SORTEOIDNOTFOUND As String = "Sorteo de Id: #, no encontrado"
Private Const MSG_FECHASORTEONOTFOUND  As String = "Sorteo de fecha:#, no encontrado"
Private Const MSG_SORTEODUPLICADO  As String = "Sorteo de fecha: # ya existe"

'--- Errores ------------------------------------------------------------------*
'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    '
    '   Creamos y abrimos la base de datos
    '
    Set mDBase = New DataBaseExcel
    mDBase.Abrir
End Sub

Private Sub Class_Terminate()
    '
    '   Eliminamos el objeto
    '
    Set mDBase = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetTablaSorteo
' Author    : CHARLY
' Date      : vi., 03/may/2019 16:44:24
' Purpose   : Obtener la tabla del sorteo según el juego
'---------------------------------------------------------------------------------------
'
Private Function GetTablaSorteo(datJuego As Juego) As NombreTabla
    '
    '   Evaluamos el juego
    '
    Select Case datJuego
        Case bonoloto: GetTablaSorteo = tblBonoloto
        Case LoteriaPrimitiva: GetTablaSorteo = tblPrimitiva
        Case Euromillones: GetTablaSorteo = tblEuromillon
        Case gordoPrimitiva: GetTablaSorteo = tblGordo
    End Select
End Function

'--- Propiedades --------------------------------------------------------------*
'---------------------------------------------------------------------------------------
' Procedure : Juego
' Author    : CHARLY
' Date      : sá., 06/abr/2019 00:46:03
' Purpose   : Juego al que pertenece el sorteo
'---------------------------------------------------------------------------------------
'
Public Property Get Juego() As Juego
    Juego = m_eJuego
End Property

Public Property Let Juego(ByVal eJuego As Juego)
    m_eJuego = eJuego
End Property
    
'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GetListaJuegos
' Fecha          : lu., 04/feb/2019 23:51:32
' Propósito      : Devuelve una matríz con los nombres de los juegos
'------------------------------------------------------------------------------*
'
Public Function GetListaJuegos() As Variant
    GetListaJuegos = Split(NOMBRE_JUEGOS, ";")
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetNewSorteo
' Fecha          : lu., 04/feb/2019 23:51:32
' Propósito      : Obtener un Sorteo nuevo para completar sus propiedades
' Parámetros     : DatJuego Juego del sorteo
'------------------------------------------------------------------------------*
'
Public Function GetNewSorteo(datJuego As Juego) As Sorteo
  On Error GoTo GetNewSorteo_Error
    '
    '   Si no está definido el juego del motor lo establecemos
    '
    If m_eJuego = SinJuego Then
        m_eJuego = datJuego
    End If
    '
    '   Definimos el Objeto
    '
    Set mObj = New Sorteo
    '
    '   Cumplimentamos las propiedades
    '
    With mObj
        .Juego = datJuego
        .EntidadNegocio.IsNew = True
        .Id = GetIdUltimoSorteo(datJuego) + 1
        '
        '   Definimos los premios por defecto
        '
        Set oPrms = New Premios
        oPrms.Juego = datJuego
        oPrms.SetPremiosDefecto datJuego
        oPrms.IdSorteo = .Id
        Set .Premios = oPrms
    End With
    '
    '   Devolvemos el objeto
    '
    Set GetNewSorteo = mObj
    '
    '   Destruimos los objetos internos
    '
    Set mObj = Nothing
    Set oPrms = Nothing
    
  On Error GoTo 0
GetNewSorteo__CleanExit:
    Exit Function
            
GetNewSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoEngine.GetNewSorteo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoEngine.GetNewSorteo", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetSorteoById
' Fecha          : lu., 04/feb/2019 23:51:32
' Propósito      :
' Parámetros     : datId código de identificación del sorteo
'                  datJuego Juego del sorteo
'------------------------------------------------------------------------------*
'
Public Function GetSorteoById(datId As Integer, datJuego As Juego) As Sorteo
    Dim mStr        As String
  On Error GoTo GetSorteoById_Error
    '
    '   Obtenemos la tabla
    '
    mTabla = GetTablaSorteo(datJuego)
    '
    '   Obtenemos el rango del ID
    '
    Set mRange = mDBase.GetRowById(datId, mTabla)
    '
    '   Si no hay error lo ha encontrado
    '
    If mDBase.ErrNumber = 0 Then
        '
        '   Creamos el objeto Sorteo
        '
        Set mObj = New Sorteo
        '
        '   Asignamos atributos al parametro
        '
        mObj.Juego = datJuego
        mObj.Constructor mRange
        '
        '   Definimos los premios por defecto
        '
        Set oPrms = New Premios
        oPrms.Juego = datJuego
        oPrms.SetPremiosDefecto datJuego
        oPrms.IdSorteo = mObj.Id
        Set mObj.Premios = oPrms
        
    Else
        mStr = Replace(MSG_SORTEOIDNOTFOUND, "#", CStr(datId))
        Err.Raise ERR_SORTEOIDNOTFOUND, "SorteoEngine.GetSorteoById", mStr
    End If
    '
    '   Devolvemos el objeto
    '
    Set GetSorteoById = mObj
    '
    '   Destruimos el sorteo
    '
    Set mObj = Nothing
  
  On Error GoTo 0
GetSorteoById__CleanExit:
    Exit Function
            
GetSorteoById_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoEngine.GetSorteoById", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoEngine.GetSorteoById", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetSorteoByFecha
' Fecha          : lu., 04/feb/2019 23:51:32
' Propósito      :
' Parámetros     : DatFecha fecha de celebración del sorteo
'                  DatJuego Juego del sorteo
'------------------------------------------------------------------------------*
'
Public Function GetSorteoByFecha(datFecha As Date, datJuego As Juego) As Sorteo
    Dim mStr        As String
  On Error GoTo GetSorteoByFecha_Error
    '
    '   Obtenemos la tabla
    '
    mTabla = GetTablaSorteo(datJuego)
    '
    '   Obtenemos el rango de la fila fecha que es la tercera (2)
    '
    Set mRange = mDBase.GetRowByColumn(datFecha, 2, mTabla)
    '
    '   Si no hay error lo ha encontrado
    '
    If mDBase.ErrNumber = 0 Then
        '
        '   Creamos el objeto Sorteo
        '
        Set mObj = New Sorteo
        '
        '   Asignamos atributos al parametro
        '
        mObj.Juego = datJuego
        mObj.Constructor mRange
        '
        '   TODO: Buscar Premios
        '
        Set oPrms = New Premios
        oPrms.Juego = datJuego
        Set mObj.Premios = oPrms
    Else
        mStr = Replace(MSG_FECHASORTEONOTFOUND, "#", Format(datFecha, "dd/MM/yyyy"))
        Err.Raise ERR_FECHASORTEONOTFOUND, "GetRowByColumn", mStr
    End If
    '
    '   Devolvemos el objeto
    '
    Set GetSorteoByFecha = mObj
    '
    '   Destruimos el sorteo
    '
    Set mObj = Nothing
    Set oPrms = Nothing
    
  On Error GoTo 0
GetSorteoByFecha__CleanExit:
    Exit Function
            
GetSorteoByFecha_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoEngine.GetSorteoByFecha", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoEngine.GetSorteoByFecha", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : SetSorteo
' Fecha          : ma., 11/jun/2019 20:23:50
' Propósito      : Guardar un sorteo en la base de datos
' Parámetros     : Sorteo que se quiere guardar
'                  Si el sorteo no existe se agrega al final
'------------------------------------------------------------------------------*
'
Public Sub SetSorteo(datSorteo As Sorteo)

  On Error GoTo SetSorteo_Error
    '
    '   Asignamos el juego al Motor
    '
    m_eJuego = datSorteo.Juego
    '
    '   Segun el juego seleccionamos la tabla
    '
    mTabla = GetTablaSorteo(m_eJuego)
    '
    '   Evaluamos el estado del Sorteo
    '
    If datSorteo.EntidadNegocio.IsNew Then
        '
        '   Obtenemos la última fila de la tabla
        '
        Set mRange = mDBase.GetLastRow(mTabla)
        '
        '   Evaluamos el juego
        '
        Select Case datSorteo.Juego
        Case bonoloto
            '
            '   Obtenido el rango asignamos datos a cada celda
            '
            With mRange
                .Cells(1, 1).Value = datSorteo.Id                        'ID
                .Cells(1, 2).Value = datSorteo.NumSorteo                 'Sorteo
                .Cells(1, 3).Value = datSorteo.Fecha                     'Fecha
                .Cells(1, 4).Value = datSorteo.Dia                       'Dia
                .Cells(1, 5).Value = datSorteo.Semana                    'Sem
                .Cells(1, 6).Value = IIf(datSorteo.Ordenado, "Si", "No") 'Orden
                .Cells(1, 7).Value = datSorteo.CombinacionGanadora.Numeros.Item(1).Valor 'N1
                .Cells(1, 8).Value = datSorteo.CombinacionGanadora.Numeros.Item(2).Valor   'N2
                .Cells(1, 9).Value = datSorteo.CombinacionGanadora.Numeros.Item(3).Valor   'N3
                .Cells(1, 10).Value = datSorteo.CombinacionGanadora.Numeros.Item(4).Valor  'N4
                .Cells(1, 11).Value = datSorteo.CombinacionGanadora.Numeros.Item(5).Valor  'N5
                .Cells(1, 12).Value = datSorteo.CombinacionGanadora.Numeros.Item(6).Valor  'N6
                .Cells(1, 13).Value = datSorteo.Complementario           'C
                .Cells(1, 14).Value = datSorteo.Reintegro                'R
            End With
        
        Case LoteriaPrimitiva
            '
            '   Obtenido el rango asignamos datos a cada celda
            '
            With mRange
                .Cells(1, 1).Value = datSorteo.Id                        'ID
                .Cells(1, 2).Value = datSorteo.NumSorteo                 'Sorteo
                .Cells(1, 3).Value = datSorteo.Fecha                     'Fecha
                .Cells(1, 4).Value = datSorteo.Dia                       'Dia
                .Cells(1, 5).Value = datSorteo.Semana                    'Sem
                .Cells(1, 6).Value = IIf(datSorteo.Ordenado, "Si", "No") 'Orden
                .Cells(1, 7).Value = datSorteo.CombinacionGanadora.Numeros.Item(1).Valor   'N1
                .Cells(1, 8).Value = datSorteo.CombinacionGanadora.Numeros.Item(2).Valor   'N2
                .Cells(1, 9).Value = datSorteo.CombinacionGanadora.Numeros.Item(3).Valor   'N3
                .Cells(1, 10).Value = datSorteo.CombinacionGanadora.Numeros.Item(4).Valor  'N4
                .Cells(1, 11).Value = datSorteo.CombinacionGanadora.Numeros.Item(5).Valor  'N5
                .Cells(1, 12).Value = datSorteo.CombinacionGanadora.Numeros.Item(6).Valor  'N6
                .Cells(1, 13).Value = datSorteo.Complementario           'C
                .Cells(1, 14).Value = datSorteo.Reintegro                'R
            End With
        
        Case Euromillones
            '
            '   Obtenido el rango asignamos datos a cada celda
            '
            With mRange
                .Cells(1, 1).Value = datSorteo.Id                        'ID
                .Cells(1, 2).Value = datSorteo.NumSorteo                 'Sorteo
                .Cells(1, 3).Value = datSorteo.Fecha                     'Fecha
                .Cells(1, 4).Value = datSorteo.Dia                       'Dia
                .Cells(1, 5).Value = datSorteo.Semana                    'Sem
                .Cells(1, 6).Value = IIf(datSorteo.Ordenado, "Si", "No") 'Orden
                .Cells(1, 7).Value = datSorteo.CombinacionGanadora.Numeros.Item(1).Valor   'N1
                .Cells(1, 8).Value = datSorteo.CombinacionGanadora.Numeros.Item(2).Valor   'N2
                .Cells(1, 9).Value = datSorteo.CombinacionGanadora.Numeros.Item(3).Valor   'N3
                .Cells(1, 10).Value = datSorteo.CombinacionGanadora.Numeros.Item(4).Valor  'N4
                .Cells(1, 11).Value = datSorteo.CombinacionGanadora.Numeros.Item(5).Valor  'N5
                .Cells(1, 12).Value = datSorteo.EstrellaUno.Valor        'E1
                .Cells(1, 13).Value = datSorteo.EstrellaDos.Valor        'E2
            End With
        
        Case gordoPrimitiva
            '
            '   Obtenido el rango asignamos datos a cada celda
            '
            With mRange
                .Cells(1, 1).Value = datSorteo.Id                        'ID
                .Cells(1, 2).Value = datSorteo.NumSorteo                 'Sorteo
                .Cells(1, 3).Value = datSorteo.Fecha                     'Fecha
                .Cells(1, 4).Value = datSorteo.Dia                       'Dia
                .Cells(1, 5).Value = datSorteo.Semana                    'Sem
                .Cells(1, 6).Value = IIf(datSorteo.Ordenado, "Si", "No") 'Orden
                .Cells(1, 7).Value = datSorteo.CombinacionGanadora.Numeros.Item(1).Valor   'N1
                .Cells(1, 8).Value = datSorteo.CombinacionGanadora.Numeros.Item(2).Valor   'N2
                .Cells(1, 9).Value = datSorteo.CombinacionGanadora.Numeros.Item(3).Valor   'N3
                .Cells(1, 10).Value = datSorteo.CombinacionGanadora.Numeros.Item(4).Valor  'N4
                .Cells(1, 11).Value = datSorteo.CombinacionGanadora.Numeros.Item(5).Valor  'N5
                .Cells(1, 12).Value = datSorteo.Clave                    'Clave
            End With
        
        End Select
    
    ElseIf datSorteo.EntidadNegocio.MarkForDelete Then
        '
        '   Si está marcado para borrarse se elimina la fila
        '
        mDBase.DeleteRow datSorteo.Id, mTabla
        
    Else
        '
        '   Obtenemos la fila con el ID
        '
        Set mRange = mDBase.GetRowById(datSorteo.Id, mTabla)
        '
        '   Evaluamos el juego
        '
        Select Case datSorteo.Juego
        Case bonoloto
            '
            '   Obtenido el rango asignamos datos a cada celda
            '
            With mRange
                .Cells(1, 1).Value = datSorteo.Id                        'ID
                .Cells(1, 2).Value = datSorteo.NumSorteo                 'Sorteo
                .Cells(1, 3).Value = datSorteo.Fecha                     'Fecha
                .Cells(1, 4).Value = datSorteo.Dia                       'Dia
                .Cells(1, 5).Value = datSorteo.Semana                    'Sem
                .Cells(1, 6).Value = IIf(datSorteo.Ordenado, "Si", "No") 'Orden
                .Cells(1, 7).Value = datSorteo.CombinacionGanadora.Numeros.Item(1).Valor   'N1
                .Cells(1, 8).Value = datSorteo.CombinacionGanadora.Numeros.Item(2).Valor   'N2
                .Cells(1, 9).Value = datSorteo.CombinacionGanadora.Numeros.Item(3).Valor   'N3
                .Cells(1, 10).Value = datSorteo.CombinacionGanadora.Numeros.Item(4).Valor  'N4
                .Cells(1, 11).Value = datSorteo.CombinacionGanadora.Numeros.Item(5).Valor  'N5
                .Cells(1, 12).Value = datSorteo.CombinacionGanadora.Numeros.Item(6).Valor  'N6
                .Cells(1, 13).Value = datSorteo.Complementario           'C
                .Cells(1, 14).Value = datSorteo.Reintegro                'R
            End With
        
        Case LoteriaPrimitiva
            '
            '   Obtenido el rango asignamos datos a cada celda
            '
            With mRange
                .Cells(1, 1).Value = datSorteo.Id                        'ID
                .Cells(1, 2).Value = datSorteo.NumSorteo                 'Sorteo
                .Cells(1, 3).Value = datSorteo.Fecha                     'Fecha
                .Cells(1, 4).Value = datSorteo.Dia                       'Dia
                .Cells(1, 5).Value = datSorteo.Semana                    'Sem
                .Cells(1, 6).Value = IIf(datSorteo.Ordenado, "Si", "No") 'Orden
                .Cells(1, 7).Value = datSorteo.CombinacionGanadora.Numeros.Item(1).Valor   'N1
                .Cells(1, 8).Value = datSorteo.CombinacionGanadora.Numeros.Item(2).Valor   'N2
                .Cells(1, 9).Value = datSorteo.CombinacionGanadora.Numeros.Item(3).Valor   'N3
                .Cells(1, 10).Value = datSorteo.CombinacionGanadora.Numeros.Item(4).Valor  'N4
                .Cells(1, 11).Value = datSorteo.CombinacionGanadora.Numeros.Item(5).Valor  'N5
                .Cells(1, 12).Value = datSorteo.CombinacionGanadora.Numeros.Item(6).Valor  'N6
                .Cells(1, 13).Value = datSorteo.Complementario           'C
                .Cells(1, 14).Value = datSorteo.Reintegro                'R
            End With
        
        Case Euromillones
            '
            '   Obtenido el rango asignamos datos a cada celda
            '
            With mRange
                .Cells(1, 1).Value = datSorteo.Id                        'ID
                .Cells(1, 2).Value = datSorteo.NumSorteo                 'Sorteo
                .Cells(1, 3).Value = datSorteo.Fecha                     'Fecha
                .Cells(1, 4).Value = datSorteo.Dia                       'Dia
                .Cells(1, 5).Value = datSorteo.Semana                    'Sem
                .Cells(1, 6).Value = IIf(datSorteo.Ordenado, "Si", "No") 'Orden
                .Cells(1, 7).Value = datSorteo.CombinacionGanadora.Numeros.Item(1).Valor   'N1
                .Cells(1, 8).Value = datSorteo.CombinacionGanadora.Numeros.Item(2).Valor   'N2
                .Cells(1, 9).Value = datSorteo.CombinacionGanadora.Numeros.Item(3).Valor   'N3
                .Cells(1, 10).Value = datSorteo.CombinacionGanadora.Numeros.Item(4).Valor  'N4
                .Cells(1, 11).Value = datSorteo.CombinacionGanadora.Numeros.Item(5).Valor  'N5
                .Cells(1, 12).Value = datSorteo.EstrellaUno.Valor        'E1
                .Cells(1, 13).Value = datSorteo.EstrellaDos.Valor        'E2
            End With
        
        Case gordoPrimitiva
            '
            '   Obtenido el rango asignamos datos a cada celda
            '
            With mRange
                .Cells(1, 1).Value = datSorteo.Id                        'ID
                .Cells(1, 2).Value = datSorteo.NumSorteo                 'Sorteo
                .Cells(1, 3).Value = datSorteo.Fecha                     'Fecha
                .Cells(1, 4).Value = datSorteo.Dia                       'Dia
                .Cells(1, 5).Value = datSorteo.Semana                    'Sem
                .Cells(1, 6).Value = IIf(datSorteo.Ordenado, "Si", "No") 'Orden
                .Cells(1, 7).Value = datSorteo.CombinacionGanadora.Numeros.Item(1).Valor   'N1
                .Cells(1, 8).Value = datSorteo.CombinacionGanadora.Numeros.Item(2).Valor   'N2
                .Cells(1, 9).Value = datSorteo.CombinacionGanadora.Numeros.Item(3).Valor   'N3
                .Cells(1, 10).Value = datSorteo.CombinacionGanadora.Numeros.Item(4).Valor  'N4
                .Cells(1, 11).Value = datSorteo.CombinacionGanadora.Numeros.Item(5).Valor  'N5
                .Cells(1, 12).Value = datSorteo.Clave                    'Clave
            End With
        
        End Select
    End If
    '
    '   Eliminamos el rango
    '
    Set mRange = Nothing

  On Error GoTo 0
SetSorteo__CleanExit:
    Exit Sub
            
SetSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoEngine.SetSorteo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoEngine.SetSorteo", ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : SetSorteos
' Fecha          :
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
'
Public Sub SetSorteos(datSorteos As Sorteos)

  On Error GoTo SetSorteos_Error
    '
    '   Para cada sorteo en la colección
    '
    For Each mObj In datSorteos.Items
        '
        '   invocamos al metodo SetSorteo
        '
        SetSorteo mObj
        
    Next mObj
  On Error GoTo 0
SetSorteos__CleanExit:
    Exit Sub
            
SetSorteos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoEngine.SetSorteos", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoEngine.SetSorteos", ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : GetSorteosInIds
' Fecha          : do., 05/may/2019 19:13:12
' Propósito      : Devuelve una colección de sorteos entre dos ID
' Parámetros     : Id Inicial, Id final del sorteo y juego
'------------------------------------------------------------------------------*
'
Public Function GetSorteosInIds(datIdInicial As Integer, _
                                datIdFinal As Integer, _
                                datJuego As Juego) As Sorteos
    Dim mStr  As String
  On Error GoTo GetSorteosInIds_Error
    '
    '   Asignamos el juego al Motor
    '
    m_eJuego = datJuego
    '
    '   Segun el juego seleccionamos la tabla
    '
    mTabla = GetTablaSorteo(datJuego)
    '
    '   Obtenemos el rango de datos entre fechas
    '
    Set mRange = mDBase.SelectByIds(datIdInicial, datIdFinal, mTabla)
    '
    '   Si no hay error lo ha encontrado
    '
    If mDBase.ErrNumber = 0 Then
        '
        '   Creamos la colección
        '
        Set mCol = New Sorteos
        mCol.Juego = datJuego
        '
        '   Para cada fila en el rango creamos un sorteo
        '
        For Each mFila In mRange.Rows
            '
            '   Creamos el objeto Sorteo
            '
            Set mObj = New Sorteo
            '
            '   Asignamos atributos al sorteo
            '
            mObj.Juego = datJuego
            mObj.Constructor mFila
            '
            '   Agregamos sorteo a la colección
            '
            mCol.Add mObj
        Next
    Else
        '
        '   TODO: Agregar un error comprensible
        '
        mStr = Replace(MSG_SORTEOIDNOTFOUND, "#", CStr(datIdInicial))
        Err.Raise ERR_FECHASORTEONOTFOUND, "SorteoEngine.GetSorteosInIds", mStr

    End If
    '
    '   Devolvemos la colección
    '
    Set GetSorteosInIds = mCol
    '
    '   Eliminamos objetos
    '
    Set mCol = Nothing
    Set mObj = Nothing
    
    
  On Error GoTo 0
GetSorteosInIds__CleanExit:
    Exit Function
            
GetSorteosInIds_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoEngine.GetSorteosInIds", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoEngine.GetSorteosInIds", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetSorteosInFechas
' Fecha          : do., 05/may/2019 18:23:28
' Propósito      : Devuelve una colección de sorteos comprendidos entre dos
'                  fechas
' Parámetros     : Fecha Inicial del rango, fecha final del rango y Juego
'------------------------------------------------------------------------------*
'
Public Function GetSorteosInFechas(datFechaIni As Date, _
                                   datFechaFin As Date, _
                                   datJuego As Juego)
    Dim mStr  As String
  On Error GoTo GetSorteosInFechas_Error
    '
    '   Asignamos el juego al Motor
    '
    m_eJuego = datJuego
    '
    '   Segun el juego seleccionamos la tabla
    '
    mTabla = GetTablaSorteo(datJuego)
    '
    '   Obtenemos el rango de datos entre fechas
    '
    Set mRange = mDBase.SelectByFechas(datFechaIni, datFechaFin, mTabla)
    '
    '   Si no hay error lo ha encontrado
    '
    If mDBase.ErrNumber = 0 Then
        '
        '   Creamos la colección
        '
        Set mCol = New Sorteos
        mCol.Juego = datJuego
        '
        '   Para cada fila en el rango creamos un sorteo
        '
        For Each mFila In mRange.Rows
            '
            '   Creamos el objeto Sorteo
            '
            Set mObj = New Sorteo
            '
            '   Asignamos atributos al sorteo
            '
            mObj.Juego = datJuego
            mObj.Constructor mFila
            '
            '   Agregamos sorteo a la colección
            '
            mCol.Add mObj
        Next
    Else
        '
        '   TODO: Agregar un error comprensible
        '
        mStr = Replace(MSG_FECHASORTEONOTFOUND, "#", Format(datFechaIni, "dd/MM/yyyy"))
        Err.Raise ERR_FECHASORTEONOTFOUND, "SorteoEngine.SelectByFechas", mStr

    End If
    '
    '   Devolvemos la colección
    '
    Set GetSorteosInFechas = mCol
    '
    '   Eliminamos objetos
    '
    Set mCol = Nothing
    Set mObj = Nothing
    
    
  On Error GoTo 0
GetSorteosInFechas__CleanExit:
    Exit Function
            
GetSorteosInFechas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoEngine.GetSorteosInFechas", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoEngine.GetSorteosInFechas", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetFechaUltimoSorteo
' Fecha          : lu., 04/feb/2019 23:51:32
' Propósito      : Devuelve la fecha del último sorteo de la Base de datos
' Parámetros     : datJuego Juego del que queremos la fecha
'------------------------------------------------------------------------------*
'
Public Function GetFechaUltimoSorteo(datJuego As Juego) As Date
  On Error GoTo GetFechaUltimoSorteo_Error
    '
    '   Asignamos el juego al Motor
    '
    m_eJuego = datJuego
    '
    '   Segun el juego seleccionamos la tabla
    '
    mTabla = GetTablaSorteo(datJuego)
    '
    '   Buscamos el ID del último registro
    '
    mId = mDBase.GetLastID(mTabla)
    '
    '   Obtenemos el primer registro
    '
    Set mRange = mDBase.GetRowById(mId, mTabla)
    '
    '   Pasamos la fecha del primer registro
    '
    GetFechaUltimoSorteo = mRange.Cells(1, 3).Value   ' Fecha
    
    
  On Error GoTo 0
GetFechaUltimoSorteo__CleanExit:
  Exit Function
            
GetFechaUltimoSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoEngine.GetFechaUltimoSorteo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoEngine.GetFechaUltimoSorteo", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetFechaPrimerSorteo
' Fecha          : lu., 04/feb/2019 23:51:32
' Propósito      : Devuelve la fecha del primer sorteo de la Base de datos
' Parámetros     : Juego del que queremos la fecha
'------------------------------------------------------------------------------*
'
Public Function GetFechaPrimerSorteo(datJuego As Juego) As Date
  On Error GoTo GetFechaPrimerSorteo_Error
    '
    '   Asignamos el juego al Motor
    '
    m_eJuego = datJuego
    '
    '   Segun el juego seleccionamos la tabla
    '
    mTabla = GetTablaSorteo(datJuego)
    '
    '   Buscamos el ID del primer registro
    '
    mId = mDBase.GetFirstID(mTabla)
    '
    '   Obtenemos el primer registro
    '
    Set mRange = mDBase.GetRowById(mId, mTabla)
    '
    '   Pasamos la fecha del primer registro
    '
    GetFechaPrimerSorteo = mRange.Cells(1, 3).Value   ' Fecha
    
    
  On Error GoTo 0
GetFechaPrimerSorteo__CleanExit:
  Exit Function
            
GetFechaPrimerSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoEngine.GetFechaPrimerSorteo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoEngine.GetFechaPrimerSorteo", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetIdUltimoSorteo
' Fecha          : lu., 04/feb/2019 23:51:32
' Propósito      : Devuelve el Id del último sorteo
' Parámetros     : Juego del que queremos la fecha
'------------------------------------------------------------------------------*
'
Public Function GetIdUltimoSorteo(datJuego As Juego) As Integer
  On Error GoTo GetIdUltimoSorteo_Error
    '
    '   Asignamos el juego al Motor
    '
    m_eJuego = datJuego
    '
    '   Segun el juego seleccionamos la tabla
    '
    mTabla = GetTablaSorteo(datJuego)
    '
    '   Buscamos el ID del último registro
    '
    GetIdUltimoSorteo = mDBase.GetLastID(mTabla)
    
  On Error GoTo 0
GetIdUltimoSorteo__CleanExit:
  Exit Function
            
GetIdUltimoSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoEngine.GetIdUltimoSorteo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoEngine.GetIdUltimoSorteo", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetIdPrimerSorteo
' Fecha          : lu., 04/feb/2019 23:51:32
' Propósito      : Devuelve el Id del primer  sorteo
' Parámetros     : Juego del que queremos la fecha
'------------------------------------------------------------------------------*
'
Public Function GetIdPrimerSorteo(datJuego As Juego) As Integer
  On Error GoTo GetIdPrimerSorteo_Error
    '
    '   Asignamos el juego al Motor
    '
    m_eJuego = datJuego
    '
    '   Segun el juego seleccionamos la tabla
    '
    mTabla = GetTablaSorteo(datJuego)
    '
    '   Buscamos el ID del último registro
    '
    GetIdPrimerSorteo = mDBase.GetFirstID(mTabla)
  
  On Error GoTo 0
GetIdPrimerSorteo__CleanExit:
  Exit Function
            
GetIdPrimerSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "SorteoEngine.GetIdPrimerSorteo", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "SorteoEngine.GetIdPrimerSorteo", ErrDescription
End Function
' *===========(EOF): SorteoEngine.cls

