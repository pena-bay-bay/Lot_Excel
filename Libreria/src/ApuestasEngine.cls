VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApuestasEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : ApuestasEngine.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : sá., 10/ago/2019 20:25:55
' *     Versión    : 1.0
' *     Propósito  : Motor
' *
' *         TODO: Resolver la autoasignación de ID si es nuevo y está
' *               en colección, o si es new y ya existe
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Constantes ---------------------------------------------------------------*
Private Const ERR_APUESTANOTFOUND = 3051
Private Const ERR_APUESTAIDNOTFOUND = 3052
Private Const ERR_FECHASAPUESTANOTFOUND = 3053
Private Const ERR_APUESTADUPLICADO = 3054
Private Const ERR_APUESTASBOLETONOTFOUND = 3055
'--- Mensajes -----------------------------------------------------------------*
Private Const MSG_APUESTANOTFOUND As String = "Apuesta no encontranda"
Private Const MSG_APUESTAIDNOTFOUND As String = "Apuesta de Id: #, no encontrada"
Private Const MSG_FECHASAPUESTANOTFOUND  As String = "Apuesta de fecha:#, no encontrada"
Private Const MSG_APUESTADUPLICADO  As String = "Apuesta de fecha: # ya existe"
Private Const MSG_APUESTASBOLETONOTFOUND  As String = "Apuesta de boleto: # no encontradas"
'--- Variables Privadas -------------------------------------------------------*
Private mDBase      As DataBaseExcel
Private mCol        As Apuestas
Private mObj        As Apuesta
Private mRange      As Range
Private mFila       As Range
Private m_eJuego    As Juego
Private mId         As Integer
Private mTabla      As NombreTabla

'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'--- Métodos Privados ---------------------------------------------------------*
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

'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GetNewApuesta
' Fecha          : vi., 23/ago/2019 20:31:27
' Propósito      : Devuelve un objeto para cumplimentar sus propiedades
' Parámetros     : Juego
' Retorno        : Apuesta
'------------------------------------------------------------------------------*
Public Function GetNewApuesta(datJuego As Juego) As Apuesta
  On Error GoTo GetNewApuesta_Error
    '
    '   Definimos el objeto Apuesta
    '
    Set mObj = New Apuesta
    '
    '   Cumplimentamos las propiedades
    '
    With mObj
        .Juego = datJuego
        .EntidadNegocio.IsNew = True
        .Id = GetIdUltimoApuesta + 1
    End With
    '
    '   Devolvemos el objeto
    '
    Set GetNewApuesta = mObj
    '
    '   Destruimos los objetos internos
    '
    Set mObj = Nothing
    
  On Error GoTo 0
GetNewApuesta__CleanExit:
    Exit Function
            
GetNewApuesta_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GetNewApuesta.GetNewApuesta", ErrSource)
    Err.Raise ErrNumber, "GetNewApuesta.GetNewApuesta", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetApuestaById
' Fecha          : vi., 23/ago/2019 20:40:40
' Propósito      : Devuelve una apuesta por su ID
' Parámetros     : Id de la apuesta
' Retorno        : Objeto Apuesta
'------------------------------------------------------------------------------*
Public Function GetApuestaById(datId As Integer) As Apuesta
    Dim i As Integer
    Dim mNum As Numero
    Dim mStr As String
    Dim mVar As Variant
    
  On Error GoTo GetApuestaById_Error
    '
    '   Buscamos la fila del ID
    '
    Set mRange = mDBase.GetRowById(datId, tblApuestas)
        '
    '   Si no hay error lo ha encontrado
    '
    If mDBase.ErrNumber = 0 Then
        '
        '   Creamos el objeto Sorteo
        '
        Set mObj = New Apuesta
        '
        '   Asignamos atributos al objeto
        '
        With mObj
            .Id = mRange.Cells(1, 1).Value
            .IdBoleto = mRange.Cells(1, 2).Value
            .Juego = mRange.Cells(1, 3).Value
            .FechaSorteo = mRange.Cells(1, 4).Value
            .EsMultiple = IIf(mRange.Cells(1, 5).Value = "Si", True, False)
            For i = 1 To 11
                mVar = mRange.Cells(1, 5 + i).Value
                If Not IsEmpty(mVar) And IsNumeric(mVar) Then
                    Set mNum = New Numero
                    mNum.Valor = CInt(mVar)
                    .Numeros.Add mNum
                End If
            Next i
            For i = 1 To 5
                mVar = mRange.Cells(1, 16 + i).Value
                If Not IsEmpty(mVar) And IsNumeric(mVar) Then
                    Set mNum = New Numero
                    mNum.Valor = CInt(mVar)
                    .Estrellas.Add mNum
                End If
            Next i
            .Coste = mRange.Cells(1, 22).Value
            .ImportePremios = mRange.Cells(1, 23).Value
        End With
    Else
        mStr = Replace(MSG_APUESTANOTFOUND, "#", CStr(datId))
        Err.Raise ERR_APUESTAIDNOTFOUND, "SorteoEngine.GetApuestaById", mStr
    End If
    '
    '   Devolvemos el objeto
    '
    Set GetApuestaById = mObj
    '
    '   Destruimos el sorteo
    '
    Set mObj = Nothing
  
  On Error GoTo 0
GetApuestaById__CleanExit:
    Exit Function
            
GetApuestaById_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GetApuestaById.GetApuestaById", ErrSource)
    Err.Raise ErrNumber, "GetApuestaById.GetApuestaById", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : SetApuesta
' Fecha          : ma., 27/ago/2019 19:20:15
' Propósito      : Guardar las propiedades de la apuesta en el repositorio
' Parámetros     : Apuesta
'------------------------------------------------------------------------------*
Public Sub SetApuesta(datApuesta As Apuesta)
    Dim i As Integer
    
  On Error GoTo SetApuesta_Error
    If datApuesta.EntidadNegocio.IsNew Then
        '
        '   Obtenemos la última fila de la tabla
        '
        Set mRange = mDBase.GetLastRow(tblApuestas)
        '
        '   Obtenido el rango asignamos datos a cada celda
        '   Id  IdBoleto    Juego   FechaSorteo EsMultiple  N1  N2  N3  N4  N5 _
        '   N6  N7  N8  N9  N10 N11 E1  E2  E3  E4  E5  Coste   ImportePremios

        With mRange
            .Cells(1, 1).Value = datApuesta.Id                       'Id
            .Cells(1, 2).Value = datApuesta.IdBoleto                 'IdBoleto
            .Cells(1, 3).Value = datApuesta.Juego                    'Juego
            .Cells(1, 4).Value = datApuesta.FechaSorteo              'FechaSorteo
            .Cells(1, 5).Value = IIf(datApuesta.EsMultiple, "Si", "No") 'EsMultiple
            For i = 1 To 11                  'N1 N2 N3 N4 N5 N6 N7 N8 N9 N10 N11
                If i <= datApuesta.Numeros.Count Then
                    .Cells(1, 5 + i).Value = datApuesta.Numeros.Numeros.Item(i).Valor
                Else
                    .Cells(1, 5 + i).Value = Empty
                End If
            Next i
            For i = 1 To 5                    ' E1  E2  E3  E4  E5
                If i <= datApuesta.Estrellas.Count Then
                    .Cells(1, 16 + i).Value = datApuesta.Estrellas.Numeros.Item(i).Valor
                Else
                    .Cells(1, 16 + i).Value = Empty
                End If
            Next i
            .Cells(1, 22).Value = datApuesta.Coste                   ' Coste
            .Cells(1, 23).Value = datApuesta.ImportePremios          ' ImportePremios
        End With
    ElseIf datApuesta.EntidadNegocio.MarkForDelete Then
        '
        '   Si está marcado para borrarse se elimina la fila
        '
        mDBase.DeleteRow datApuesta.Id, tblApuestas
    Else
        '
        '   Obtenemos la fila con el ID
        '
        Set mRange = mDBase.GetRowById(datApuesta.Id, tblApuestas)
        With mRange
            .Cells(1, 2).Value = datApuesta.IdBoleto                 'IdBoleto
            .Cells(1, 3).Value = datApuesta.Juego                    'Juego
            .Cells(1, 4).Value = datApuesta.FechaSorteo              'FechaSorteo
            .Cells(1, 5).Value = IIf(datApuesta.EsMultiple, "Si", "No") 'EsMultiple
            For i = 1 To 11                  'N1 N2 N3 N4 N5 N6 N7 N8 N9 N10 N11
                If i <= datApuesta.Numeros.Count Then
                    .Cells(1, 5 + i).Value = datApuesta.Numeros.Numeros.Item(i).Valor
                Else
                    .Cells(1, 5 + i).Value = Empty
                End If
            Next i
            For i = 1 To 5                    ' E1  E2  E3  E4  E5
                If i <= datApuesta.Estrellas.Count Then
                    .Cells(1, 16 + i).Value = datApuesta.Estrellas.Numeros.Item(i).Valor
                Else
                    .Cells(1, 16 + i).Value = Empty
                End If
            Next i
            .Cells(1, 22).Value = datApuesta.Coste                   ' Coste
            .Cells(1, 23).Value = datApuesta.ImportePremios          ' ImportePremios
        End With
    End If
    '
    '   Eliminamos el rango
    '
    Set mRange = Nothing
    
  On Error GoTo 0
SetApuesta__CleanExit:
    Exit Sub
            
SetApuesta_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GetNewApuesta.SetApuesta", ErrSource)
    Err.Raise ErrNumber, "GetNewApuesta.SetApuesta", ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : SetApuestas
' Fecha          : ma., 27/ago/2019 19:24:56
' Propósito      : Guardar varias apuestas recursivamente
' Parámetros     : Colección de Apuestas
'------------------------------------------------------------------------------*
Public Sub SetApuestas(datApuestas As Apuestas)
  On Error GoTo SetApuestas_Error
    '
    '   Para cada sorteo en la colección
    '
    For Each mObj In datApuestas.Items
        '
        '   Si es una apuesta nueva recalculamos el ID
        '
        If mObj.EntidadNegocio.IsNew Then
            mObj.Id = GetIdUltimoApuesta + 1
        End If
        '
        '   invocamos al metodo SetSorteo
        '
        SetApuesta mObj
        
    Next mObj
  
  On Error GoTo 0
SetApuestas__CleanExit:
    Exit Sub
            
SetApuestas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "GetNewApuesta.SetApuestas", ErrSource)
    Err.Raise ErrNumber, "GetNewApuesta.SetApuestas", ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetApuestasBoleto
' Fecha          : sá., 12/oct/2019 12:50:00
' Propósito      : Devuelve las apuestas de un boleto
' Parámetros     : Id del Boleto
' Retorno        : Colección de apuestas
'------------------------------------------------------------------------------*
Public Function GetApuestasBoleto(datIdBoleto As Integer) As Apuestas
    Dim mTuplas As Collection
    Dim mTupl As TuplaAparicion
    Dim mStr As String
    
  On Error GoTo GetApuestasBoleto_Error
    '
    '   Obtenemos la colección de registros con el mismo Id de Boleto
    '
    Set mTuplas = mDBase.SelectApuestasByBoleto(datIdBoleto, tblApuestas)
    '
    '   Si no ha encontrado nada
    '
    If mDBase.ErrNumber <> 0 Then
        '
        '   No ha encontrado registros para el Id Solicado
        '
        mStr = Replace(MSG_APUESTASBOLETONOTFOUND, "#", CStr(datIdBoleto))
        Err.Raise ERR_APUESTASBOLETONOTFOUND, "ApuestasEngine.GetApuestasBoleto", mStr
    End If
    '
    '   Si ha encontrado registros
    '
    If mTuplas.Count > 0 Then
        '
        '   Creamos la colección
        '
        Set mCol = New Apuestas
        '
        '   Para cada fila en la co el rango creamos un sorteo
        '
        For Each mTupl In mTuplas
            '
            '   Creamos el objeto Sorteo
            '
            Set mObj = GetApuestaById(mTupl.NumeroRegistro)
            '
            '   Agregamos sorteo a la colección
            '
            mCol.Add mObj
        Next
    End If
    '
    '   Devolvemos la colección
    '
    Set GetApuestasBoleto = mCol
    '
    '   Eliminamos objetos
    '
    Set mCol = Nothing
    Set mObj = Nothing
  
  On Error GoTo 0
GetApuestasBoleto__CleanExit:
    Exit Function
            
GetApuestasBoleto_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ApuestasEngine.GetApuestasBoleto", ErrSource)
    Err.Raise ErrNumber, "ApuestasEngine.GetApuestasBoleto", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetApuestasInIds
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetApuestasInIds(datIni As Integer, datFin As Integer, _
                                datJuego As Juego) As Apuestas

  On Error GoTo GetApuestasInIds_Error
    '
    '
    Err.Raise ERR_TODO, "ApuestasEngine.GetApuestasInIds", MSG_TODO
    
  On Error GoTo 0
GetApuestasInIds__CleanExit:
    Exit Function
            
GetApuestasInIds_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ApuestasEngine.GetApuestasInIds", ErrSource)
    Err.Raise ErrNumber, "ApuestasEngine.GetApuestasInIds", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetApuestasInFechas
' Fecha          : dd/MMM/yyyy
' Propósito      :
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Function GetApuestasInFechas(datFechaIni As Date, datFechaFin As Date, _
                                   datJuego As Juego) As Apuestas

  On Error GoTo GetApuestasInFechas_Error
    '
    '
    Err.Raise ERR_TODO, "ApuestasEngine.GetApuestasInFechas", MSG_TODO
    
  On Error GoTo 0
GetApuestasInFechas__CleanExit:
    Exit Function
            
GetApuestasInFechas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ApuestasEngine.GetApuestasInFechas", ErrSource)
    Err.Raise ErrNumber, "ApuestasEngine.GetApuestasInFechas", ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetIdUltimaApuesta
' Fecha          :
' Propósito      : Devuelve el Id de la última apuesta
' Parámetros     :
'------------------------------------------------------------------------------*
'
Public Function GetIdUltimoApuesta() As Integer
  On Error GoTo GetIdUltimoApuesta_Error
    '
    '   Buscamos el ID del último registro
    '
    GetIdUltimoApuesta = mDBase.GetLastID(tblApuestas)
    
  On Error GoTo 0
GetIdUltimoApuesta__CleanExit:
  Exit Function
            
GetIdUltimoApuesta_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ApuestasEngine.GetIdUltimoApuesta", ErrSource)
    Err.Raise ErrNumber, "ApuestasEngine.GetIdUltimoApuesta", ErrDescription
End Function

'' *===========(EOF): BoletosEngine.cls
