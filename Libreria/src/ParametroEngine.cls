VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParametroEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : ParametroEngine.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creaci�n   : vi., 01/feb/2019 23:22:35
' *     Versi�n    : 1.0
' *     Prop�sito  : Proporcionar y almacenar parametros de aplicaci�n
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mDBase      As DataBaseExcel
Private mCol        As Parametros
Private mPar        As Parametro
Private mRange      As Range

'--- Constantes ---------------------------------------------------------------*
'--- Mensajes -----------------------------------------------------------------*
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

'--- Propiedades --------------------------------------------------------------*
'--- Metodos Publicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GetNewParametro
' Fecha          : vi., 01/feb/2019 23:23:53
' Prop�sito      : Obtiene un parametro nuevo para agregar al repositorio
' Retorno        : Objeto Parametro
'------------------------------------------------------------------------------*
'
Public Function GetNewParametro() As Parametro
    
 On Error GoTo GetNewParametro_Error
    '
    '   Creamos el parametro nuevo
    '
    Set mPar = New Parametro
    '
    '   Inicializamos el estado del objeto
    '
    With mPar
        .Descripcion = LT_PARAMSINDESCRIPCION
        .Nombre = LT_PARAMSINNOMBRE
        .EntidadNegocio.IsNew = True
        '
        '   Calculamos el ID
        '
        .Id = mDBase.GetLastID(tblParametros)
    End With
    '
    '   Devolvemos el objeto
    '
    Set GetNewParametro = mPar
    Set mPar = Nothing

   On Error GoTo 0
   Exit Function

GetNewParametro_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "ParametroEngine.GetNewParametro")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetParametroById
' Fecha          : lu., 04/feb/2019 23:48:09
' Prop�sito      : Obtiene un parametros con el id de la variable
' Par�metros     : datId Clave �nica de la variable
' Retorno        : Objeto Parametro
'------------------------------------------------------------------------------*
'
Public Function GetParametroById(datId As Integer) As Parametro

 On Error GoTo GetParametroById_Error
    '
    '   Consultamos la base de datos
    '
    'On Error Resume Next
    Set mRange = mDBase.GetRowById(datId, tblParametros)
    If mDBase.ErrNumber = 0 Then
        '
        '   Creamos el parametro nuevo
        '
        Set mPar = New Parametro
        '
        '   Asignamos atributos al
        '
        With mPar
            .EntidadNegocio.ClassStorage = True
            .Id = mRange.Cells(1, 1).Value              'N
            .Nombre = mRange.Cells(1.2).Value           'Nombre
            .Orden = mRange.Cells(1, 3).Value           'Orden
            .Valor = mRange.Cells(1, 4).Value           'Valor
            .Tipo = mRange.Cells(1, 5).Value            'Tipo
            .Descripcion = mRange.Cells(1, 6).Value     'Concepto
            .FechaAlta = mRange.Cells(1, 7).Value       'FechaAlta
            .FechaModificacion = mRange.Cells(1, 8).Value 'FechaModificacion
        End With
        mPar.EntidadNegocio.IsNew = False
        mPar.EntidadNegocio.ClassStorage = False
    Else
        Err.Raise ERR_REGISTRONOTFOUND, "GetParametroById", MSG_REGISTRONOTFOUND
    End If
    '
    '   Devolvemos el objeto
    '
    Set GetParametroById = mPar

   On Error GoTo 0
   Exit Function

GetParametroById_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "ParametroEngine.GetParametroById")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetParametroByName
' Fecha          : vi., 01/feb/2019 23:28:48
' Prop�sito      : Obtiene un parametros con el nombre de la variable
' Par�metros     : datNombreParametro Nombre de la variable
' Retorno        : Colecci�n de parametros
'------------------------------------------------------------------------------*
'
Public Function GetParametroByName(datNombreParametro As String) As Parametros
 
 On Error GoTo GetParametro_Error
    '
    '   Accedemos a la tabla por la columna (1: Nombre
    '
    Set mRange = mDBase.GetRowByColumn(datNombreParametro, 1, tblParametros)
    '
    '   Si no hay error lo ha enconrado
    '
    If mDBase.ErrNumber = 0 Then
        '
        '   Creamos la colecci�n
        '
        Set mCol = New Parametros
        '
        '   Creamos el parametro nuevo
        '
        Set mPar = New Parametro
        '
        '   Asignamos atributos al parametro
        '
        With mPar
            .EntidadNegocio.ClassStorage = True
            .Id = mRange.Cells(1, 1).Value              'N
            .Nombre = mRange.Cells(1, 2).Value           'Nombre
            .Orden = mRange.Cells(1, 3).Value           'Orden
            .Valor = mRange.Cells(1, 4).Value           'Valor
            .Tipo = mRange.Cells(1, 5).Value            'Tipo
            .Descripcion = mRange.Cells(1, 6).Value     'Concepto
            .FechaAlta = mRange.Cells(1, 7).Value       'FechaAlta
            .FechaModificacion = mRange.Cells(1, 8).Value 'FechaModificacion
        End With
        mPar.EntidadNegocio.IsNew = False
        mPar.EntidadNegocio.ClassStorage = False
        '
        '   Agregamos el parametro a la colecci�n
        '
        mCol.Add mPar
    Else
        Err.Raise ERR_REGISTRONOTFOUND, "GetParametroById", MSG_REGISTRONOTFOUND
    End If
    '
    '   Devolvemos el objeto
    '
    Set GetParametroByName = mCol
    '
    '   Destruimos el parametro y colecci�n
    '
    Set mPar = Nothing
    
 On Error GoTo 0
    Exit Function

GetParametro_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "ParametroEngine.GetParametroByName")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : SetParametro
' Fecha          : lu., 04/feb/2019 23:51:32
' Prop�sito      : Almacena un parametro en la base de datos seg�n su estado
' Par�metros     : datParametro Objeto a almacenar
'------------------------------------------------------------------------------*
'
Public Sub SetParametro(datParametro As Parametro)
    Dim mRange As Range
    
 On Error GoTo SetParametro_Error
    '
    '   Evaluamos si el parametro existe o no
    '
    If datParametro.EntidadNegocio.IsNew Then
        '
        '   Obtenemos la �ltima fila de la tabla
        '
        Set mRange = mDBase.GetLastRow(tblParametros)
        '
        '   Obtenido el rango asignamos datos a cada celda
        '
        With mRange
            .Cells(1, 1).Value = datParametro.Id                'N
            .Cells(1, 2).Value = datParametro.Nombre            'Nombre
            .Cells(1, 3).Value = datParametro.Orden             'Orden
            .Cells(1, 4).Value = datParametro.Valor             'Valor
            .Cells(1, 5).Value = datParametro.Tipo              'Tipo
            .Cells(1, 6).Value = datParametro.Descripcion       'Concepto
            .Cells(1, 7).Value = datParametro.FechaAlta         'FechaAlta
            .Cells(1, 8).Value = datParametro.FechaModificacion 'FechaModificacion
        End With
    
    ElseIf datParametro.EntidadNegocio.MarkForDelete Then
        '
        '   Si est� marcado para borrarse se elimina la fila
        '
        mDBase.DeleteRow datParametro.Id, tblParametros
        
    Else
        '
        '   Obtenemos la fila con el ID
        '
        Set mRange = mDBase.GetRowById(datParametro.Id, tblParametros)
        '
        '   Obtenido el rango asignamos datos a cada celda
        '
        With mRange
'           .Cells(1, 1).Value = datParametro.ID                'N
            .Cells(1, 2).Value = datParametro.Nombre            'Nombre
            .Cells(1, 3).Value = datParametro.Orden             'Orden
            .Cells(1, 4).Value = datParametro.Valor             'Valor
            .Cells(1, 5).Value = datParametro.Tipo              'Tipo
            .Cells(1, 6).Value = datParametro.Descripcion       'Concepto
'           .Cells(1, 7).Value = datParametro.FechaAlta         'FechaAlta
            .Cells(1, 8).Value = datParametro.FechaModificacion 'FechaModificacion
        End With
    End If
    '
    '   Eliminamos el rango
    '
    Set mRange = Nothing
   On Error GoTo 0
   Exit Sub

SetParametro_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "ParametroEngine.SetParametro")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub


Public Function GetParametros(datNombreParametro As String) As Parametros
    Dim objResult As Parametros
   On Error GoTo GetParametros_Error

    Set GetParametros = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

GetParametros_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "ParametroEngine.GetParametros")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetListaVariable
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   : Devuelve una lista de variables en una matriz
'---------------------------------------------------------------------------------------
'
Public Function GetListaVariable(datNombreParametro As String) As Variant
    Dim vResult As Variant
 
 On Error GoTo GetListaVariable_Error

    Err.Raise ERR_TODO, "ParametroEngine.GetListaVariable", MSG_TODO

 On Error GoTo 0
    Exit Function

GetListaVariable_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "ParametroEngine.GetListaVariable")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetTipoParametros
' Fecha          : ma., 05/feb/2019 00:07:50
' Prop�sito      : Devuelve en una matriz los nombres de tipo de parametros
' Retrono        : Matriz con Nombres de tipo
'------------------------------------------------------------------------------*
'
Public Function GetTipoParametros() As Variant
    Dim mMatriz As Variant
    mMatriz = Split(LT_NOMBRESTIPOSPARAMETROS, ";")
    GetTipoParametros = mMatriz
End Function


' *===========(EOF): ParametroEngine.cls
