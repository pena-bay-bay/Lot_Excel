VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sorteos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : Sorteos
' Author    : CHARLY
' Date      : 27/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit
Option Base 0
Private mCol As Collection                  ' Usado para almacenar la Clase

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : CHARLY
' Date      : 27/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : CHARLY
' Date      : 27/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Add
' Author    : CHARLY
' Date      : 27/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Add(ByVal paramSorteo As Sorteo) As Variant
   On Error GoTo Add_Error

    Add = -1
    
    '
    '   Valida si el objeto es correcto
    If paramSorteo Is Nothing Then
'        Err.Raise ERR_SOCIOFAILEDADDCHECK, "ColAccount.Add PROC", "You must specify: "
        GoTo Add_CleanExit
    End If

    If mCol Is Nothing Then
        Set mCol = New Collection
    End If
    
    'Add the item to the End of the Collection
    If mCol.Count <> 0 Then
        mCol.Add paramSorteo, , , mCol.Count
    Else
        mCol.Add paramSorteo
    End If
    
    Add = mCol.Count
    
Add_CleanExit:
   On Error GoTo 0
   Exit Function

Add_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteos.Add")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : CHARLY
' Date      : 27/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
       
    Dim cls As Sorteo
   On Error GoTo Clear_Error

   On Error GoTo Clear_Error
    '
    '
    ' borra las subcolecciones
    For Each cls In Me
        Set cls.Combinacion = Nothing
    Next
    
    'Clear Collection
    Set mCol = Nothing

    'Create Collection
    Set mCol = New Collection

    Set cls = Nothing

   On Error GoTo 0
   Exit Sub

Clear_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteos.Clear")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : CHARLY
' Date      : 27/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Count() As Variant

   On Error GoTo Count_Error

   On Error GoTo Count_Error

    If mCol Is Nothing Then
        Count = 0
    Else
        Count = mCol.Count
    End If

   On Error GoTo 0
   Exit Property

Count_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteos.Count")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property
'---------------------------------------------------------------------------------------
' Procedure : Delete
' Author    : CHARLY
' Date      : 27/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Delete(Optional ByVal Index As Variant) As Variant
    Dim oSorteo As Sorteo
    Dim LowerLimit As Long
    Dim UpperLimit As Long
    Dim inx As Long

   On Error GoTo Delete_Error

    Delete = False
    
    If Me.Count = 0 Then    'Nothing to Update
        Delete = True
        GoTo Delete_CleanExit
    End If

    'If Index is supplied then Delete only the supplied record
    If Not IsMissing(Index) Then
        If Index < 1 Or Index > Me.Count Then
'             Err.Raise ERR_COLINDEXOUTOFRANGE, "ColAccount.Delete PROC", "Index out of range"
             GoTo Delete_CleanExit
        Else
            LowerLimit = Index
            UpperLimit = Index
        End If
    Else
        LowerLimit = 1
        UpperLimit = Me.Count
    End If

    For inx = UpperLimit To LowerLimit Step -1
        Set oSorteo = Me.Items(inx)
        '
        ' Elimina fila de
        '
'        If Not oSocio.IsNew Then
'            'Delete from DB If Not New
'            ReDim vParameters(PARMUBOUND, 1)
'
'            With oCAccount
'                'Fill Parameter Array
'                .ClassStorage = True
'                vParameters(PARMNAME, 0) = PARMNAMESP_ACCOUNTACCOUNTID    'Name
'                vParameters(PARMTYPE, 0) = PARMTYPESP_ACCOUNTACCOUNTID      'Type
'                vParameters(PARMLENGTH, 0) = 4     'Size
'                vParameters(PARMDIR, 0) = adInput     'Direction
'                vParameters(PARMVALUE, 0) = .AccountID     'Value
'
'                vParameters(PARMNAME, 1) = PARMNAMESP_ACCOUNTRECORDTIMESTAMP    'Name
'                vParameters(PARMTYPE, 1) = PARMTYPESP_ACCOUNTRECORDTIMESTAMP      'Type
'                vParameters(PARMLENGTH, 1) = 8     'Size
'                vParameters(PARMDIR, 1) = adInput     'Direction
'                vParameters(PARMVALUE, 1) = .RecordTimestamp     'Value
'                .ClassStorage = False
'            End With
'
'            If Not oDALEng.Execute(SecurityToken, SP_D_ACCOUNT, vParameters) Then
'                Err.Raise ERR_ACCOUNTDALDELETEFAILED, "ColAccount.Delete PROC", "Delete Failed. SPName was: " & SP_D_ACCOUNT
'                GoTo CleanExit
'            End If
'        End If
        Set oSorteo = Nothing

        'Remove from Collection
        mCol.Remove (inx)
    Next inx

    Delete = True

Delete_CleanExit:
   On Error GoTo 0
   Exit Function

Delete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteos.Delete")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

' *============================================================================*
' *     Procedure  : MarkForDelete
' *     Fichero    : Socios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 28/02/2009
' *     Asunto     :
' *============================================================================*
'
Public Function MarkForDelete(Optional ByVal Index As Variant) As Variant

    Dim LowerLimit As Long
    Dim UpperLimit As Long
    Dim inx As Long
    
   On Error GoTo MarkForDelete_Error

    MarkForDelete = False

    If Not IsMissing(Index) Then
        'Check Index
        If (Not IsNumeric(Index)) Or (Index < 1 Or Index > Me.Count) Then
'            Err.Raise ERR_COLINDEXOUTOFRANGE, "Socios.MarkForDelete PROC", "MarkForDelete Failed: Index out of range."
            GoTo MarkForDelete_CleanExit
        End If
        
        'Toggle DeleteFlag
'        Me.Item(Index).DeleteFlag = True
    Else
        LowerLimit = 1
        UpperLimit = Me.Count
        For inx = LowerLimit To UpperLimit
'            Me.Item(inx).DeleteFlag = True
        Next
    End If

    MarkForDelete = True

MarkForDelete_CleanExit:
   On Error GoTo 0
    Exit Function

MarkForDelete_Error:

    Call HandleException(Err.Number, Err.Description, "Socios.MarkForDelete")

End Function

' *============================================================================*
' *     Procedure  : Items
' *     Fichero    : Socios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : jue, 19/03/2009
' *     Asunto     :
' *============================================================================*
'
Public Property Get Items() As Collection
   On Error GoTo Items_Error

    Set Items = mCol

Items_CleanExit:
   On Error GoTo 0
    Exit Property

Items_Error:

    Call HandleException(Err.Number, Err.Description, "Socios.Items")
End Property

' *============================================================================*
' *     Procedure  : Undelete
' *     Fichero    : Socios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 28/02/2009
' *     Asunto     :
' *============================================================================*
'
Public Function Undelete(Optional ByVal Index As Variant) As Variant
    Dim LowerLimit As Long
    Dim UpperLimit As Long
    Dim inx As Long
    
   On Error GoTo Undelete_Error

    Undelete = False

    'Check Index
    If Not IsMissing(Index) Then
        If (Not IsNumeric(Index)) Or (Index < 1 Or Index > Me.Count) Then
'            Err.Raise ERR_COLINDEXOUTOFRANGE, "Socios.UnDelete PROC", "UnDelete Failed: Index out of range."
            GoTo Undelete_CleanExit
        End If
    
        'Toggle DeleteFlag
'        Me.Item(Index).DeleteFlag = False
    Else
        LowerLimit = 1
        UpperLimit = Me.Count
        For inx = LowerLimit To UpperLimit
'            Me.Item(inx).DeleteFlag = False
        Next
    End If

    Undelete = True

Undelete_CleanExit:
   On Error GoTo 0
    Exit Function

Undelete_Error:

    Call HandleException(Err.Number, Err.Description, "Socios.Undelete")

End Function

