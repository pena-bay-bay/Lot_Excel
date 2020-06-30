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
        Set oSorteo = Me.Item(inx)
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
'---------------------------------------------------------------------------------------
' Procedure : Item
' Author    : CHARLY
' Date      : 27/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Item(ByVal Index As Variant) As Sorteo

    Dim mintCodeID As Integer
   On Error GoTo Item_Error

    Set Item = Nothing
    
    If mCol Is Nothing Then
        GoTo Item_CleanExit
    End If
    
    If mCol.Count = 0 Then
        GoTo Item_CleanExit
    End If
    
    'Make Sure you cast the index to a Integer otherwise sets won't work
    If Trim(Index & "") = "" Or Index <= 0 Then
        'raise an error?
        GoTo Item_CleanExit
    Else
        mintCodeID = CInt(Index)
        Set Item = mCol.Item(mintCodeID)
    End If

Item_CleanExit:
   On Error GoTo 0
   Exit Property

Item_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteos.Item")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Load
' Author    : CHARLY
' Date      : 27/03/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Friend Function Load(ByVal FilledStorage As Variant) As Variant

   On Error GoTo Load_Error

    Load = False

    Set mCol = FilledStorage

    Load = True

   On Error GoTo 0
   Exit Function

Load_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Sorteos.Load")
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
' *     Procedure  : Remove
' *     Fichero    : Socios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 28/02/2009
' *     Asunto     :
' *============================================================================*
'
Public Function Remove(Optional ByVal Index As Variant) As Variant
   On Error GoTo Remove_Error

    Remove = False

    If IsMissing(Index) Then
        Me.Clear
        Remove = True
        GoTo Remove_CleanExit
    End If

    If (Not IsNumeric(Index)) Or (Index < 1 Or Index > Me.Count) Then
'        Err.Raise ERR_COLINDEXOUTOFRANGE, "Socios.Remove PROC", "Remove Failed: Index out of range."
        GoTo Remove_CleanExit
    End If

    mCol.Remove (Index)
    Remove = True

Remove_CleanExit:
   On Error GoTo 0
    Exit Function

Remove_Error:

    Call HandleException(Err.Number, Err.Description, "Socios.Remove")

End Function

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

' *============================================================================*
' *     Procedure  : Update
' *     Fichero    : Socios
' *     Autor      : Carlos Almela Baeza
' *     Creacion   : sáb, 28/02/2009
' *     Asunto     :
' *============================================================================*
'
Public Function Update(Optional ByVal Index As Variant) As Variant
    Dim oSorteo         As Sorteo
    Dim LowerLimit      As Long
    Dim UpperLimit      As Long
    Dim inx             As Long
    
   On Error GoTo Update_Error

    Update = False
    
    If Me.Count = 0 Then    'Nothing to Update
        Update = True
        GoTo Update_CleanExit
    End If

    'If Index is supplied then update only the supplied record
    If Not IsMissing(Index) Then
        If Index < 1 Or Index > Me.Count Then
'             Err.Raise ERR_COLINDEXOUTOFRANGE, "ColAccount.Update PROC", "Index out of range"
             GoTo Update_CleanExit
        Else
            LowerLimit = Index
            UpperLimit = Index
        End If
    Else
        LowerLimit = 1
        UpperLimit = Me.Count
    End If

    For inx = LowerLimit To UpperLimit
        Set oSorteo = Me.Item(inx)
        If oSorteo.EntidadNegocio.IsDirty Then
            With oSorteo
                .EntidadNegocio.ClassStorage = True
                '
                '   Propiedades a columnas
                '
                '
                                
                .EntidadNegocio.ClassStorage = False
            End With
            
'            'Check to see if updating existing record to database
'            If oSocio.IsNew = False Then
'                'Update the record
'                spname = SP_U_ACCOUNT
'            Else
'                'Inserting the record
'                spname = SP_I_ACCOUNT
'            End If
            
'            If Not oDALEng.Execute(SecurityToken, spname, vParameters, rsAccount) Then
'                Err.Raise ERR_ACCOUNTDALUPDATEFAILED, "ColAccount.Update PROC", "Update to Database Failed. SPName was: " & spname
'                GoTo CleanExit
'            ElseIf (Not rsAccount Is Nothing) Then  'Set to True if nothing returned in DB
'                If rsAccount.RecordCount > 1 Then
'                    Err.Raise ERR_ACCOUNTUPDATEFAILEDRETURNTOOMANYRECORDS, "ColAccount.Update PROC", "Update to Database Failed. Returned more than one record. SPName was: " & spname
'                    GoTo CleanExit
'                Else
'                    'Update collection with returned data
'                    With oCAccount
'                        .ClassStorage = True    'Set this to False to Enable Data Validation
'                        .AccountID = rsAccount.Fields(FN_ACCOUNTACCOUNTID)
'                        .Number = rsAccount.Fields(FN_ACCOUNTNUMBER)
'                        .StatusID = rsAccount.Fields(FN_ACCOUNTSTATUSID)
'                        .UserNumber = rsAccount.Fields(FN_ACCOUNTUSERNUMBER)
'                        .EnrollDate = rsAccount.Fields(FN_ACCOUNTENROLLDATE)
'                        .DenrolledDate = rsAccount.Fields(FN_ACCOUNTDENROLLEDDATE)
'                        .RecordTimestamp = rsAccount.Fields(FN_ACCOUNTRECORDTIMESTAMP)
'                        .CreateDateTime = rsAccount.Fields(FN_ACCOUNTCREATEDATETIME)
'                        .DeleteFlag = rsAccount.Fields(FN_ACCOUNTDELETEFLAG)
'                        .UpdateDateTime = rsAccount.Fields(FN_ACCOUNTUPDATEDATETIME)
'                        .IsNew = False
'                        .SecurityToken = SecurityToken
'                        .ClassStorage = False
'                        .Dirty = False
'                    End With
'                End If
'            End If
        End If 'Dirty
        Set oSorteo = Nothing
    Next inx

    Update = True

Update_CleanExit:
   On Error GoTo 0
    Exit Function

Update_Error:

    Call HandleException(Err.Number, Err.Description, "Socios.Update")

End Function


