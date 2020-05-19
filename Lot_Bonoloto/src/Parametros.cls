VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : Parametros
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
' Version   :
'---------------------------------------------------------------------------------------
Option Explicit
Private m_iCount As Integer



'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Count() As Integer

   On Error GoTo Count_Error

    Count = m_iCount

   On Error GoTo 0
   Exit Property

Count_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.Count")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Add
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Add(datParametro As Parametro) As Integer

    Dim iResult As Integer


   On Error GoTo Add_Error

    Add = iResult

   On Error GoTo 0
   Exit Function

Add_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.Add")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : Delete
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Delete(datIndex As Integer) As Boolean

    Dim bResult As Boolean


   On Error GoTo Delete_Error

    Delete = bResult

   On Error GoTo 0
   Exit Function

Delete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.Delete")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : Item
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Item(datIndex As Integer) As Parametro

    Dim objResult As Parametro


   On Error GoTo Item_Error

    Set Item = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

Item_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.Item")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : MarkForDelete
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub MarkForDelete(datIndex As Integer)

   On Error GoTo MarkForDelete_Error

    

   On Error GoTo 0
   Exit Sub

MarkForDelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.MarkForDelete")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Undelete
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Undelete(datIndex As Integer)

   On Error GoTo Undelete_Error

    

   On Error GoTo 0
   Exit Sub

Undelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.Undelete")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Update
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Update(datIndex As Integer)

   On Error GoTo Update_Error

    

   On Error GoTo 0
   Exit Sub

Update_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.Update")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Clear()

   On Error GoTo Clear_Error

    

   On Error GoTo 0
   Exit Sub

Clear_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Parametros.Clear")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
