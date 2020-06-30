VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SorteoEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : SorteoEngine
' Author    : CHARLY
' Date      : mar, 09/sep/2014 23:28:15
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit
Option Base 0
'
'
'
Private DB                     As New BdDatos           'Objeto Base de Datos

'---------------------------------------------------------------------------------------
' Procedure : GetSorteoByFecha
' Author    : CHARLY
' Date      : mar, 09/sep/2014 23:33:57
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetSorteoByFecha(vNewValue As Date) As Sorteo
    Dim rgSorteos As Range
    Dim oSorteo   As Sorteo
    Dim oFila     As Range
    Dim sKey      As String
   
   On Error Resume Next
    

    Set rgSorteos = DB.Resultados_Fechas(vNewValue, _
                                         vNewValue)
    If Not (rgSorteos Is Nothing) Then
        Set oSorteo = New Sorteo
        oSorteo.Constructor rgSorteos.Rows
        If oSorteo.Fecha = vNewValue Then
            Set GetSorteoByFecha = oSorteo
        Else
            Set GetSorteoByFecha = Nothing
        End If
    Else
        Set GetSorteoByFecha = Nothing
    End If
                
  On Error GoTo 0
    Exit Function
            
GetSorteoByFecha_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, ErrSource, "SorteoEngine.GetSorteoByFecha")
   '   Lanza el error
   Err.Raise ErrNumber, "SorteoEngine.GetSorteoByFecha", ErrDescription
End Function

