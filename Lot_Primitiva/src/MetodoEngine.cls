VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetodoEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : MetodoEngine
' Author    : Charly
' Date      : 19/03/2012
' Purpose   : Esta clase es un motor de metodos y suministra elementos relacionados con
'             el método
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : CargaOrdenaciones
' Author    : Charly
' Date      : 11/04/2012
' Purpose   : Carga el control desplegable de una ventana con la lista de valores de
'             Ordenaciones
'---------------------------------------------------------------------------------------
'
Public Sub CargaOrdenaciones(cmbLista As Object)
    Dim m_vMatriz       As Variant
    Dim m_sNombre       As String
    Dim i               As Integer
  On Error GoTo CargaOrdenaciones_Error

    m_vMatriz = Split(NOMBRES_ORDENACION, ";")
    
    'Si el tipo de objeto no es un control salir
    If Not (TypeOf cmbLista Is Control) Then
        Return
    End If

    'Inicializar el control
    cmbLista.Clear
       
    For i = 0 To UBound(m_vMatriz)
        m_sNombre = m_vMatriz(i)
        'A±adir elementos de la lista
        cmbLista.AddItem m_sNombre, i
    Next i
    

   On Error GoTo 0
   Exit Sub

CargaOrdenaciones_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "MetodoEngine.CargaOrdenaciones")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub
'---------------------------------------------------------------------------------------
' Procedure : CargaAgrupaciones
' Author    : Charly
' Date      : 11/04/2012
' Purpose   : Carga el control desplegable de una ventana con la lista de valores de
'             Agrupaciones
'---------------------------------------------------------------------------------------
'
Public Sub CargaAgrupaciones(cmbLista As Object)
    Dim m_vMatriz       As Variant
    Dim m_sNombre       As String
    Dim i               As Integer
  On Error GoTo CargaAgrupaciones_Error

    m_vMatriz = Split(NOMBRES_AGRUPACION, ";")
    
    'Si el tipo de objeto no es un control salir
    If Not (TypeOf cmbLista Is Control) Then
        Return
    End If

    'Inicializar el control
    cmbLista.Clear
       
    For i = 0 To UBound(m_vMatriz)
        m_sNombre = m_vMatriz(i)
        'A±adir elementos de la lista
        cmbLista.AddItem m_sNombre, i
    Next i
    

   On Error GoTo 0
   Exit Sub

CargaAgrupaciones_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "MetodoEngine.CargaAgrupaciones")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

