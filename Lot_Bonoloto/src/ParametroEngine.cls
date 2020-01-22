VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParametroEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ParametroEngine
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
' Version   :
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : GetNewParametro
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetNewParametro() As Parametro

    Dim objResult As Parametro


   On Error GoTo GetNewParametro_Error

    Set GetNewParametro = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

GetNewParametro_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "ParametroEngine.GetNewParametro")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function

'---------------------------------------------------------------------------------------
' Procedure : GetParametro
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetParametro(datNombreParametro As String) As Parametro

    Dim objResult As Parametro


   On Error GoTo GetParametro_Error

    Set GetParametro = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

GetParametro_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "ParametroEngine.GetParametro")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function

'---------------------------------------------------------------------------------------
' Procedure : GetParametros
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
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
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function

'---------------------------------------------------------------------------------------
' Procedure : GetListaVariable
' Author    : CHARLY
' Date      : 06/08/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetListaVariable(datNombreParametro As String) As Variant

    Dim vResult As Variant


   On Error GoTo GetListaVariable_Error

    GetListaVariable = vResult

   On Error GoTo 0
   Exit Function

GetListaVariable_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "ParametroEngine.GetListaVariable")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function
