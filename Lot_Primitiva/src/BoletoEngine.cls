VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoletoEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : GetNewBoleto
' Author    : Charly
' Date      : 11/06/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Procedure : GetNewBoleto
' Author    : Charly
' Date      : 11/06/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetNewBoleto(datTipoJuego As Juego) As Boleto
    Dim objResult As Boleto


  On Error GoTo GetNewBoleto_Error
    Set objResult = New Boleto
    objResult.Juego = datTipoJuego
    Set GetNewBoleto = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

GetNewBoleto_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "BoletoEngine.GetNewBoleto")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetBoletoById
' Author    : Charly
' Date      : 11/06/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetBoletoById(datId As Long) As Boleto

    Dim objResult As Boleto


  On Error GoTo GetBoletoById_Error

    Set GetBoletoById = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

GetBoletoById_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "BoletoEngine.GetBoletoById")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetBoletoByFecha
' Author    : Charly
' Date      : 11/06/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetBoletoByFecha(datFecha As Date, datTipoJuego As ModalidadJuego) As Boleto

    Dim objResult As Boleto


  On Error GoTo GetBoletoByFecha_Error

    Set GetBoletoByFecha = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

GetBoletoByFecha_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "BoletoEngine.GetBoletoByFecha")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : SearchBoletos
' Author    : Charly
' Date      : 11/06/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function SearchBoletos(datFechaIni As Date, datFechaFin As Date, datTipoJuego As ModalidadJuego) As Boleto

    Dim objResult As Boleto


  On Error GoTo SearchBoletos_Error

    Set SearchBoletos = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

SearchBoletos_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "BoletoEngine.SearchBoletos")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetDataSet
' Author    : Charly
' Date      : 11/06/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetDataSet(datTipoJuego As Juego) As Variant

    Dim vResult As Variant


  On Error GoTo GetDataSet_Error

    GetDataSet = vResult

   On Error GoTo 0
   Exit Function

GetDataSet_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "BoletoEngine.GetDataSet")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : Parse
' Author    : Charly
' Date      : 11/06/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Parse(datRecordset As Recordset) As Variant

    Dim vResult As Variant


  On Error GoTo Parse_Error

    Parse = vResult

   On Error GoTo 0
   Exit Function

Parse_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "BoletoEngine.Parse")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")

End Function

'---------------------------------------------------------------------------------------
' Procedure : DelBoleto
' Author    : Charly
' Date      : 11/06/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function DelBoleto(datId As Long) As Boolean

    Dim bResult As Boolean


  On Error GoTo DelBoleto_Error

    DelBoleto = bResult

   On Error GoTo 0
   Exit Function

DelBoleto_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "BoletoEngine.DelBoleto")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")

End Function

'---------------------------------------------------------------------------------------
' Procedure : SetBoleto
' Author    : Charly
' Date      : 11/06/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function SetBoleto(datBoleto As Boleto) As Boolean

    Dim bResult As Boolean


  On Error GoTo SetBoleto_Error

    SetBoleto = bResult

   On Error GoTo 0
   Exit Function

SetBoleto_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "BoletoEngine.SetBoleto")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")

End Function
