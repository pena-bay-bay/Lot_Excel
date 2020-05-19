VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Premios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : Premios
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   : Colección de Premios de un sorteo
' Version   : 1.0
'---------------------------------------------------------------------------------------
Option Explicit
Option Base 0
'
'  Variables privadas
'
Private m_iIdSorteo As Integer
Private m_eJuego As Juego

'---------------------------------------------------------------------------------------
' Procedure : IdSorteo
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get IdSorteo() As Integer

   On Error GoTo IdSorteo_Error

    IdSorteo = m_iIdSorteo

   On Error GoTo 0
   Exit Property

IdSorteo_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.IdSorteo")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : IdSorteo
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let IdSorteo(ByVal iIdSorteo As Integer)

   On Error GoTo IdSorteo_Error

    m_iIdSorteo = iIdSorteo

   On Error GoTo 0
   Exit Property

IdSorteo_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.IdSorteo")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Juego
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Juego() As Juego

   On Error GoTo Juego_Error

    Juego = m_eJuego

   On Error GoTo 0
   Exit Property

Juego_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Juego")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Juego
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Juego(ByVal eJuego As Juego)

   On Error GoTo Juego_Error

    m_eJuego = eJuego

   On Error GoTo 0
   Exit Property

Juego_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Juego")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'---------------------------------------------------------------------------------------
' Procedure : Add
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Add() As Premio2

    Dim objResult As Premio2


   On Error GoTo Add_Error

    Set Add = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

Add_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Add")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : Delete
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Delete() As Boolean

    Dim bResult As Boolean


   On Error GoTo Delete_Error

    Delete = bResult

   On Error GoTo 0
   Exit Function

Delete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Delete")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'---------------------------------------------------------------------------------------
' Procedure : Item
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Item() As Premio2

    Dim objResult As Premio2


   On Error GoTo Item_Error

    Set Item = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

Item_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Item")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function

'---------------------------------------------------------------------------------------
' Procedure : Load
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Load() As Variant

    Dim vResult As Variant


   On Error GoTo Load_Error

    Load = vResult

   On Error GoTo 0
   Exit Function

Load_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Load")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function

'---------------------------------------------------------------------------------------
' Procedure : Update
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Update() As Boolean

    Dim bResult As Boolean


   On Error GoTo Update_Error

    Update = bResult

   On Error GoTo 0
   Exit Function

Update_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Update")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function

'---------------------------------------------------------------------------------------
' Procedure : Store
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Store() As Variant

    Dim vResult As Variant


   On Error GoTo Store_Error

    Store = vResult

   On Error GoTo 0
   Exit Function

Store_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Store")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function

'---------------------------------------------------------------------------------------
' Procedure : Remove
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Remove() As Boolean

    Dim bResult As Boolean


   On Error GoTo Remove_Error

    Remove = bResult

   On Error GoTo 0
   Exit Function

Remove_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Remove")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function

'---------------------------------------------------------------------------------------
' Procedure : MarkForDelete
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function MarkForDelete() As Boolean

    Dim bResult As Boolean


   On Error GoTo MarkForDelete_Error

    MarkForDelete = bResult

   On Error GoTo 0
   Exit Function

MarkForDelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.MarkForDelete")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function

'---------------------------------------------------------------------------------------
' Procedure : Undelete
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Undelete() As Integer

    Dim iResult As Integer


   On Error GoTo Undelete_Error

    Undelete = iResult

   On Error GoTo 0
   Exit Function

Undelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Undelete")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function

'---------------------------------------------------------------------------------------
' Procedure : IndexOf
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function IndexOf() As Premio2

    Dim objResult As Premio2


   On Error GoTo IndexOf_Error

    Set IndexOf = objResult
    Set objResult = Nothing

   On Error GoTo 0
   Exit Function

IndexOf_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.IndexOf")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function

'---------------------------------------------------------------------------------------
' Procedure : Conteins
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Conteins() As Boolean

    Dim bResult As Boolean


   On Error GoTo Conteins_Error

    Conteins = bResult

   On Error GoTo 0
   Exit Function

Conteins_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Conteins")
   '   Lanza el Error
   'Err.Raise ErrNumber, ErrSource, ErrDescription
   '   Informa del error
   'Call MsgBox(ErrDescription, vbError Or vbSystemModal, NOMBRE_APLICACION)
   'Call Trace("CERRAR")


End Function
