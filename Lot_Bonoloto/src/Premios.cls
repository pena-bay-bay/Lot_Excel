VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Premios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : Premios.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : do., 06/abr/2014 00:00:00
' *     Modificado : mi., 15/jul/2020 17:30:01
' *     Versión    : 1.1
' *     Propósito  : Representa la colección de premios de un sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private m_iIdSorteo As Integer
Private m_eJuego As Juego
Private mCol As Collection

'--- Constantes ---------------------------------------------------------------*
Private Const LT_PREMIOSDEFBONOLOTO As String = _
              "1,0,450000;2,0,100000;3,0,1000;4,0,30;5,0,4;15,0,0.5"
Private Const LT_PREMIOSDEFPRIMITIVA As String = _
              "14,0,10000000;1,0,1000000;2,0,75000;3,0,3000;4,0,85;5,0,8;15,0,1"
Private Const LT_PREMIOSDEFGORDO As String = _
              "1,0,1000000;2,0,200000;3,0,3000;4,0,250;5,0,50;6,0,20;7,0,8;" & _
              "8,0,3;15,0,1.5"
Private Const LT_PREMIOSDEFEUROMILLONES As String = _
              "1,0,1000000,0;2,0,250000,0;3,0,60000,0;4,0,2000,0;5,0,150,0;" & _
              "6,0,100,0;7,0,30,0;8,0,12,0;9,0,12,0;10,0,12,0;11,0,6,0;12,0" & _
              ",6,0;13,0,4,0"
              
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedure : IdSorteo
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'------------------------------------------------------------------------------*
'
Public Property Get IdSorteo() As Integer
    IdSorteo = m_iIdSorteo
End Property

Public Property Let IdSorteo(ByVal vNewValue As Integer)
    m_iIdSorteo = vNewValue
End Property


'------------------------------------------------------------------------------*
' Procedure : Juego
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'------------------------------------------------------------------------------*
'
Public Property Get Juego() As Juego
    Juego = m_eJuego
End Property

Public Property Let Juego(ByVal vNewValue As Juego)
    If vNewValue <> m_eJuego Then
        m_eJuego = vNewValue
        SetPremiosDefecto
    Else
        m_eJuego = vNewValue
    End If
End Property

Public Property Get Count() As Integer
    Count = mCol.Count
End Property
'------------------------------------------------------------------------------*
' Procedimiento  : Items
' Fecha          : do., 10/feb/2019 09:54:46
' Propósito      : Devolver la colección interna
' Retorno        : Colección de Objetos
'------------------------------------------------------------------------------*
Public Property Get Items() As Collection
    Set Items = mCol
End Property

Public Property Let Items(ByVal vNewValue As Collection)
    Set mCol = vNewValue
End Property

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mCol = New Collection
    m_eJuego = JUEGO_DEFECTO
    SetPremiosDefecto
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : SetPremiosDefecto
' Fecha          : mi., 15/jul/2020 17:33:15
' Propósito      : Establece una colección de premios por defecto
'------------------------------------------------------------------------------*
'
Private Sub SetPremiosDefecto()
 On Error GoTo SetPremiosDefecto_Error
    '
    '   Seleccionamos los premios y categorias según el juego
    '
    Select Case m_eJuego
        Case Bonoloto: Parse LT_PREMIOSDEFBONOLOTO
        Case LoteriaPrimitiva: Parse LT_PREMIOSDEFPRIMITIVA
        Case gordoPrimitiva: Parse LT_PREMIOSDEFGORDO
        Case Euromillones: Parse LT_PREMIOSDEFEUROMILLONES
    End Select
    
  On Error GoTo 0
    Exit Sub
SetPremiosDefecto_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Premios.SetPremiosDefecto", ErrSource)
    Err.Raise ErrNumber, "Premios.SetPremiosDefecto", ErrDescription
End Sub



'--- Métodos Públicos ---------------------------------------------------------*

'------------------------------------------------------------------------------*
' Procedimiento  : Parse
' Fecha          : mi., 15/jul/2020 17:33:15
' Propósito      : Interpreta la información de los premios de cadena
'                  a colección
' Parámetros     : Cadena de caracteres con los valores de los premios
'------------------------------------------------------------------------------*
'
Public Sub Parse(vNewValue As String)
    Dim mVar As Variant
    Dim i    As Integer
    Dim oPrm As Premio
    
 On Error GoTo Parse_Error
    '
    '   Inicializamos colección
    '
    Set mCol = New Collection
    '
    '   obtenemos las tuplas de premios
    '
    mVar = Split(vNewValue, ";")
    
    For i = 0 To UBound(mVar)
        '
        '   Creamos un premio
        '
        Set oPrm = New Premio
        '
        '   asignamos Id
        '
        oPrm.Id = i
        oPrm.Juego = m_eJuego
        oPrm.UnPack CStr(mVar(i))
        '
        '   agregamos a la colección
        '
        mCol.Add oPrm, Str(oPrm.CategoriaPremio)
    Next i
  
  
  On Error GoTo 0
    Exit Sub
Parse_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Premios.Parse", ErrSource)
    Err.Raise ErrNumber, "Premios.Parse", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : mi., 15/jul/2020 18:47:20
' Propósito      : Componer los premios de la colección en una cadena
' Parámetros     : Cadena de caracteres con los valores de los premios
'------------------------------------------------------------------------------*
'
Public Function ToString() As String
    Dim mStr As String
    Dim mPrm As Premio
 On Error GoTo ToString_Error
    '
    '   Inicializamos la cadena
    '
    mStr = ""
    '
    '   Recorremos la colección
    '
    For Each mPrm In mCol
        mStr = mStr & mPrm.Pack & ";"
    Next mPrm
    
    mStr = Left(mStr, Len(mStr) - 1)
    ToString = mStr
    
 On Error GoTo 0
   Exit Function

ToString_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Premios.ToString")
   Err.Raise ErrNumber, "Premios.ToString", ErrDescription
End Function



'------------------------------------------------------------------------------*
' Procedimiento  : GetPremioByCategoria
' Fecha          : mi., 15/jul/2020 18:49:52
' Propósito      : Obtiene el premio de una categoria
' Parámetros     :
'------------------------------------------------------------------------------*
'
Public Function GetPremioByCategoria(vNewValue As CategoriaPremio) As Premio
    Dim mPrm As Premio
    
  On Error GoTo GetPremioByCategoria_Error
    '
    ' Creamos la variable
    '
    Set mPrm = New Premio
    Set mPrm = mCol.Item(Str(vNewValue))
    
    Set GetPremioByCategoria = mPrm
    Set mPrm = Nothing
    
 On Error GoTo 0
   Exit Function

GetPremioByCategoria_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Premios.GetPremioByCategoria")
   Err.Raise ErrNumber, "Premios.GetPremioByCategoria", ErrDescription
End Function



'---------------------------------------------------------------------------------------
' Procedure : Item
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Item(vNewValue As Integer) As Premio
    Dim mPrm As Premio

  On Error GoTo Item_Error
    '
    '   Comprobamos el indice de la colección
    '
    If vNewValue < 1 Or vNewValue > mCol.Count Then
        Set Item = Nothing
        Exit Function
    End If
    '
    '   Creamos el objeto
    '
    Set mPrm = New Premio
    Set mPrm = mCol.Item(vNewValue)
    '
    '   devolvemos el objeto
    '
    Set Item = mPrm
    Set mPrm = Nothing

   On Error GoTo 0
   Exit Function

Item_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Premios.Item")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function




'---------------------------------------------------------------------------------------
' Procedure : IndexOf
' Author    : CHARLY
' Date      : 06/04/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function IndexOf(vNewValue As Premio) As Integer
    Dim i As Integer
    Dim mPrem As Premio
    
  On Error GoTo IndexOf_Error
    '
    '   Inicializamos la función
    '
    IndexOf = -1
    '
    '   Comprobamos si nos pasan un objeto
    '
    If vNewValue Is Nothing Then
        Exit Function
    End If
    '
    '   Verificamos la categoria del premio
    '
    i = 1
    For Each mPrem In mCol
        If mPrem.CategoriaPremio = vNewValue.CategoriaPremio Then
            IndexOf = i
            Exit For
        Else
            i = i + 1
        End If
    Next mPrem
  
  On Error GoTo 0
   Exit Function

IndexOf_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Premios.IndexOf")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

' *===========(EOF): Premios.cls

