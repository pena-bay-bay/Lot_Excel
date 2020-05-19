VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Metodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Modulo    : Metodo
' Creado    : 16-dic-2006 21:28
' Autor     : Carlos Almela Baeza
' Version   : 1.0.0 Fecha
' Objeto    : Clase que representa el método de obtención de la sugerencia
'---------------------------------------------------------------------------------------
Option Explicit
'
Public EntidadNegocio       As EntidadNegocio
Public Parametros           As ParametrosMetodo
Public TipoProcedimiento    As ProcedimientoMetodo
Public EsMultiple           As Boolean
Public TipoMuestra          As Boolean              'False = dias, True = Registros


Private Sub Class_Initialize()
    Set EntidadNegocio = New EntidadNegocio
    Set Parametros = New ParametrosMetodo
    TipoProcedimiento = mtdEstadistico
    EsMultiple = False
    TipoMuestra = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TipoProcedimientoTostring
' Author    : Charly
' Date      : 10/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function TipoProcedimientoTostring() As String
    Dim m_vMatriz As Variant
    m_vMatriz = Split(NOMBRES_PROCEDIMIENTOMETODO, ";")
    TipoProcedimientoTostring = m_vMatriz(TipoProcedimiento)

End Function
