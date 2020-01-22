Attribute VB_Name = "Lot_CU_DefinirSorteo"
' *============================================================================*
' *
' *     Fichero    : Lot_CU_DefinirSorteo.bas
' *
' *     Autor      : Carlos Almela Baeza
' *     Creaci�n   : vi., 18/oct/2019 18:23:16
' *     Versi�n    : 1.0
' *     Prop�sito  : Suministrar las clases del CU Definir Sorteo a otros
' *                  libros que referencien a la libreria
' *============================================================================*
Option Explicit
Option Base 0
'------------------------------------------------------------------------------*
' Funcion        : GetSorteoModel
' Fecha          : vi., 18/oct/2019 18:25:23
' Prop�sito      : Suministrar clases de la libreria LotProject
' Retorno        : SorteoModel
'------------------------------------------------------------------------------*
'
Public Function GetSorteoModel() As SorteoModel
    Dim mObj As SorteoModel
    Set mObj = New SorteoModel
    Set GetSorteoModel = mObj
End Function
'------------------------------------------------------------------------------*
' Funcion        : GetPeriodo
' Fecha          :
' Prop�sito      : Suministrar clases de la libreria LotProject
' Retorno        : Periodo
'------------------------------------------------------------------------------*
'
Public Function GetPeriodo() As Periodo
    Dim mObj As Periodo
    Set mObj = New Periodo
    Set GetPeriodo = mObj
End Function


' *===========(EOF): Lot_CU_DefinirSorteo.bas
