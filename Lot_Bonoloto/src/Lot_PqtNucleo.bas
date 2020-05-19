Attribute VB_Name = "Lot_PqtNucleo"
Option Explicit

Public Function GetColParametros() As ParametrosProceso
    Dim mObj As ParametrosProceso
    Set mObj = New ParametrosProceso
    Set GetColParametros = mObj
End Function

Public Function GetParamProceso() As ParamProceso
    Dim mObj As ParamProceso
    Set mObj = New ParamProceso
    Set GetParamProceso = mObj
End Function

Public Function GetPoblacion() As Poblacion
    Dim mObj As Poblacion
    Set mObj = New Poblacion
    Set GetPoblacion = mObj
End Function


Public Function GetIndividuo() As Individuo
    Dim mObj As Individuo
    Set mObj = New Individuo
    Set GetIndividuo = mObj
End Function

Public Function GetBombo() As BomboV2
    Dim mObj As BomboV2
    Set mObj = New BomboV2
    Set GetBombo = mObj
End Function

