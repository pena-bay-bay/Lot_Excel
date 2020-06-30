VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TuplaAparicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public NumeroRegistro As Integer
Public FechaAparicion As Date
Public OrdenAparicion As Integer

Public Function ToString() As String
    Dim mStr As String
    mStr = "#{0} => {1} Orden: {2}"
    mStr = Replace(mStr, "{0}", CStr(NumeroRegistro))
    mStr = Replace(mStr, "{1}", Format(FechaAparicion, "dd/mm/yyyy"))
    mStr = Replace(mStr, "{2}", CStr(OrdenAparicion))
    ToString = mStr
End Function
