VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParametrosMPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
'   Variables Publicas
'
Public BolasCombinacion As Integer          ' Bolas de la combinacion
Public NumerosContador As Integer           ' Numero que se utilizan para combinar
Public NumerosRegistrosCombinados As Long   ' Numeros de registros que se combinan
Public NumerosRegistrosOK As Long           ' Numeros de registros que se combinan
Public NumerosRegistrosKO As Long           ' Numeros de registros que se combinan
Public TiempoGeneracion As Date             ' Tiempo en realizar la generacion
Public TiempoFiltrado As Date               ' Tiempo en realizar la filtración
Public Filtros As Collection                ' Filtros a aplicar en el proceso
Public NombreFileCombinacion As String      ' Ruta y nombre del fichero de salida de combinaciones
Public NombreFileFiltroOK As String         ' Ruta y nombre del fichero de combinaciones filtradas
Public NombreFileFiltroKO As String         ' Ruta y nombre del fichero de combinaciones rechazadas
Public NombreFileEvaluacion As String       ' Ruta y nombre del fichero de combinaciones evaluadas
Public PeriodoControl As Periodo            ' Periodo de resultados para evaluar


Private Sub Class_Initialize()
    Set Filtros = New Collection
End Sub
