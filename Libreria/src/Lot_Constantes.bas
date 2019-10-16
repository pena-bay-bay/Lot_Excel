Attribute VB_Name = "Lot_Constantes"
'---------------------------------------------------------------------------------------
' Module    : Lot_Constantes
' Author    : CHARLY
' Date      : 25/05/2004 21:52
' Purpose   : M�dulo de definici�n de constantes lot
' Version   : 1.8.01
'---------------------------------------------------------------------------------------
'// Variables de versi�n de la librer�a
Public Const lotVersion = "2019.02"
Public Const lotFeVersion = "ju., 05/sep/2019 18:05:06"
'Public Const NOMBRE_APLICACION = "Bonoloto"

'// Variables relacionadas con el bombo
Public Const lotVacio = 1       '
Public Const lotLleno = 2
Public Const lotCargado = 3
Public Const lotTiempo = 1
Public Const lotGiros = 2
'
'   Mensajes de error
'
Public Const MSG_NOERROR = "No existen errores"
Public Const MSG_DESCONOCIDO = "Mensaje no registrado"
Public Const MSG_MALRANGO = "El rango del Numero debe estar comprendido entre 1 y 49, ambos inclusive"
Public Const MSG_FALTANumero = "Falta el n�mero, no se pueden realizar evaluaciones"
Public Const MSG_HAYERRORES = "Existen errores de inconsistencias:"
Public Const MSG_FECHAANALISCERO = "* La Fecha de Analisis no puede ser 0."
Public Const MSG_FECHAINICIALCERO = "* La Fecha Inicial no puede ser 0."
Public Const MSG_FECHAFINALCERO = "* La Fecha Final no puede ser 0."
Public Const MSG_NUMSORTEOSCERO = "* El n�mero de sorteos es 0."
Public Const MSG_FECHAANALISMENOR = "* La Fecha de Analisis es menor que la Fecha Final."
Public Const MSG_FECHAFINALMENOR = "* La Fecha Final es Menor que la Fecha Inicial."
Public Const MSG_FECHAANALISNOJUEGO = "* La Fecha de Analisis no pertenece al Juego."
Public Const MSG_FECHAFINALNOJUEGO = "* La Fecha Final no pertenece al Juego."
Public Const MSG_FECHAINICIALNOJUEGO = "* La Fecha Inicial no pertenece al Juego."

Public Const LT_ERROR = "#Err"
Public Const LT_PAR = "par"
Public Const LT_IMPAR = "impar"
Public Const LT_ALTO = "alto"
Public Const LT_BAJO = "bajo"

'*-----------------| Estados del Formulario |--------------------------+
Public Const ESTADO_INICIAL = 0
Public Const BOTON_CERRAR = 1
Public Const EJECUTAR = 5
Public Const COLOREAR_NumeroS = 2
Public Const COLOREAR_UNAFECHA = 3
Public Const COLOREAR_CARACTERISTICAS = 4
Public Const SIMULAR_METODOS = 6

'*-----------------| COLORES |------------------------------------------+
Public Const COLOR_ROJO = 3
Public Const COLOR_MARRON = 45
Public Const COLOR_AMARILLO = 6
Public Const COLOR_VERDE_CLARO = 35
Public Const COLOR_AZUL_CLARO = 8
Public Const COLOR_A�IL = 8
Public Const COLOR_AZUL_OSCURO = 41

Public Const COLOR_ERROR = 42           'Verde azulado
Public Const COLOR_NUMCOMPLE = 6        'Amarillo
Public Const COLOR_VERDE = 4            'Verde
Public Const COLOR_ANARANJADO = 45      'Anaranjado

'Colores de las Decenas
Public Const COLOR_DECENA1 = 36
Public Const COLOR_DECENA2 = 6
Public Const COLOR_DECENA3 = 40
Public Const COLOR_DECENA4 = 44
Public Const COLOR_DECENA5 = 46

'Colores de las terminaciones
Public Const COLOR_TERMINACION0 = 49
Public Const COLOR_TERMINACION1 = 23
Public Const COLOR_TERMINACION2 = 33
Public Const COLOR_TERMINACION3 = 4
Public Const COLOR_TERMINACION4 = 50
Public Const COLOR_TERMINACION5 = 10
Public Const COLOR_TERMINACION6 = 43
Public Const COLOR_TERMINACION7 = 44
Public Const COLOR_TERMINACION8 = 36
Public Const COLOR_TERMINACION9 = 6

'*----------------| Formatos
'Public Const FMT_IMPORTE = "_-* #,##0.0 _�_-;-* #,##0.0 _�_-;_-* ""-""?? _�_-;_-@_-"
Public Const FMT_IMPORTE = "* #,##0.00 �;-* #,##0.00 �"


'*-----------------| Barras de funciones |------------------------------+
Public Const BARRA_ESTUDIOS = "bar_studio"
Public Const BARRA_COLORES = "bar_colores"
Public Const BARRA_FUNCIONES = "bar_baybay"

'*----------------| Nombres de Referencia |-----------------------------+
Public THISLIBRO As String
Public Rango_Frecuencias    As Variant
Public Const MX_FECHA As Date = #1/1/2100#
Public Const PI As Double = 3.141592654
Public Enum TiposAciertos
    SinAciertos = 0
    SeisAciertosMasCyR = 1
    SeisAciertosMasR = 2
    SeisAciertosMasC = 3
    SeisAciertos = 4
    CincoAciertosMasC = 5
    CincoAciertosMasR = 6
    CincoAciertos = 7
    CuatroAciertosMasC = 8
    CuatroAciertos = 9
    TresAciertos = 10
    DosAciertosMasC = 11
    DosAciertos = 12
    UnAcierto = 13
End Enum
'
'   Categor�as de los premios
'
Public Enum CategoriaPremio
    Ninguna = 0
    Primera = 1
    Segunda = 2
    Tercera = 3
    Cuarta = 4
    Quinta = 5
    sexta = 6
    Septima = 7
    Octava = 8
    Novena = 9
    Decima = 10
    Undecima = 11
    duodecima = 12
    Trigesimotercera = 13
    Especial = 14
    Reintegro = 15
End Enum
'
'   Literal de las categorias
'
Public Const NOMBRE_CATEGORIASPREMIOS As String = "Primera;Segunda;Tercera;Cuarta;Quinta"
'
'   Tipos de juegos de Loterias y apuestas
'
'
Public Enum Juego
    SinJuego = 0
    LoteriaPrimitiva = 1
    bonoloto = 2
    gordoPrimitiva = 3
    Euromillones = 4
End Enum
'
'
'
Public Const JUEGO_DEFECTO As Integer = bonoloto
'
'   Literales de las constantes
'
Public Const NOMBRE_JUEGOS As String = "Loter�a Primitiva;Bonoloto;El Gordo de la Primitiva;Euro Millones"

'
'   Modalidades de juegos, en funci�n del n�mero de bolas
'
Public Enum ModalidadJuego
    LP_LB_6_49 = 1
    GP_5_54 = 2
    EU_5_50 = 3
    EU_2_12 = 4
End Enum
'
'
'
Public Const NOMBRE_MODALIDADES_JUEGO As String = "6 Bolas de 49;5 bolas de 54;5 bolas de 50"
'
'
'
Public Enum TipoOrdenacion
    ordSinDefinir = 0
    ordProbabilidad = 1
    ordProbTiempoMedio = 2
    ordFrecuencia = 3
    ordAusencia = 4
    ordTiempoMedio = 5
    ordDesviacion = 6
    ordProximaFecha = 7
    ordModa = 8
End Enum
'
'
'
Public Const NOMBRES_ORDENACION = "Sin Definir; Probabilidad; Prob.Tiempo Medio; Frecuencia; Ausencia; Tiempo Medio; Desviacion; Proxima fecha; Moda"
'
'
'
Public Enum TipoAgrupacion
    grpSinDefinir = 0
    grpDecenas = 1
    grpSeptenas = 2
    grpParidad = 3
    grpPeso = 4
    grpTerminacion = 5
End Enum
'
'
'
Public Const NOMBRES_PROCEDIMIENTOMETODO = "Sin Definir;Estadistico;Algoritmo Gen�tico;Red Neuronal;Estadisticas Combinaciones"
'
'
'
Public Enum ProcedimientoMetodo
    mtdSinDefinir = 0
    mtdEstadistico = 1
    mtdAlgoritmoAG = 2
    mtdRedNeuronal = 3
    mtdEstadCombinacion = 4
End Enum
'
'
'
Public Enum TipoAperturaFichero
    OpenForInput = 1
    OpenForAppend = 2
    OpenForOutput = 3
End Enum
'
'
'
Public Const NOMBRES_ROLES = "Sin Definir;Administrador;Escritor;Lector"
Public Enum Roles
    rolSinDefinir = 0
    rolAdministrador = 1
    rolEscritor = 2
    rolLector = 3
End Enum
'
'
'
Public Const NOMBRES_SITBOLETO = "Borrador;Validado;Premiado;Sorteado"
Public Enum SitBoleto
    blBorrador = 0
    blValidado = 1
    blPremiado = 2
    blSorteado = 3
End Enum
'
'
'
Public Const NOMBRES_AGRUPACION = "Sin Definir; Decenas; Septenas; Paridad; Peso; Terminacion "
'
'   Definici�n de premios
'                                          R   ;5; 4;   3;     2;     1
Public Const PREMIOS_BONOLOTO As String = "0,50;4;30;1000;100000;450000"
'                                          R;5�;4�;3�;2�;1�
Public Const PREMIOS_PRIMITIVA As String = "1;8;85;3000;75000;1000000"
'                                          13�;12�;11�;10�;9�;8�;7�;6�;5�;4�;3�;2�;1�
Public Const PREMIOS_EUROMILLON As String = "4;6;6;12;12;12;30;100;150;2000;60000;250000;1000000"
'                                       R;8�;7�;6�;5�;4�;3�;2�;1�
Public Const PREMIOS_GORDO As String = "1,5;3;8;20;50;250;3000;200000;1000000"
'
Public Const ERR_TODO = 999
Public Const MSG_TODO = "Rutina pendiente de codificar."
Public Const ERR_IDXINDIVIDUO = 10201
Public Const MSG_IDXINDIVIDUO = "Indice desbordado al obtener el individuo de una poblacion"


Public Const LT_EUROMILLON As String = "Euromillon"
Public Const LT_ESTRELLAS As String = "Estrellas"
Public Const LT_GORDO As String = "Gordo Primitiva"
Public Const LT_BONOLOTO As String = "Bonoloto"
Public Const LT_PRIMITIVA As String = "Primitiva"
Public Const LT_CLAVE As String = "Clave"
Public Const LT_COMPLEMENTARIO As String = "Complementario"
Public Const LT_REINTEGRO As String = "Reintegro"
Public Const LP_PREMIOS_EURO As String = "25.326.022,00;253.763,79;35.462,69;" & _
                                         "3.097,48;141,67;102,16;44,39;20,82;13" & _
                                         ",76;10,27;11,31;8,00;4,00"
Public Const LP_PREMIOS_GORDO As String = "5.438.778,89;165.842,81;7.911,24;" & _
                                         "105,37;28,33;9,17;4,89;3,00;1,50"
Public Const LP_PREMIOS_BONO As String = "981.440,37;47.921,93;895,74;37,09;4,00;0,50"
Public Const LP_PREMIOS_PRIMI As String = "13.085.952,17;1.598.135,28;59.930" & _
                                          ",07;1.438,76;51,58;8,00;1,00"
































