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
' *     Creación   : do., 10/feb/2019 00:49:51
' *     Versión    : 1.0
' *     Propósito  : Colección de premios de un sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private mCol            As Collection
Private mPremio         As Premio
Private mMatMult        As Collection
Private mKey            As String
Private mKeyAciertos    As String

'--- Constantes ---------------------------------------------------------------*
'
'   Matrices de multiples premios según juego y apuestas
'
Private Const LT_MATRIZMULTIBL001 As Variant = "776cg+C1,1;776cg+C2,6;776cg1,1;" & _
    "776cg3,1;775cg+C2,1;775cg+C3,1;775cg+C4,5;775cg3,2;775cg4,2;774cg4,3;774c" & _
    "g5,4;773cg,4;8286cg+C1,1;8286cg+C2,6;8286cg+C3,6;8286cg+C4,15;8286cg1,1;8" & _
    "286cg3,12;8286cg4,15;8285cg+C2,1;8285cg+C3,2;8285cg+C4,15;8285cg+C5,20;82" & _
    "85cg3,3;8285cg4,15;8285cg5,10;8284cg4,6;8284cg3,16;8283cg5,10;5445cg1,1;5" & _
    "445cg2,1;5445cg3,42;5444cg+C2,2;5444cg+C4,42;5444cg3,2;5444cg4,42;5443cg4" & _
    ",3;5443cg5,41;5442cg,4;9846cg+C1,1;9846cg+C2,6;9846cg+C3,12;9846cg+C4,45;" & _
    "9846cg+C5,20;9846cg1,1;9846cg3,18;9846cg4,45;9846cg5,20;9845cg+C2,1;9845c" & _
    "g+C3,3;9845cg+C4,30;9845cg+C5,40;9845cg3,4;9845cg4,30;9845cg5,40;9844cg4," & _
    "10;9844cg5,40;9843cg5,20;102106cg+C1,1;102106cg+C2,6;102106cg+C3,18;10210" & _
    "6cg+C4,90;102106cg+C5,80;102106cg1,1;102106cg3,24;102106cg4,90;102106cg5," & _
    "80;102105cg+C2,1;102105cg+C3,4;102105cg+C4,50;102105cg+C5,100;102105cg3,5" & _
    ";102105cg4,50;102105cg5,100;102104cg4,15;102104cg5,80;102103cg5,35;114626" & _
    "cg+C1,1;114626cg+C2,6;114626cg+C3,24;114626cg+C4,150;114626cg+C5,200;1146" & _
    "26cg1,1;114626cg3,30;114626cg4,150;114626cg5,200;114625cg+C2,1;114625cg+C" & _
    "3,5;114625cg+C4,75;114625cg+C5,200;114625cg3,6;114625cg4,75;114625cg5,200" & _
    ";114624cg4,21;114624cg5,140;114623cg5,56"

Private Const LT_MATRIZMULTILP001 As Variant = "776cg+R+C14,1;776cg+R+C1,1;776c" & _
    "g+R+C2,6;776cg+R14,1;776cg+R1,1;776cg+R3,6;776cg+C1,1;776cg+C2,6;776cg1,1;" & _
    "776cg3,6;775cg+C2,1;775cg+C3,1;775cg+C4,5;775cg3,2;775cg4,5;774cg4,3;774cg" & _
    "5,4;773cg5,4;8286cg+R+C14,1;8286cg+R+C1,1;8286cg+R+C2,6;8286cg+R+C3,6;8286" & _
    "cg+R+C4,15;8286cg+R14,1;8286cg+R1,1;8286cg+R3,12;8286cg+R4,15;8286cg+C1,1;" & _
    "8286cg+C2,6;8286cg+C3,6;8286cg+C4,15;8286cg1,1;8286cg3,12;8286cg4,15;8285c" & _
    "g+C2,1;8285cg+C3,2;8285cg+C4,15;8285cg+C5,10;8285cg3,3;8285cg4,15;8285cg5," & _
    "10;8284cg4,6;8284cg5,16;8283cg5,10;5445cg+R14,1;5445cg+R1,1;5445cg+R2,1;54" & _
    "45cg+R3,42;5445cg1,1;5445cg2,1;5445cg3,42;5444cg+C2,2;5444cg+C4,42;5444cg3" & _
    ",2;5444cg4,42;5443cg4,3;5443cg5,41;5442cg5,4;9846cg+R+C14,1;9846cg+R+C1,1;" & _
    "9846cg+R+C2,6;9846cg+R+C3,12;9846cg+R+C4,45;9846cg+R+C5,20;9846cg+R14,1;98"

Private Const LT_MATRIZMULTILP002 As Variant = "46cg+R1,1;9846cg+R3,18;9846cg+R" & _
    "4,45;9846cg+R5,20;9846cg+C1,1;9846cg+C2,6;9846cg+C3,12;9846cg+C4,45;9846cg" & _
    "+C5,20;9846cg1,1;9846cg3,18;9846cg4,45;9846" & _
    "cg5,20;9845cg+C2,1;9845cg+C3,3;9845cg+C4,30;9845cg+C5,40;9845cg3,4;9845cg4" & _
    ",30;9845cg5,40;9844cg4,10;9844cg5,40;9843cg5,20;102106cg+R+C14,1;102106cg+" & _
    "R+C1,1;102106cg+R+C2,6;102106cg+R+C3,18;102106cg+R+C4,90;102106cg+R+C5,80;" & _
    "102106cg+R14,1;102106cg+R1,1;102106cg+R3,24;102106cg+R4,90;102106cg+R5,80;" & _
    "102106cg+C1,1;102106cg+C2,6;102106cg+C3,18;102106cg+C4,90;102106cg+C5,80;1" & _
    "02106cg1,1;102106cg3,24;102106cg4,90;102106cg5,80;102105cg+C2,1;102105cg+C" & _
    "3,4;102105cg+C4,50;102105cg+C5,100;102105cg3,5;102105cg4,50;102105cg5,100;" & _
    "102104cg4,15;102104cg5,80;102103cg5,35;114626cg+R+C14,1;114626cg+R+C1,1;11" & _
    "4626cg+R+C2,6;114626cg+R+C3,24;114626cg+R+C4,150;114626cg+R+C5,200;114626c" & _
    "g+R14,1;114626cg+R1,1;114626cg+R3,30;114626cg+R4,150;114626cg+R5,200;11462" & _
    "6cg+C1,1;114626cg+C2,6;114626cg+C3,24;114626cg+C4,150;114626cg+C5,200;1146" & _
    "26cg1,1;114626cg3,30;114626cg4,150;114626cg5,200;114625cg+C2,1;114625cg+C3" & _
    ",5;114625cg+C4,75;114625cg+C5,200;114625cg3,6;114625cg4,75;114625cg5,200;1" & _
    "14624cg4,21;114624cg5,140;114623cg5,56"
    
Private Const LT_MATRIZMULTIGP001 As Variant = "665+11,1;665+13,5;6652,1;6654,5" & _
    ";664+13,2;664+15,4;6644,2;6646,4;663+15,3;663+17,3;6636,3;6638,3;662+17,4;" & _
    "6628,4;7215+11,1;7215+13,10;7215+15,10;72152,1;72154,10;72156,10;7214+13,3" & _
    ";7214+15,12;7214+17,6;72144,3;72146,12;72148,6;7213+15,6;7213+17,12;72136," & _
    "6;72138,12;7212+17,10;72128,10;8565+11,1;8565+13,15;8565+15,30;8565+17,10;" & _
    "85652,1;85654,15;85656,30;85658,10;8564+13,4;8564+15,24;8564+17,24;85644,4" & _
    ";85646,24;85648,24;8563+15,10;8563+17,30;85636,10;85638,30;8562+17,20;8562"

Private Const LT_MATRIZMULTIGP002 As Variant = "8,20;91265+11,1;91265+13,20;912" & _
    "65+15,60;91265+17,40;912652,1;912654,20;912656,60;912658,40;91264+13,5;912" & _
    "64+15,40;91264+17,60;912644,5;912646,40;912648,60;91263+15,15;91263+17,60;" & _
    "912636,15;912638,60;91262+17,35;912628,35;102525+11,1;102525+13,25;102525+" & _
    "15,100;102525+17,100;1025252,1;1025254,25;1025256,100;1025258,100;102524+1" & _
    "3,6;102524+15,60;102524+17,120;1025244,6;1025246,60;1025248,120;102523+15," & _
    "21;102523+17,105;1025236,21;1025238,105;102522+17,56;1025228,56;114625+11,"
    
Private Const LT_MATRIZMULTIGP003 As Variant = "1;114625+13,30;114625+15,150;11" & _
    "4625+17,200;1146252,1;1146254,30;1146256,150;1146258,200;114624+13,7;11462" & _
    "4+15,84;114624+17,210;1146244,7;1146246,84;1146248,210;114623+15,28;114623" & _
    "+17,168;1146236,28;1146238,168;114622+17,84;1146228,84"

Private Const LT_MATRIZMULTIEU001 As Variant = ""
Private Const LT_MATRIZMULTIEU002 As Variant = ""
Private Const LT_MATRIZMULTIEU003 As Variant = ""
'
'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*
Public IdSorteo                   As Integer
'------------------------------------------------------------------------------*
' Procedimiento  : Juego
' Fecha          : do., 10/feb/2019 09:48:57
' Propósito      : Juego al que pertenece el premio
'------------------------------------------------------------------------------*
Public Juego                      As Juego
'------------------------------------------------------------------------------*
' Procedimiento  : ComplementarioAcertado
' Fecha          : do., 10/feb/2019 09:48:57
' Propósito      : Indicador de que se ha acertado el complementario
'------------------------------------------------------------------------------*
Public ComplementarioAcertado     As Boolean
'------------------------------------------------------------------------------*
' Procedimiento  : EstrellasAcertadas
' Fecha          : do., 10/feb/2019 09:48:57
' Propósito      : Numero de estrellas acertadas {0,1,2}
'------------------------------------------------------------------------------*
Public EstrellasAcertadas         As Integer
'------------------------------------------------------------------------------*
' Procedimiento  : NumerosAcertados
' Fecha          : do., 10/feb/2019 09:48:57
' Propósito      : Numero de pronosticos acertados, incluido el complementario
'------------------------------------------------------------------------------*
Public NumerosAcertados           As Integer
'------------------------------------------------------------------------------*
' Procedimiento  : NumerosAcertados
' Fecha          : do., 10/feb/2019 09:48:57
' Propósito      : Indicador de reintegro o clave, según juego, acertado
'------------------------------------------------------------------------------*
Public ReintegroAcertado          As Boolean
'------------------------------------------------------------------------------*
' Procedimiento  : PronosticosApostados
' Fecha          : do., 10/feb/2019 09:48:57
' Propósito      : Numero de pronosticos apostados en el boleto {5.. 11}
'------------------------------------------------------------------------------*
Public PronosticosApostados       As Integer
'------------------------------------------------------------------------------*
' Procedimiento  : EstrellasApostadas
' Fecha          : do., 10/feb/2019 09:48:57
' Propósito      : Numero de estrellas apostadas en el boleto {2..5}
'------------------------------------------------------------------------------*
Public EstrellasApostadas         As Integer
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
'------------------------------------------------------------------------------*
' Procedimiento  : Count
' Fecha          : do., 10/feb/2019 09:54:46
' Propósito      : Devuelve el número de premios de la colección
' Retorno        : Numero de elementos
'------------------------------------------------------------------------------*
Public Property Get Count() As Integer

   On Error GoTo Count_Error

    Count = mCol.Count

   On Error GoTo 0
   Exit Property

Count_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Count")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Property

'--- Metodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    ComplementarioAcertado = False
    ReintegroAcertado = False
    EstrellasAcertadas = 0
    NumerosAcertados = 0
    PronosticosApostados = 0
    EstrellasApostadas = 0
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : GetPremioObtenido
' Fecha          : sá., 16/feb/2019 21:32:37
' Propósito      : Con los numeros y estrellas acertadas se obtiene la categoria
'                  del premio
'------------------------------------------------------------------------------*
Private Function GetPremioObtenido() As CategoriaPremio
 On Error GoTo GetPremioObtenido_Error
    '
    '   Inicializamos la categoria del premio
    '
    GetPremioObtenido = Ninguna
    mKeyAciertos = ""
    '
    '   Evaluamos la categoria de premio obtenido segun el juego
    '   los numeros acertados y estrellas acertadas
    '
    Select Case Juego
        Case Bonoloto:
            If PronosticosApostados = 5 Then
                Select Case True
                    Case (NumerosAcertados = 5)
                        GetPremioObtenido = Primera
                        mKeyAciertos = "5cg"
                    Case (NumerosAcertados = 5 And ComplementarioAcertado)
                        GetPremioObtenido = Segunda
                        mKeyAciertos = "4cg+C"
                    Case (NumerosAcertados = 5 And Not ComplementarioAcertado)
                        GetPremioObtenido = Tercera
                        mKeyAciertos = "4cg"
                    Case (NumerosAcertados = 4 And Not ComplementarioAcertado)
                        GetPremioObtenido = Tercera
                        mKeyAciertos = "4cg"
                    Case (NumerosAcertados = 4 And ComplementarioAcertado)
                        GetPremioObtenido = Cuarta
                        mKeyAciertos = "3cg"
                    Case (NumerosAcertados = 3 And Not ComplementarioAcertado)
                        GetPremioObtenido = Quinta
                        mKeyAciertos = "2cg"
                    Case (NumerosAcertados = 2 And Not ComplementarioAcertado)
                        GetPremioObtenido = Quinta
                        mKeyAciertos = "2cg"
                    Case ReintegroAcertado: GetPremioObtenido = Reintegro
                End Select
            Else
                Select Case True
                    Case (NumerosAcertados = 7 And ComplementarioAcertado)
                        GetPremioObtenido = Primera
                        mKeyAciertos = "6cg+C"
                    Case (NumerosAcertados = 6 And Not ComplementarioAcertado)
                        GetPremioObtenido = Primera
                        mKeyAciertos = "6cg"
                    Case (NumerosAcertados = 6 And ComplementarioAcertado)
                        GetPremioObtenido = Segunda
                        mKeyAciertos = "5cg+C"
                    Case (NumerosAcertados = 5 And Not ComplementarioAcertado)
                        GetPremioObtenido = Tercera
                        mKeyAciertos = "5cg"
                    Case (NumerosAcertados = 5 And ComplementarioAcertado)
                        GetPremioObtenido = Cuarta
                        mKeyAciertos = "4cg"
                    Case (NumerosAcertados = 4 And Not ComplementarioAcertado)
                        GetPremioObtenido = Cuarta
                        mKeyAciertos = "4cg"
                    Case (NumerosAcertados = 4 And ComplementarioAcertado)
                        GetPremioObtenido = Quinta
                        mKeyAciertos = "3cg"
                    Case (NumerosAcertados = 3 And Not ComplementarioAcertado):
                        GetPremioObtenido = Quinta
                        mKeyAciertos = "3cg"
                    Case ReintegroAcertado: GetPremioObtenido = Reintegro
                End Select
            End If
        Case LoteriaPrimitiva:
            Select Case True
                Case (NumerosAcertados = 6 And ReintegroAcertado And Not ComplementarioAcertado): GetPremioObtenido = Especial
                Case (NumerosAcertados = 6 And Not ComplementarioAcertado): GetPremioObtenido = Primera
                Case (NumerosAcertados = 6 And ComplementarioAcertado): GetPremioObtenido = Segunda
                Case (NumerosAcertados = 5 And ComplementarioAcertado): GetPremioObtenido = Segunda
                Case (NumerosAcertados = 5 And Not ComplementarioAcertado): GetPremioObtenido = Tercera
                Case (NumerosAcertados = 5 And ComplementarioAcertado): GetPremioObtenido = Cuarta
                Case (NumerosAcertados = 4 And Not ComplementarioAcertado): GetPremioObtenido = Cuarta
                Case (NumerosAcertados = 4 And ComplementarioAcertado): GetPremioObtenido = Quinta
                Case (NumerosAcertados = 3 And Not ComplementarioAcertado): GetPremioObtenido = Quinta
                Case ReintegroAcertado: GetPremioObtenido = Reintegro
            End Select
        '
        '   Euromillones
        '
        Case Euromillones
            Select Case True
                Case (NumerosAcertados = 5 And EstrellasAcertadas = 2): GetPremioObtenido = Primera
                Case (NumerosAcertados = 5 And EstrellasAcertadas = 1): GetPremioObtenido = Segunda
                Case (NumerosAcertados = 5 And EstrellasAcertadas = 0): GetPremioObtenido = Tercera
                Case (NumerosAcertados = 4 And EstrellasAcertadas = 2): GetPremioObtenido = Cuarta
                Case (NumerosAcertados = 4 And EstrellasAcertadas = 1): GetPremioObtenido = Quinta
                Case (NumerosAcertados = 3 And EstrellasAcertadas = 2): GetPremioObtenido = sexta
                Case (NumerosAcertados = 4 And EstrellasAcertadas = 0): GetPremioObtenido = Septima
                Case (NumerosAcertados = 2 And EstrellasAcertadas = 2): GetPremioObtenido = Octava
                Case (NumerosAcertados = 3 And EstrellasAcertadas = 1): GetPremioObtenido = Novena
                Case (NumerosAcertados = 3 And EstrellasAcertadas = 0): GetPremioObtenido = Decima
                Case (NumerosAcertados = 1 And EstrellasAcertadas = 2): GetPremioObtenido = Undecima
                Case (NumerosAcertados = 2 And EstrellasAcertadas = 1): GetPremioObtenido = duodecima
                Case (NumerosAcertados = 2 And EstrellasAcertadas = 0): GetPremioObtenido = Trigesimotercera
            End Select
        '
        '   Gordo Prmitiva
        '
        Case gordoPrimitiva
            Select Case True
                Case (NumerosAcertados = 5 And ReintegroAcertado): GetPremioObtenido = Primera
                Case (NumerosAcertados = 5 And Not ReintegroAcertado): GetPremioObtenido = Segunda
                Case (NumerosAcertados = 4 And ReintegroAcertado): GetPremioObtenido = Tercera
                Case (NumerosAcertados = 4 And Not ReintegroAcertado): GetPremioObtenido = Cuarta
                Case (NumerosAcertados = 3 And ReintegroAcertado): GetPremioObtenido = Quinta
                Case (NumerosAcertados = 3 And Not ReintegroAcertado): GetPremioObtenido = sexta
                Case (NumerosAcertados = 2 And ReintegroAcertado): GetPremioObtenido = Septima
                Case (NumerosAcertados = 2 And Not ReintegroAcertado): GetPremioObtenido = Octava
                Case ReintegroAcertado: GetPremioObtenido = Reintegro
            End Select
            
    End Select
    
 On Error GoTo 0
   Exit Function

GetPremioObtenido_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.GetPremioObtenido")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : GetNumApuestas
' Fecha          : vi., 22/feb/2019 23:06:58
' Propósito      : Determina si con el numero de bolas acertadas hay mas de
'                  una apuesta
'------------------------------------------------------------------------------*
Private Function GetNumApuestas() As Integer
 On Error GoTo GetNumApuestas_Error
    '
    '   Inicializamos el numero de apuestas a 0
    '
    GetNumApuestas = 0
    '
    '   Evaluamos el número de apuestas segun las bolas acertadas y el juego
    '   EstrellasApostadas,PronosticosApostados
    Select Case Juego
        Case Bonoloto, LoteriaPrimitiva:
            Select Case True
                Case PronosticosApostados = 6: GetNumApuestas = 1
                Case PronosticosApostados = 7: GetNumApuestas = 7
                Case PronosticosApostados = 8: GetNumApuestas = 28
                Case PronosticosApostados = 5: GetNumApuestas = 44
                Case PronosticosApostados = 9: GetNumApuestas = 84
                Case PronosticosApostados = 10: GetNumApuestas = 210
                Case PronosticosApostados = 11: GetNumApuestas = 462
            End Select
        '
        '   Euromillones
        '
        Case Euromillones
            Select Case True
                Case PronosticosApostados = 5 And EstrellasApostadas = 2: GetNumApuestas = 1
                Case PronosticosApostados = 6 And EstrellasApostadas = 2: GetNumApuestas = 6
                Case PronosticosApostados = 7 And EstrellasApostadas = 2: GetNumApuestas = 21
                Case PronosticosApostados = 8 And EstrellasApostadas = 2: GetNumApuestas = 56
                Case PronosticosApostados = 9 And EstrellasApostadas = 2: GetNumApuestas = 126
                Case PronosticosApostados = 10 And EstrellasApostadas = 2: GetNumApuestas = 252
                Case PronosticosApostados = 5 And EstrellasApostadas = 3: GetNumApuestas = 3
                Case PronosticosApostados = 6 And EstrellasApostadas = 3: GetNumApuestas = 18
                Case PronosticosApostados = 7 And EstrellasApostadas = 3: GetNumApuestas = 63
                Case PronosticosApostados = 8 And EstrellasApostadas = 3: GetNumApuestas = 168
                Case PronosticosApostados = 9 And EstrellasApostadas = 3: GetNumApuestas = 378
                Case PronosticosApostados = 10 And EstrellasApostadas = 3: GetNumApuestas = 756
                Case PronosticosApostados = 5 And EstrellasApostadas = 4: GetNumApuestas = 6
                Case PronosticosApostados = 6 And EstrellasApostadas = 4: GetNumApuestas = 36
                Case PronosticosApostados = 7 And EstrellasApostadas = 4: GetNumApuestas = 126
                Case PronosticosApostados = 8 And EstrellasApostadas = 4: GetNumApuestas = 336
                Case PronosticosApostados = 9 And EstrellasApostadas = 4: GetNumApuestas = 756
                Case PronosticosApostados = 10 And EstrellasApostadas = 4: GetNumApuestas = 1512
                Case PronosticosApostados = 5 And EstrellasApostadas = 5: GetNumApuestas = 10
                Case PronosticosApostados = 6 And EstrellasApostadas = 5:  GetNumApuestas = 60
                Case PronosticosApostados = 7 And EstrellasApostadas = 5:  GetNumApuestas = 210
                Case PronosticosApostados = 8 And EstrellasApostadas = 5:  GetNumApuestas = 560
                Case PronosticosApostados = 9 And EstrellasApostadas = 5:  GetNumApuestas = 1260
                Case PronosticosApostados = 10 And EstrellasApostadas = 5:  GetNumApuestas = 2520
            End Select
        '
        '   Gordo Prmitiva
        '
        Case gordoPrimitiva
            Select Case True
                Case PronosticosApostados = 5: GetNumApuestas = 1
                Case PronosticosApostados = 6: GetNumApuestas = 6
                Case PronosticosApostados = 7: GetNumApuestas = 21
                Case PronosticosApostados = 8: GetNumApuestas = 56
                Case PronosticosApostados = 9: GetNumApuestas = 126
                Case PronosticosApostados = 10: GetNumApuestas = 252
                Case PronosticosApostados = 11: GetNumApuestas = 462
            End Select
    End Select
    
 On Error GoTo 0
   Exit Function

GetNumApuestas_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.GetNumApuestas")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetMatrizApuestas
' Fecha          : do., 24/feb/2019 21:22:32
' Propósito      : Construye una colección con la matriz de apuestas multiples
'                  de cada juego
'------------------------------------------------------------------------------*
Private Function GetMatrizApuestas() As Collection
    Dim mMatriz     As Variant
    Dim mElem       As Variant
    Dim i           As Integer
    Dim mTmpKey     As String
    Dim mVal        As Integer
    Dim mDupla      As Dupla
 On Error GoTo GetMatrizApuestas_Error
    '
    '   Creamos la colección de elementos de matriz
    '
    Set mMatMult = New Collection
    '
    '   Seleccionamos la matriz según el juego
    '
    Select Case Juego
        Case Bonoloto:
            mMatriz = Split(LT_MATRIZMULTIBL001, ";")
        
        Case LoteriaPrimitiva:
            mMatriz = Split(LT_MATRIZMULTILP001 & LT_MATRIZMULTILP002, ";")
        
        Case gordoPrimitiva:
            mMatriz = Split(LT_MATRIZMULTIGP001 & LT_MATRIZMULTIGP002 & LT_MATRIZMULTIGP003, ";")
        
        Case Euromillones:
            mMatriz = Split(LT_MATRIZMULTIEU001 & LT_MATRIZMULTIEU002 & LT_MATRIZMULTIEU003, ";")
            
    End Select
    '
    '   Para cada elemento en la matriz
    '
   
    For i = 0 To UBound(mMatriz)
        '
        '   Separamos el elemento de la matriz en
        '   una dupla: clave y apuestas
        '
        mElem = Split(mMatriz(i), ",")
        '
        '   Creamos la dupla
        '
        Set mDupla = New Dupla
        '
        '   Asignamos la clave y el valor
        '
        mDupla.Key = mElem(0)
        mDupla.Value = mElem(1)
        '
        '   Si la matriz a tiene elementos
        '
        If mMatMult.Count <> 0 Then
            '
            '   Agregamos al final
            '
             mMatMult.Add mDupla, mDupla.Key, , mMatMult.Count
        Else
            '
            '   Agregamos a la colección
            '
             mMatMult.Add mDupla, mDupla.Key
        End If
    '
    '   Siguiente elemento de la matriza
    '
    Next i
    '
    '   Devolvemos la matriz
    '
    Set GetMatrizApuestas = mMatMult
 On Error GoTo 0
   Exit Function

GetMatrizApuestas_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.GetMatrizApuestas")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function
'--- Metodos Publicos ---------------------------------------------------------*

'------------------------------------------------------------------------------*
' Procedimiento  : Clear
' Fecha          : do., 10/feb/2019 09:58:21
' Propósito      : Elminar todos los elementos de la colección
'------------------------------------------------------------------------------*
Public Sub Clear()
   On Error GoTo Clear_Error
    Set mCol = Nothing
    Set mCol = New Collection
    
   On Error GoTo 0
   Exit Sub

Clear_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Clear")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetImporteTotalPremios
' Fecha          : do., 10/feb/2019 10:03:46
' Propósito      : Obtiene el importe total de los premios
' Parámetros     :
'------------------------------------------------------------------------------*
'
Public Function GetImporteTotalPremios() As Currency
    Dim mCatPremio      As CategoriaPremio
    Dim mNumApuestas    As Integer
    Dim mMtzApuestas    As Variant
    Dim mDupla          As Dupla
    Dim mApuestas       As Integer
    Dim mStrKey         As String
    
 On Error GoTo GetImporteTotalPremios_Error
    '
    '  inicializamos el importe
    '
    GetImporteTotalPremios = 0
    '
    '   Si el juego no está establecido sale
    '
    If Juego = SinJuego Then
        Exit Function
    End If
    '
    '   Evaluamos el premio obtenido
    '
    mCatPremio = GetPremioObtenido
    '
    '   Si es un boleto múltiple obtenemos matriz de apuestas
    '
    mNumApuestas = GetNumApuestas
    '
    '   Si es una apuesta múltiple se obtiene la matriz de apuestas
    '
    If mNumApuestas > 1 Then
        Set mMtzApuestas = GetMatrizApuestas
    End If
    '
    '   Si el numero de apuestas es 0 no se calcula el premio
    '
    If mNumApuestas = 0 Then
        Exit Function
    End If
    '
    '   Creamos el objeto Dupla
    '
    Set mDupla = New Dupla
    '
    '  calculamos el importe para todos los premios igualo inferiores a la categoria
    '  correspondiente
    '
    For Each mPremio In mCol
        '
        '   Si la apuesta es multiple obtenemos las apuestas correspondientes a la categoria
        '   analizada
        '
        If mNumApuestas > 1 Then
            mStrKey = PronosticosApostados & mNumApuestas & mKeyAciertos & mPremio.Categoria
            '
            '   controlamos el error de clave no existe en la colección
            '
            On Error Resume Next
            Set mDupla = mMtzApuestas(mStrKey)
            If Err.Number = 5 Then
                mApuestas = 0
            Else
                mApuestas = Val(mDupla.Value)
            End If
        Else
            mApuestas = 1
        End If
        
        '
        '  Si es multiple debemmos ver si corresponde la categoria
        '  si es simple
        If mPremio.Categoria >= mCatPremio Then
            GetImporteTotalPremios = GetImporteTotalPremios + mPremio.GetImportePremio(mApuestas)
        End If
    Next mPremio

    
 On Error GoTo 0
GetImporteTotalPremios__CleanExit:
    Exit Function
            
GetImporteTotalPremios_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Premios.GetImporteTotalPremios", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Premios.GetImporteTotalPremios", ErrDescription
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : SetPremiosDefecto
' Fecha          : do., 10/feb/2019 10:07:36
' Propósito      : Establece una colección de premios por defecto
' Parámetros     :
'------------------------------------------------------------------------------*
'
Public Sub SetPremiosDefecto(datJuego As Juego)
    Dim mStr    As String
    Dim mVar    As Variant
    Dim mVar2   As Variant
    Dim i       As Integer
 On Error GoTo SetPremiosDefecto_Error
    '
    '   Inicializamos la colección
    '
    Me.Clear
    '
    '   Establecemos el juego de los premios
    '
    Me.Juego = datJuego
    '
    '   Determinamos datos por defecto de la matriz
    '
    Select Case datJuego
        Case Bonoloto: mStr = LT_PREMIOSBONOLOTO
        Case LoteriaPrimitiva: mStr = LT_PREMIOSPRIMITIVA
        Case Euromillones: mStr = LT_PREMIOSEUROMILLONES
        Case gordoPrimitiva: mStr = LT_PREMIOSGORDO
    End Select
    '
    '   Creamos una mariz con las categorias y premio por defecto
    '
    mVar = Split(mStr, ";")
    
    For i = 0 To UBound(mVar)
        '
        '   Creamos un premio
        '
        Set mPremio = New Premio
        '
        '   Separamos la categoria del importe
        '
        mVar2 = Split(mVar(i), ",")
        '
        '   Asignamos valores a premio
        '
        With mPremio
            .Categoria = Val(mVar2(0))
            .Importe = Val(mVar2(1))
            .ImporteDefault = True
            .Juego = datJuego
        End With
        '
        '   Agregamos a la colección
        '
        Me.Add mPremio
        '
        '   Eliminamos Premio
        '
        Set mPremio = Nothing
    Next i
    '
    '
    '
 On Error GoTo 0
SetPremiosDefecto__CleanExit:
    Exit Sub
            
SetPremiosDefecto_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Premios.SetPremiosDefecto", ErrSource)
    '   Lanza el Error
    Err.Raise ErrNumber, "Premios.SetPremiosDefecto", ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Add
' Fecha          : 13/01/2018
' Propósito      : Agrega un parámetro a la colección
' Parámetros     : datParametro
'------------------------------------------------------------------------------*
Public Sub Add(datPremio As Premio)
    Dim strKey  As String
  On Error GoTo Add_Error
    '
    '   Calculamos la clave
    '
    strKey = CStr(datPremio.Categoria)
    '
    '   Si ya tenemos elementos agregamos al final
    '
    If mCol.Count <> 0 Then
        mCol.Add datPremio, strKey, , mCol.Count
    Else
        mCol.Add datPremio, strKey
    End If

  On Error GoTo 0
   Exit Sub

Add_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Add")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Delete
' Fecha          : do., 10/feb/2019 10:01:35
' Propósito      : Elimina un premio de la colección por su índice
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub Delete(datIndex As Integer)
    Dim strError As String
    
  On Error GoTo Delete_Error
    If datIndex > mCol.Count Then
        strError = Replace(MSG_DELETEINDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_DELETEINDEXERROR, "Premios.Delete", strError
    Else
        mCol.Remove datIndex
    End If
    
  On Error GoTo 0
    Exit Sub

Delete_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    '   Audita el error
    Call HandleException(ErrNumber, ErrDescription, "Premios.Delete")
    '   Lanza el Error
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : MarkForDelete
' Fecha          : do., 10/feb/2019 10:01:35
' Propósito      : Marcar un elemento para borrar de la BBDD
' Parámetros     : Indice del elemento a borrar
' Retorno        :
'------------------------------------------------------------------------------*
Public Sub MarkForDelete(datIndex As Integer)
    Dim strError As String
    
   On Error GoTo MarkForDelete_Error
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Premios.MarkForDelete", strError
    Else
        Set mPremio = mCol.Item(datIndex)
        mPremio.MarkForDelete = True
    End If
    
   On Error GoTo 0
   Exit Sub

MarkForDelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.MarkForDelete")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

'------------------------------------------------------------------------------*
' Procedimiento  : Undelete
' Fecha          : 13/01/2018
' Propósito      : Desmarcar un elemento para borrar
' Parámetros     : Indice de la colección
'------------------------------------------------------------------------------*
Public Sub Undelete(datIndex As Integer)
    Dim strError As String
    Dim mPar As Parametro
   On Error GoTo Undelete_Error
    
    If datIndex > mCol.Count Then
        strError = Replace(MSG_INDEXERROR, "#", "#" & datIndex)
        Err.Raise ERR_INDEXERROR, "Premios.Undelete", strError
    Else
        Set mPar = mCol.Item(datIndex)
        mPar.EntidadNegocio.MarkForDelete = False
    End If
   
   On Error GoTo 0
   Exit Sub

Undelete_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "Premios.Undelete")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : ToString
' Fecha          : vi., 22/nov/2019 19:54:21
' Propósito      : Establecer una cadena de caracteres con los premios
'------------------------------------------------------------------------------*
Public Function ToString() As String
    Dim mStr    As String
  On Error GoTo ToString_Error
    '
    '   Inicializamos resultado
    '
    mStr = ""
    If mCol.Count > 0 Then
        For Each mPremio In mCol
            mStr = mStr & CStr(mPremio.Categoria) & "," & CStr(mPremio.Importe) & ";"
        Next mPremio
        '
        '   Eliminamos el último ";"
        '
        ToString = Left(mStr, Len(mStr) - 1)
    Else
        ToString = "#Sin Definir"
    End If
  On Error GoTo 0
   Exit Function

ToString_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   Call HandleException(ErrNumber, ErrDescription, "Premios.ToString")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : Parse
' Fecha          : vi., 02/dic/2019 19:54:21
' Propósito      : Establece los premios desde una cadena de caracteres
'------------------------------------------------------------------------------*
Public Sub Parse(datStr As String, datJuego As Juego)
    Dim mStr    As String
    Dim mVar    As Variant
    Dim mVar2   As Variant
    Dim i       As Integer
 On Error GoTo Parse_Error
    '
    '   Inicializamos la colección
    '
    Me.Clear
    '
    '   Establecemos el juego de los premios
    '
    Me.Juego = datJuego
    '
    '   Creamos una mariz con las categorias y premio por defecto
    '
    mVar = Split(datStr, ";")
    
    For i = 0 To UBound(mVar)
        '
        '   Creamos un premio
        '
        Set mPremio = New Premio
        '
        '   Separamos la categoria del importe
        '
        mVar2 = Split(mVar(i), ",")
        '
        '   Asignamos valores a premio
        '
        With mPremio
            .Categoria = Val(mVar2(0))
            .Importe = Val(mVar2(1))
            .ImporteDefault = False
            .Juego = datJuego
        End With
        '
        '   Agregamos a la colección
        '
        Me.Add mPremio
        '
        '   Eliminamos Premio
        '
        Set mPremio = Nothing
    Next i
    '
    '
    '
 On Error GoTo 0
Parse__CleanExit:
    Exit Sub
            
Parse_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Premios.Parse", ErrSource)
    Err.Raise ErrNumber, "Premios.SetPremiosDefecto", ErrDescription
End Sub
'' *===========(EOF): Premios.cls

