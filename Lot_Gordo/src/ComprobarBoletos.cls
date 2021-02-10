VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComprobarBoletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : ComprobarBoletos.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ju., 16/jul/2020 19:49:41
' *     Modificado :
' *     Versión    : 1.0
' *     Propósito  : Funcionalidad de comprobar un boleto
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'--- Variables Privadas -------------------------------------------------------*
Private oSortEng                As SorteoEngine
Private oSorteo                 As Sorteo
Private oPremio                 As Premio
Private nBolasAcertadas         As Integer
Private nEstrellasAcertadas     As Integer
Private bComplementarioAcertado As Boolean
Private bReintegroAcertado      As Boolean
Private bElMillonAcertado       As Boolean
Private bJokerAcertado          As Boolean
Private mImporteApta            As Currency
Private mImporteBlto            As Currency
Private mCategoriaApuesta       As CategoriaPremio
Private mMatMult                As Collection

'--- Constantes ---------------------------------------------------------------*
Private Const LT_ACIERTOSBL As String = "6cg+R+C;6cg+R;6cg+C;6cg;5cg+C;5cg;" & _
                                "4cg;3cg;2cg;1cg;0cg"
Private Const LT_ACIERTOSLP As String = "6cg+R+C;6cg+R;6cg+C;6cg;5cg+C;5cg;" & _
                                "4cg;3cg;2cg;1cg;0cg"
Private Const LT_ACIERTOSGP As String = "5+1;5;4+1;4;3+1;3;2+1;2;1;0"
Private Const LT_ACIERTOSEU As String = "5+2;5+1;5;4+2;4+1;4;3+2;2+2,3+1;3;1+2;" & _
                                "2+1;2;1+1;1;0+2;0+1;0"

'
'   Matrices de multiples premios según juego y apuestas
'
Private Const LT_MATRIZMULTIBL001 As Variant = "76cg+C1,1;76cg+C2,6;76cg1,1;76cg3" & _
       ",6;75cg+C2,1;75cg+C3,1;75cg+C4,5;75cg3,2;75cg4,2;74cg4,3;74cg5,4;73cg5,4;" & _
       "86cg+C1,1;86cg+C2,6;86cg+C3,6;86cg+C4,15;86cg1,1;86cg3,12;86cg4,15;85cg+C" & _
       "2,1;85cg+C3,2;85cg+C4,15;85cg+C5,20;85cg3,3;85cg4,15;85cg5,10;84cg4,6;84c" & _
       "g3,16;83cg5,10;55cg1,1;55cg2,1;55cg3,42;54cg+C2,2;54cg+C4,42;54cg3,2;54cg" & _
       "4,42;53cg4,3;53cg5,41;52cg5,4;96cg+C1,1;96cg+C2,6;96cg+C3,12;96cg+C4,45;9" & _
       "6cg+C5,20;96cg1,1;96cg3,18;96cg4,45;96cg5,20;95cg+C2,1;95cg+C3,3;95cg+C4," & _
       "30;95cg+C5,40;95cg3,4;95cg4,30;95cg5,40;94cg4,10;94cg5,40;93cg5,20;106cg+" & _
       "C1,1;106cg+C2,6;106cg+C3,18;106cg+C4,90;106cg+C5,80;106cg1,1;106cg3,24;10" & _
       "6cg4,90;106cg5,80;105cg+C2,1;105cg+C3,4;105cg+C4,50;105cg+C5,100;105cg3,5" & _
       "105cg4,50;105cg5,100;104cg4,15;104cg5,80;103cg5,35;116cg+C1,1;116cg+C2,6;" & _
       "116cg+C3,24;116cg+C4,150;116cg+C5,200;116cg1,1;116cg3,30;116cg4,150;116cg" & _
       "5,200;115cg+C2,1;115cg+C3,5;115cg+C4,75;115cg+C5,200;115cg3,6;115cg4,75;1" & _
       "15cg5,200;114cg4,21;114cg5,140;113cg5,56"

Private Const LT_MATRIZMULTILP001 As Variant = "76cg+R+C14,1;76cg+R+C1,1;76cg+R" & _
    "+C2,6;76cg+R14,1;76cg+R1,1;76cg+R3,6;76cg+C1,1;76cg+C2,6;76cg1,1;76cg3,6;7" & _
    "5cg+C2,1;75cg+C3,1;75cg+C4,5;75cg3,2;75cg4,5;74cg4,3;74cg5,4;73cg5,4;86cg+" & _
    "R+C14,1;86cg+R+C1,1;86cg+R+C2,6;86cg+R+C3,6;86cg+R+C4,15;86cg+R14,1;86cg+R" & _
    "1,1;86cg+R3,12;86cg+R4,15;86cg+C1,1;86cg+C2,6;86cg+C3,6;86cg+C4,15;86cg1,1" & _
    ";86cg3,12;86cg4,15;85cg+C2,1;85cg+C3,2;85cg+C4,15;85cg+C5,10;85cg3,3;85cg4" & _
    ",15;85cg5,10;84cg4,6;84cg5,16;83cg5,10;55cg+R14,1;55cg+R1,1;55cg+R2,1;55cg" & _
    "+R3,42;55cg1,1;55cg2,1;55cg3,42;54cg+C2,2;54cg+C4,42;54cg3,2;54cg4,42;53cg" & _
    "4,3;53cg5,41;52cg5,4;96cg+R+C14,1;96cg+R+C1,1;96cg+R+C2,6;96cg+R+C3,12;96c" & _
    "g+R+C4,45;96cg+R+C5,20;96cg+R14,1;96cg+R1,1;96cg+R3,18;96cg+R4,45;96cg+R5," & _
    "20;96cg+C1,1;96cg+C2,6;96cg+C3,12;96cg+C4,45;96cg+C5,20;96cg1,1;96cg3,18;9" & _
    "6cg4,45;96cg5,20;95cg+C2,1;95cg+C3,3;95cg+C4,30;95cg+C5,40;95cg3,4;95cg4,3" & _
    "0;95cg5,40;94cg4,10;94cg5,40;93cg5,20;106cg+R+C14,1;106cg+R+C1,1;106cg+R+C" & _
    "2,6;106cg+R+C3,18;106cg+R+C4,90;106cg+R+C5,80;106cg+R14,1;106cg+R1,1;106cg" & _
    "+R3,24;106cg+R4,90;106cg+R5,80;106cg+C1,1;106cg+C2,6;106cg+C3,18;106cg+C4," & _
    "90;106cg+C5,80;106cg1,1;106cg3,24;106cg4,90;106cg5,80;105cg+C2,1;105cg+C3," & _
    "4;105cg+C4,50;105cg+C5,100;105cg3,5;105cg4,50;105cg5,100;104cg4,15;104cg5," & _
    "80;103cg5,35;116cg+R+C14,1;116cg+R+C1,1;116cg+R+C2,6;116cg+R+C3,24;116cg+R" & _
    "+C4,150;116cg+R+C5,200;116cg+R14,1;116cg+R1,1;116cg+R3,30;116cg+R4,150;116" & _
    "cg+R5,200;116cg+C1,1;116cg+C2,6;116cg+C3,24;116cg+C4,150;116cg+C5,200;116c" & _
    "g1,1;116cg3,30;116cg4,150;116cg5,200;115cg+C2,1;115cg+C3,5;115cg+C4,75;115" & _
    "cg+C5,200;115cg3,6;115cg4,75;115cg5,200;114cg4,21;114cg5,140;113cg5,56"

Private Const LT_MATRIZMULTIGP001 As Variant = "65+11,1;65+13,5;652,1;654,5;64+13" & _
    ",2;64+15,4;644,2;646,4;63+15,3;63+17,3;636,3;638,3;62+17,4;628,4;75+11,1;75+" & _
    "13,10;75+15,10;752,1;754,10;756,10;74+13,3;74+15,12;74+17,6;744,3;746,12;748" & _
    ",6;73+15,6;73+17,12;736,6;738,12;72+17,10;728,10;85+11,1;85+13,15;85+15,30;8" & _
    "5+17,10;852,1;854,15;856,30;858,10;84+13,4;84+15,24;84+17,24;844,4;846,24;84" & _
    "8,24;83+15,10;83+17,30;836,10;838,30;82+17,20;828,20;95+11,1;95+13,20;95+15," & _
    "60;95+17,40;952,1;954,20;956,60;958,40;94+13,5;94+15,40;94+17,60;944,5;946,4" & _
    "0;948,60;93+15,15;93+17,60;936,15;938,60;92+17,35;928,35;105+11,1;105+13,25;" & _
    "105+15,100;105+17,100;1052,1;1054,25;1056,100;1058,100;104+13,6;104+15,60;10" & _
    "4+17,120;1044,6;1046,60;1048,120;103+15,21;103+17,105;1036,21;1038,105;102+1" & _
    "7,56;1028,56;115+11,1;115+13,30;115+15,150;115+17,200;1152,1;1154,30;1156,15" & _
    "0;1158,200;114+13,7;114+15,84;114+17,210;1144,7;1146,84;1148,210;113+15,28;1" & _
    "13+17,168;1136,28;1138,168;112+17,84;1128,84"
    
Private Const LT_MATRIZMULTIEU001 As Variant = ""
Private Const LT_MATRIZMULTIEU002 As Variant = ""
Private Const LT_MATRIZMULTIEU003 As Variant = ""
'


'--- Mensajes -----------------------------------------------------------------*
'--- Errores ------------------------------------------------------------------*
'--- Propiedades --------------------------------------------------------------*

Public CatPremioBoleto  As CategoriaPremio

Public Property Get BolasAcertadas() As Integer
    If JUEGO_DEFECTO = Euromillones Then
        BolasAcertadas = nBolasAcertadas + nEstrellasAcertadas
    Else
        BolasAcertadas = nBolasAcertadas
    End If
End Property

Public Property Get Premio() As Premio
    Set Premio = oPremio
End Property

Public Property Get CatPremioApuesta() As CategoriaPremio
    CatPremioApuesta = mCategoriaApuesta
End Property

Public Property Get Sorteo() As Sorteo
    Set Sorteo = oSorteo
End Property

Public Property Set Sorteo(ByVal vNewValue As Sorteo)
    Set oSorteo = vNewValue
End Property

Public Property Get ImporteApuesta() As Currency
    ImporteApuesta = mImporteApta
End Property

Public Property Get ImporteBoleto() As Currency
    ImporteBoleto = mImporteBlto
End Property

Public Property Get NumerosAcertados() As String
    NumerosAcertados = GetAciertos()
End Property

Public Property Get CategoriaPremioTxt() As String
    If mCategoriaApuesta <> Ninguna Then
        Select Case mCategoriaApuesta
            Case Especial: CategoriaPremioTxt = "Esp"
            Case Primera: CategoriaPremioTxt = "1ª"
            Case Segunda: CategoriaPremioTxt = "2ª"
            Case Tercera: CategoriaPremioTxt = "3ª"
            Case Cuarta: CategoriaPremioTxt = "4ª"
            Case Quinta: CategoriaPremioTxt = "5ª"
            Case Sexta: CategoriaPremioTxt = "6ª"
            Case Septima: CategoriaPremioTxt = "7ª"
            Case Octava: CategoriaPremioTxt = "8ª"
            Case Novena: CategoriaPremioTxt = "9ª"
            Case Decima: CategoriaPremioTxt = "10ª"
            Case Undecima: CategoriaPremioTxt = "11ª"
            Case Duodecima: CategoriaPremioTxt = "12ª"
            Case Trigesimotercera: CategoriaPremioTxt = "13ª"
            Case Reintegro: CategoriaPremioTxt = "R"
        End Select
    Else
        CategoriaPremioTxt = CStr(nBolasAcertadas)
    End If
End Property

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set oSortEng = New SorteoEngine
    Set oSorteo = New Sorteo
    Set oPremio = New Premio
    nBolasAcertadas = 0
    nEstrellasAcertadas = 0
    bComplementarioAcertado = False
    bReintegroAcertado = False
    bElMillonAcertado = False
    mImporteApta = 0
    mImporteBlto = 0
    CatPremioBoleto = Ninguna
End Sub

Private Sub Class_Terminate()
    Set oSortEng = Nothing
    Set oSorteo = Nothing
    Set oPremio = Nothing
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : GetMatrizApuestas
' Fecha          : do., 24/feb/2019 21:22:32
' Propósito      : Construye una colección con la matriz de apuestas multiples
'                  de cada juego
'------------------------------------------------------------------------------*
Private Function GetMatrizApuestas(vNewPronostico As Integer) As Collection
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
    Select Case oSorteo.Juego
        Case Bonoloto:
            mMatriz = Split(LT_MATRIZMULTIBL001, ";")
        
        Case LoteriaPrimitiva:
            mMatriz = Split(LT_MATRIZMULTILP001, ";")
        
        Case GordoPrimitiva:
            mMatriz = Split(LT_MATRIZMULTIGP001, ";")
        
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
        '   Si la clave pertenece al pronostico
        '
        If Left(mElem(0), Len(CStr(vNewPronostico))) = CStr(vNewPronostico) Then
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
        End If
    '
    '   Siguiente elemento de la matriz
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
   Call HandleException(ErrNumber, ErrDescription, "Premios.GetMatrizApuestas")
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function




'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ComprobarApuesta
' Fecha          : lu., 20/jul/2020 18:03:54
' Propósito      : Comprueba los numeros de la apuesta
' Parámetros     : Apuesta
'------------------------------------------------------------------------------*
'
Public Function ComprobarApuesta(vNewApuesta As Apuesta, NoSetFechaApuesta As Boolean) As String
    Dim mNum As Numero
    
  On Error GoTo ComprobarApuesta_Error
    '
    '   Inicializamos la función
    '
    ComprobarApuesta = ""
    '
    '   Cargamos el sorteo
    '
    If oSorteo.Fecha <> vNewApuesta.Fecha And _
    Not (NoSetFechaApuesta) Then
        SetSorteo vNewApuesta.Fecha
        '
        '   Si no hay sorteo para la fecha no evaluamos
        '
        If oSorteo Is Nothing Then
            Exit Function
        End If
    End If
    '
    '   Inicialización de contadores
    '
    nBolasAcertadas = 0
    nEstrellasAcertadas = 0
    bComplementarioAcertado = False
    mCategoriaApuesta = Ninguna
    '
    '   Segun el juego comprobamos
    '
    Select Case vNewApuesta.Juego
    
        Case Bonoloto, LoteriaPrimitiva:
            '
            '   Para cada número en la apuesta
            '
            For Each mNum In vNewApuesta.Combinacion.Numeros
                '
                '
                '
                If oSorteo.Combinacion.Contiene(mNum.Valor) Then
                    nBolasAcertadas = nBolasAcertadas + 1
                End If
                If oSorteo.Complementario = mNum.Valor Then
                    nBolasAcertadas = nBolasAcertadas + 1
                    bComplementarioAcertado = True
                End If
            Next mNum
        
        Case GordoPrimitiva:
            '
            '   Para cada número en la apuesta
            '
            For Each mNum In vNewApuesta.Combinacion.Numeros
                If oSorteo.Combinacion.Contiene(mNum.Valor) Then
                    nBolasAcertadas = nBolasAcertadas + 1
                End If
            Next mNum
        
        Case Euromillones:
            '
            '   Para cada número en la combinación
            '
            For Each mNum In vNewApuesta.Combinacion.Numeros
                If oSorteo.Combinacion.Contiene(mNum.Valor) Then
                    nBolasAcertadas = nBolasAcertadas + 1
                End If
            Next mNum
            '
            '   Para cada número en la apuesta
            '
            For Each mNum In vNewApuesta.Estrellas.Numeros
                If oSorteo.Estrellas.Contiene(mNum.Valor) Then
                    nEstrellasAcertadas = nEstrellasAcertadas + 1
                End If
            Next mNum
    End Select
    '
    '   Definir la categoría del premio
    '
    ComprobarApuesta = GetAciertos
    '
    '
    If mCategoriaApuesta <> Ninguna Then
        '
        '   Calculamos el importe de la apuesta
        '
        mImporteApta = CalImporteApuestas(vNewApuesta)
    Else
        mImporteApta = 0
    End If
        
  On Error GoTo 0
ComprobarApuesta__CleanExit:
    Exit Function
ComprobarApuesta_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ComprobarBoletos.ComprobarApuesta", ErrSource)
    Err.Raise ErrNumber, "ComprobarBoletos.ComprobarApuesta", ErrDescription
End Function




'------------------------------------------------------------------------------*
' Procedimiento  : ComprobarBoleto
' Fecha          : vi., 17/jul/2020 18:31:51
' Propósito      : Comprueba un boleto con el sorteo
' Parámetros     : Boleto a comprobar
'------------------------------------------------------------------------------*
'
Public Function ComprobarBoleto(vNewValue As Boleto) As String
    Dim mPrem           As Premio
    Dim mCatPremApt     As String
    Dim oApt            As Apuesta
    
  On Error GoTo ComprobarBoleto_Error
    '
    '   #TODO: Estudiar vigencia del boleto y bucle de fechas para varios sorteos
    '
    ComprobarBoleto = ""
    CatPremioBoleto = Ninguna
    mImporteBlto = 0
    '
    '   Cargamos el sorteo bien porque no existe o cambia la fecha
    '
    If oSorteo Is Nothing Then
        SetSorteo vNewValue.FechaValidez
        '
        '   Si no hay sorteo para la fecha no evaluamos
        '
        If oSorteo Is Nothing Then
            Exit Function
        End If
    Else
        If oSorteo.Fecha <> vNewValue.FechaValidez Then
            SetSorteo vNewValue.FechaValidez
            '
            '   Si no hay sorteo para la fecha no evaluamos
            '
            If oSorteo Is Nothing Then
                Exit Function
            End If
        End If
    End If
    
    Select Case vNewValue.Juego
        Case Bonoloto, LoteriaPrimitiva, GordoPrimitiva:
            If vNewValue.Reintegro = oSorteo.Reintegro Then
                bReintegroAcertado = True
                Set mPrem = oSorteo.Premios.GetPremioByCategoria(Reintegro)
                mImporteBlto = mPrem.Importe * vNewValue.NumeroApuestas
                ComprobarBoleto = mPrem.CategoriaTexto
                CatPremioBoleto = Reintegro
            Else
                bReintegroAcertado = False
            End If
        
        Case Euromillones:
            If vNewValue.ElMillon = oSorteo.ElMillon Then
                mImporteBlto = 1000000#
                ComprobarBoleto = "El Millon"
                CatPremioBoleto = Reintegro '#TODO: crear una categoria para esto
            End If
            
    End Select
    '
    '   Bucle de apuestas del boleto
    '
    For Each oApt In vNewValue.Apuestas
        '    Iniciamos la categoria
        '
        mCatPremApt = ""
        '   Comprobamos la apuesta
        '
        mCatPremApt = Me.ComprobarApuesta(oApt, False)
        '
        '   Si hay premio se agrega al Boleto
        '
        If mCategoriaApuesta <> Ninguna Then
            If Len(ComprobarBoleto) = 0 Then
                ComprobarBoleto = mCatPremApt
            Else
            '
            '  TODO si es El gordo queda la categoria de premio mas alta de la apuesta
                ComprobarBoleto = ComprobarBoleto & "+" & mCatPremApt
            End If
            mImporteBlto = mImporteBlto + mImporteApta
            CatPremioBoleto = mCategoriaApuesta
        End If
    
    Next oApt
        
  On Error GoTo 0
ComprobarBoleto__CleanExit:
    Exit Function
ComprobarBoleto_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ComprobarBoletos.ComprobarBoleto", ErrSource)
    Err.Raise ErrNumber, "ComprobarBoletos.ComprobarBoleto", ErrDescription
End Function


'------------------------------------------------------------------------------*
' Procedimiento  : GetAciertos
' Fecha          : lu., 20/jul/2020 18:03:54
' Propósito      : Obtiene el las categorias del premio en texto
'------------------------------------------------------------------------------*
'
Private Function GetAciertos() As String
    Dim mVar As Variant
    
  On Error GoTo GetAciertos_Error
    GetAciertos = ""
    mCategoriaApuesta = Ninguna
    
    Select Case oSorteo.Juego
        Case Bonoloto:
            mVar = Split(LT_ACIERTOSBL, ";")
            Select Case nBolasAcertadas
                Case 7:
                    mCategoriaApuesta = Primera
                    If bComplementarioAcertado Then
                        If bReintegroAcertado Then
                            GetAciertos = mVar(0)
                        Else
                            GetAciertos = mVar(1)
                        End If
                    Else
                        If bReintegroAcertado Then
                            GetAciertos = mVar(2)
                        Else
                            GetAciertos = mVar(3)
                        End If
                    End If
                
                Case 6:
                    
                    If bComplementarioAcertado Then
                        mCategoriaApuesta = Segunda
                        GetAciertos = mVar(4)
                    Else
                        mCategoriaApuesta = Primera
                        GetAciertos = mVar(3)
                    End If
                    
                Case 5:
                    If bComplementarioAcertado Then
                        GetAciertos = mVar(6)
                        mCategoriaApuesta = Cuarta
                    Else
                        GetAciertos = mVar(5)
                        mCategoriaApuesta = Tercera
                    End If
                    
                Case 4:
                    If bComplementarioAcertado Then
                        GetAciertos = mVar(7)
                        mCategoriaApuesta = Quinta
                    Else
                        GetAciertos = mVar(6)
                        mCategoriaApuesta = Cuarta
                    End If
                    
                Case 3:
                    If bComplementarioAcertado Then
                        GetAciertos = mVar(8)
                        mCategoriaApuesta = Ninguna
                    Else
                        GetAciertos = mVar(7)
                        mCategoriaApuesta = Quinta
                    End If
                    
                Case 2:
                    GetAciertos = mVar(8)
                    mCategoriaApuesta = Ninguna
                    
                Case 1:
                    GetAciertos = mVar(9)
                    mCategoriaApuesta = Ninguna
                Case Else:
                    GetAciertos = mVar(10)
                    mCategoriaApuesta = Ninguna
            End Select
            
        Case LoteriaPrimitiva:
            mVar = Split(LT_ACIERTOSLP, ";")
            Select Case nBolasAcertadas
                Case 7:
                    If bComplementarioAcertado Then
                        If bReintegroAcertado Then
                            GetAciertos = mVar(0)
                            mCategoriaApuesta = Especial
                        Else
                            GetAciertos = mVar(2)
                            mCategoriaApuesta = Primera
                        End If
                    Else
                        If bReintegroAcertado Then
                            GetAciertos = mVar(1)
                            mCategoriaApuesta = Especial
                        Else
                            GetAciertos = mVar(3)
                            mCategoriaApuesta = Primera
                        End If
                    End If
                
                Case 6:
                    If bComplementarioAcertado Then
                        If bReintegroAcertado Then
                            GetAciertos = mVar(0)
                            mCategoriaApuesta = Especial
                        Else
                            GetAciertos = mVar(2)
                            mCategoriaApuesta = Primera
                        End If
                    Else
                        If bReintegroAcertado Then
                            GetAciertos = mVar(1)
                            mCategoriaApuesta = Especial
                        Else
                            GetAciertos = mVar(3)
                            mCategoriaApuesta = Primera
                        End If
                    End If
                    
                Case 5:
                    If bComplementarioAcertado Then
                        GetAciertos = mVar(4)
                        mCategoriaApuesta = Segunda
                    Else
                        GetAciertos = mVar(5)
                        mCategoriaApuesta = Tercera
                    End If
                    
                Case 4:
                    If bComplementarioAcertado Then
                        GetAciertos = mVar(7)
                        mCategoriaApuesta = Quinta
                    Else
                        GetAciertos = mVar(6)
                        mCategoriaApuesta = Cuarta
                    End If
                    
                Case 3:
                    If bComplementarioAcertado Then
                        GetAciertos = mVar(8)
                        mCategoriaApuesta = Ninguna
                    Else
                        GetAciertos = mVar(7)
                        mCategoriaApuesta = Quinta
                    End If
                    
                Case 2:
                    GetAciertos = mVar(8)
                    mCategoriaApuesta = Ninguna
                    
                Case 1:
                    GetAciertos = mVar(9)
                    mCategoriaApuesta = Ninguna
                Case Else:
                    GetAciertos = mVar(10)
                    mCategoriaApuesta = Ninguna
            End Select
        
        
        
        Case GordoPrimitiva:
            mVar = Split(LT_ACIERTOSGP, ";")
            Select Case nBolasAcertadas
                Case 5:
                    If bReintegroAcertado Then
                        GetAciertos = mVar(0)
                        mCategoriaApuesta = Primera
                    Else
                        GetAciertos = mVar(1)
                        mCategoriaApuesta = Segunda
                    End If
                    
                Case 4:
                    If bReintegroAcertado Then
                        GetAciertos = mVar(2)
                        mCategoriaApuesta = Tercera
                    Else
                        GetAciertos = mVar(3)
                        mCategoriaApuesta = Cuarta
                    End If
                    
                Case 3:
                    If bReintegroAcertado Then
                        GetAciertos = mVar(4)
                        mCategoriaApuesta = Quinta
                    Else
                        GetAciertos = mVar(5)
                        mCategoriaApuesta = Sexta
                    End If
                    
                Case 2:
                    If bReintegroAcertado Then
                        GetAciertos = mVar(6)
                        mCategoriaApuesta = Septima
                    Else
                        GetAciertos = mVar(7)
                        mCategoriaApuesta = Octava
                    End If
                    
                Case 1:
                    GetAciertos = mVar(8)
                    mCategoriaApuesta = Ninguna
                Case Else:
                    GetAciertos = mVar(9)
                    mCategoriaApuesta = Ninguna
            End Select
        
        
        Case Euromillones:
            mVar = Split(LT_ACIERTOSEU, ";")
            
            Select Case nBolasAcertadas
                Case 5:
                    Select Case nEstrellasAcertadas
                        Case 2:
                            GetAciertos = mVar(0)
                            mCategoriaApuesta = Primera
                        Case 1:
                            GetAciertos = mVar(1)
                            mCategoriaApuesta = Segunda
                        Case Else:
                            GetAciertos = mVar(2)
                            mCategoriaApuesta = Tercera
                    End Select
                        
                Case 4:
                    Select Case nEstrellasAcertadas
                        Case 2:
                            GetAciertos = mVar(3)
                            mCategoriaApuesta = Cuarta
                        Case 1:
                            GetAciertos = mVar(4)
                            mCategoriaApuesta = Quinta
                        Case Else:
                            GetAciertos = mVar(5)
                            mCategoriaApuesta = Sexta
                    End Select
                
                Case 3:
                    Select Case nEstrellasAcertadas
                        Case 2:
                            GetAciertos = mVar(6)
                            mCategoriaApuesta = Septima
                        Case 1:
                            GetAciertos = mVar(8)
                            mCategoriaApuesta = Novena
                        Case Else:
                            GetAciertos = mVar(9)
                            mCategoriaApuesta = Decima
                    End Select
                
                Case 2:
                    Select Case nEstrellasAcertadas
                        Case 2:
                            GetAciertos = mVar(7)
                            mCategoriaApuesta = Octava
                        Case 1:
                            GetAciertos = mVar(11)
                            mCategoriaApuesta = Duodecima
                        Case Else:
                            GetAciertos = mVar(12)
                            mCategoriaApuesta = Trigesimotercera
                    End Select
                
                Case 1:
                    Select Case nEstrellasAcertadas
                        Case 2:
                            GetAciertos = mVar(10)
                            mCategoriaApuesta = Undecima
                        Case 1:
                            GetAciertos = mVar(14)
                            mCategoriaApuesta = Ninguna
                        Case Else:
                            GetAciertos = mVar(15)
                            mCategoriaApuesta = Ninguna
                    End Select
                Case Else:
                    mCategoriaApuesta = Ninguna
                    Select Case nEstrellasAcertadas
                        Case 2:
                            GetAciertos = mVar(16)
                        Case 1:
                            GetAciertos = mVar(17)
                        Case Else:
                            GetAciertos = mVar(18)
                    End Select
            End Select
            
            
    End Select
        
  On Error GoTo 0
    Exit Function
GetAciertos_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ComprobarBoletos.GetAciertos", ErrSource)
    Err.Raise ErrNumber, "ComprobarBoletos.GetAciertos", ErrDescription
End Function


'------------------------------------------------------------------------------*
' Procedimiento  : SetSorteo
' Fecha          : lu., 20/jul/2020 18:03:54
' Propósito      : Obtiene el sorteo para su evaluación
' Parámetros     : Fecha de sorteo
'------------------------------------------------------------------------------*
'
Private Sub SetSorteo(vNewValue As Date)
    
  On Error GoTo SetSorteo_Error
    '
    '   Cargamos el motor
    '
    Set oSortEng = New SorteoEngine
    '
    '   Buscamos el sorteo
    '
    Set oSorteo = oSortEng.GetSorteoByFecha(vNewValue)
        
  On Error GoTo 0
    Exit Sub
SetSorteo_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ComprobarBoletos.SetSorteo", ErrSource)
    Err.Raise ErrNumber, "ComprobarBoletos.SetSorteo", ErrDescription
End Sub



'------------------------------------------------------------------------------*
' Procedimiento  : CalImporteApuestas
' Fecha          : #TODO: Calcular apuestas múltiples
' Propósito      :
' Parámetros     :
'------------------------------------------------------------------------------*
'
Private Function CalImporteApuestas(vNewValue As Apuesta) As Currency
    Dim mMtzApuestas    As Variant
    Dim mDupla          As Dupla
    Dim mStrKey         As String
    Dim mApuestas       As Integer
    Dim mPremio         As Premio
     
  On Error GoTo CalImporteApuestas_Error
    CalImporteApuestas = 0
    '
    '   Inicializamos el premio
    '
    Set mPremio = New Premio
    '
    '   Segun el juego calculamos importes de apuesta múltiple
    '
    Select Case oSorteo.Juego
    
        Case Bonoloto, LoteriaPrimitiva:
            '
            '   Si es apuesta múltiple
            '
            If vNewValue.EsMultiple Then
                '
                '   Si es multiple cargamos la matriz de premios
                '
                Set mMtzApuestas = GetMatrizApuestas(vNewValue.Pronosticos)
                '
                '   Creamos el objeto Dupla
                '
                Set mDupla = New Dupla
                '
                '  calculamos el importe para todos los premios igual o inferiores a la categoria
                '  correspondiente
                '
                For Each mPremio In Me.Sorteo.Premios.Items
                    '
                    '   Clave = Pronosticos + Numero Apuestas + Numeros Acertados + Categoria Premio
                    '
                    mStrKey = CStr(vNewValue.Pronosticos) & NumerosAcertados & mPremio.CategoriaPremio
                    On Error Resume Next
                    Set mDupla = mMtzApuestas(mStrKey)
                    If Err.Number = 5 Then
                        mApuestas = 0
                    Else
                        mApuestas = Val(mDupla.Value)
                    End If
                    If mPremio.CategoriaPremio >= mCategoriaApuesta _
                    And mPremio.CategoriaPremio < Reintegro Then
                        CalImporteApuestas = CalImporteApuestas + (mPremio.Importe * mApuestas)
                    End If
                Next mPremio
            Else
                '
                '   Obtenemos el premio correspondiente a la categoria
                '
                Set oPremio = oSorteo.Premios.GetPremioByCategoria(mCategoriaApuesta)
                CalImporteApuestas = oPremio.Importe
            End If
        
        Case GordoPrimitiva:
            '
            '   Si es apuesta múltiple
            '
            If vNewValue.EsMultiple Then
                '
                '   Si es multiple cargamos la matriz de premios
                '
                Set mMtzApuestas = GetMatrizApuestas(vNewValue.Pronosticos)
                '
                '   Creamos el objeto Dupla
                '
                Set mDupla = New Dupla
                '
                '  calculamos el importe para todos los premios igual o inferiores a la categoria
                '  correspondiente
                '
                For Each mPremio In Me.Sorteo.Premios.Items
                    '
                    '   Clave = Pronosticos + Numero Apuestas + Numeros Acertados + Categoria Premio
                    '
                    mStrKey = CStr(vNewValue.Pronosticos) & NumerosAcertados & mPremio.CategoriaPremio
                    On Error Resume Next
                    Set mDupla = mMtzApuestas(mStrKey)
                    If Err.Number = 5 Then
                        mApuestas = 0
                    Else
                        mApuestas = Val(mDupla.Value)
                    End If
                    If mPremio.CategoriaPremio >= mCategoriaApuesta _
                    And mPremio.CategoriaPremio < Reintegro Then
                        CalImporteApuestas = CalImporteApuestas + (mPremio.Importe * mApuestas)
                    End If
                Next mPremio
            Else
                '
                '   Obtenemos el premio correspondiente a la categoria
                '
                Set oPremio = oSorteo.Premios.GetPremioByCategoria(mCategoriaApuesta)
                CalImporteApuestas = oPremio.Importe
            End If
        
        Case Euromillones:
            '#TODO:
                CalImporteApuestas = 0
        
    End Select

  On Error GoTo 0
    Exit Function
CalImporteApuestas_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ComprobarBoletos.CalImporteApuestas", ErrSource)
    Err.Raise ErrNumber, "ComprobarBoletos.CalImporteApuestas", ErrDescription
End Function
' *===========(EOF): ComprobarBoletos.cls


