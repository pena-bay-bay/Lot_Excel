VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComprobarSugerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : ComprobarSugerencia.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : ju., 04/feb/2021 20:09:04
' *     Modificado :
' *     Versión    : 1.0
' *     Propósito  : Funcionalidad de comprobar una sugerencia contra un Sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0

'--- Variables Privadas -------------------------------------------------------*
Private oSorteo                 As Sorteo
Private oPremio                 As Premio
Private oSugerencia             As Sugerencia
Private nBolasAcertadas         As Integer
Private nEstrellasAcertadas     As Integer
Private bComplementarioAcertado As Boolean
Private bReintegroAcertado      As Boolean
Private mImporteApta            As Currency
Private mCategoriaApuesta       As CategoriaPremio
Private mPuntuacion             As Integer
Private mAciertos               As String
Private mMatMult                As Collection

'--- Constantes ---------------------------------------------------------------*
Private Const LT_ACIERTOSBL As String = "6cg+R+C;6cg+R;6cg+C;6cg;5cg+C;5cg;" & _
                                "4cg;3cg;2cg;1cg;0cg"
Private Const LT_ACIERTOSLP As String = "6cg+R+C;6cg+R;6cg+C;6cg;5cg+C;5cg;" & _
                                "4cg;3cg;2cg;1cg;0cg"
Private Const LT_ACIERTOSGP As String = "5+1;5;4+1;4;3+1;3;2+1;2;1;0"
Private Const LT_ACIERTOSEU As String = "5+2;5+1;5;4+2;4+1;4;3+2;2+2;3+1;3;1+2;" & _
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
    
Private Const LT_MATRIZMULTIEU001 As Variant = "535+21,1;535+22,2;535+12,2;535+13" & _
    ",1;5353,3;534+24,1;534+25,2;534+15,2;534+16,1;5346,3;533+27,1;533+29,2;532+2" & _
    "8,1;532+212,2;533+19,2;533+110,1;53310,3;531+211,1;532+112,2;532+113,1;53213" & _
    ",3;545+21,1;545+22,4;545+23,1;545+12,3;545+13,3;5453,6;544+24,1;544+25,4;544" & _
    "+26,1;544+15,3;544+16,3;5446,6;543+27,1;543+29,4;543+210,1;542+28,1;542+212," & _
    "4;542+213,1;543+19,3;543+110,3;54310,6;541+211,1;542+112,3;542+113,3;54213,6" & _
    ";555+21,1;555+22,4;555+23,1;555+12,3;555+13,3;5553,6;554+24,1;554+25,4;554+2" & _
    "6,1;554+15,3;554+16,3;5546,6;553+27,1;553+29,4;553+210,1;552+28,1;552+212,4;" & _
    "552+213,1;553+19,3;553+110,3;55310,6;551+211,1;552+112,3;552+113,3;55213,6;"

Private Const LT_MATRIZMULTIEU002 As Variant = "625+21,1;625+24,5;625+12,1;625+15," & _
    "5;6253,1;6256,5;624+24,2;624+27,4;624+15,2;624+19,4;6246,2;62410,4;623+27,3;" & _
    "623+28,3;622+28,4;622+211,4;623+19,3;623+112,3;62310,3;62313,3;621+211,11;62" & _
    "2+112,4;62213,4;635+21,1;635+22,2;635+24,5;635+25,10;635+12,2;635+13,1;635+1" & _
    "5,10;635+16,5;6353,3;6356,15;634+24,2;634+25,4;634+27,4;634+29,8;634+15,4;63" & _
    "4+16,2;634+19,8;634+110,4;6346,6;63410,12;633+27,3;633+28,3;633+29,6;633+212" & _
    ",6;632+28,4;632+211,2;632+212,8;633+19,6;633+110,3;633+112,6;633+113,3;63310" & _
    ",9;63313,9;631+211,5;632+112,8;632+113,4;63213,12;645+21,1;645+22,4;645+23,1" & _
    ";645+24,5;645+25,20;645+26,5;645+12,3;645+13,3;645+15,15;645+16,15;6453,6;64" & _
    "56,30;644+24,2;644+25,8;644+26,2;644+27,4;644+29,16;644+210,4;644+15,6;644+1" & _
    "6,6;644+19,12;644+110,12;6446,12;64410,24;643+27,3;643+28,3;643+29,12;643+21" & _
    "0,3;643+212,12;643+213,3;642+28,4;642+211,2;642+212,16;642+213,4;643+19,9;64" & _
    "3+110,9;643+112,9;643+113,9;64310,18;64313,18;641+211,5;642+112,12;642+113,1" & _
    "2;64213,24;655+21,1;655+22,6;655+23,3;655+24,5;655+25,30;655+26,15;655+12,4;" & _
    "655+13,6;655+15,20;655+16,30;6553,10;6556,50;654+24,2;654+25,12;654+26,6;654" & _
    "+27,4;654+29,24;654+210,12;654+15,8;654+16,12;654+19,16;654+110,24;6546,20;6" & _
    "5410,40;653+27,3;653+28,3;653+29,18;653+210,9;653+212,18;653+213,9;652+28,4;" & _
    "652+211,2;652+212,24;652+213,12;653+19,12;653+110,18;653+112,12;653+113,18;6" & _
    "5310,30;65313,30;651+211,5;652+112,16;652+113,24;65213,40"

Private Const LT_MATRIZMULTIEU003a As Variant = "725+21,1;725+24,10;725+27,10;725+" & _
    "12,1;725+15,10;725+19,10;7253,3;7256,3;72510,3;724+24,3;724+27,12;724+28,6;7" & _
    "24+15,3;724+19,12;724+112,6;7246,3;72410,12;72413,6;723+27,6;723+28,12;723+2" & _
    "11,3;722+28,10;722+211,10;723+19,6;723+112,12;72310,3;72313,12;721+211,15;72" & _
    "2+112,12;72213,10;735+21,1;735+22,2;735+24,10;735+25,20;735+27,10;735+29,20;" & _
    "735+12,2;735+13,1;735+15,20;735+16,10;735+19,20;735+110,10;7353,3;7356,30;73" & _
    "510,30734+24,3;734+25,6;734+27,12;734+28,6;734+29,24;734+212,12;734+15,6;734" & _
    "+16,3;734+19,24;734+110,12;734+112,12;734+113,6;7346,9;73410,36;73413,18;733" & _
    "+27,6;733+28,12;733+29,12;733+211,3;733+212,24;732+28,10;732+211,10;732+212," & _
    "20;733+19,12;733+110,6;733+112,24;733+113,12;73310,18;73313,36;731+211,15;73" & _
    "2+112,20;732+113,10;73213,30;745+21,1;745+22,4;745+23,1;745+24,10;745+25,40;" & _
    "745+26,10;745+27,10;745+29,40;745+210,10;745+12,3;745+13,3;745+15,30;745+16," & _
    "30;745+19,30;745+110,30;7453,6;7456,60;74510,60;744+24,3;744+25,12;744+26,3;" & _
    "744+27,12;744+28,6;744+29,48;744+210,12;744+212,24;744+213,6;744+15,9;744+16" & _
    ",9;744+18,36;744+19,36;744+111,18;744+112,18;7446,18;74410,72;74413,36;743+2" & _
    "7,6;743+28,12;743+29,24;743+210,6;743+211,3;743+212,48;743+213,12;742+28,10;" & _
    "742+211,10;742+212,40;742+213,10;743+19,18;743+110,18;743+112,36;743+113,36;" & _
    "74310,36;74313,72;741+211,15;742+112,30;742+113,30;74213,60;755+21,1;755+22," & _
    "6;755+23,3;755+24,10;755+25,60;755+26,30;755+27,10;755+29,60;755+210,30;755+" & _
    "12,4;755+13,6;755+15,40;755+15,60;755+16,30;755+19,40;755+110,60;7553,10;755" & _
    "6,100;75510,100;754+24,3;754+25,18;754+26,9;754+27,12;754+28,6;754+29,72;754" & _
    "+210,36;754+212,36;754+213,18;754+15,12;754+16,18;754+19,48;754+110,72;754+1" & _
    "12,24;754+113,36;7546,30;75410,120;75413,60;753+27,6;753+28,12;753+29,36;753" & _
    "+210,18;753+211,3;753+212,72;753+213,36;752+28,10;752+211,10;752+212,60;752+" & _
    "213,30;753+19,24;753+110,36;753+112,48;753+113,72;75310,60;75313,120;751+211"
Private Const LT_MATRIZMULTIEU003b As Variant = ",15;752+112,40;752+113,60;75213,100"


'--- Propiedades --------------------------------------------------------------*
Public Property Get Aciertos() As String
    Aciertos = mAciertos
End Property

Public Property Get NumerosAcertados() As Integer
    NumerosAcertados = nBolasAcertadas
End Property

Public Property Get EstrellasAcertadas() As Integer
    EstrellasAcertadas = nEstrellasAcertadas
End Property

Public Property Get ComplementarioAcertado() As Boolean
    ComplementarioAcertado = bComplementarioAcertado
End Property

Public Property Get ReintegroAcertado() As Boolean
    ReintegroAcertado = bReintegroAcertado
End Property

Public Property Get Premio() As Premio
    Set Premio = oPremio
End Property

Public Property Get CatPremioApuesta() As CategoriaPremio
    CatPremioApuesta = mCategoriaApuesta
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
        CategoriaPremioTxt = ""
    End If
End Property


Public Property Get Sorteo() As Sorteo
    Set Sorteo = oSorteo
End Property

Public Property Set Sorteo(ByVal vNewValue As Sorteo)
    Set oSorteo = vNewValue
End Property



Public Property Get Sugerencia() As Sugerencia
    Set Sugerencia = oSugerencia
End Property

Public Property Set Sugerencia(ByVal vNewValue As Sugerencia)
    Set oSugerencia = vNewValue
End Property


Public Property Get ImporteApuesta() As Currency
    ImporteApuesta = mImporteApta
End Property

Public Property Get Puntuacion() As Integer
    Puntuacion = mPuntuacion
End Property


'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set oSorteo = New Sorteo
    Set oSugerencia = New Sugerencia
    Set oPremio = New Premio
    nBolasAcertadas = 0
    nEstrellasAcertadas = 0
    bComplementarioAcertado = False
    bReintegroAcertado = False
    mImporteApta = 0
    mCategoriaApuesta = Ninguna
    mPuntuacion = 0
End Sub

Private Sub Class_Terminate()
    Set oSugerencia = Nothing
    Set oSorteo = Nothing
    Set oPremio = Nothing
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : SetCategoriaPremio
' Fecha          : vi., 05/feb/2021 19:30:40
' Propósito      : Obtiene el las categorias del premio en texto
'------------------------------------------------------------------------------*
'
Private Sub SetCategoriaPremio()
    Dim mVar As Variant
  On Error GoTo SetCategoriaPremio_Error
    '
    '   Inicializamos categoria del premio
    '
    mCategoriaApuesta = Ninguna
    mAciertos = ""
    
    Select Case oSorteo.Juego
        Case Bonoloto:
            mVar = Split(LT_ACIERTOSBL, ";")
            Select Case nBolasAcertadas
                Case 7:
                    mCategoriaApuesta = Primera
                    If bComplementarioAcertado Then
                        If bReintegroAcertado Then
                            mAciertos = mVar(0)
                        Else
                            mAciertos = mVar(1)
                        End If
                    Else
                        If bReintegroAcertado Then
                            mAciertos = mVar(2)
                        Else
                            mAciertos = mVar(3)
                        End If
                    End If
                
                Case 6:
                    
                    If bComplementarioAcertado Then
                        mCategoriaApuesta = Segunda
                        mAciertos = mVar(4)
                    Else
                        mCategoriaApuesta = Primera
                        mAciertos = mVar(3)
                    End If
                    
                Case 5:
                    If bComplementarioAcertado Then
                        mAciertos = mVar(6)
                        mCategoriaApuesta = Cuarta
                    Else
                        mAciertos = mVar(5)
                        mCategoriaApuesta = Tercera
                    End If
                    
                Case 4:
                    If bComplementarioAcertado Then
                        mAciertos = mVar(7)
                        mCategoriaApuesta = Quinta
                    Else
                        mAciertos = mVar(6)
                        mCategoriaApuesta = Cuarta
                    End If
                    
                Case 3:
                    If bComplementarioAcertado Then
                        mAciertos = mVar(8)
                        mCategoriaApuesta = Ninguna
                    Else
                        mAciertos = mVar(7)
                        mCategoriaApuesta = Quinta
                    End If
                    
                Case 2:
                    mAciertos = mVar(8)
                    mCategoriaApuesta = Ninguna
                    
                Case 1:
                    mAciertos = mVar(9)
                    mCategoriaApuesta = Ninguna
                Case Else:
                    mAciertos = mVar(10)
                    mCategoriaApuesta = Ninguna
            End Select
            
        Case LoteriaPrimitiva:
            mVar = Split(LT_ACIERTOSLP, ";")
            Select Case nBolasAcertadas
                Case 7:
                    If bComplementarioAcertado Then
                        If bReintegroAcertado Then
                            mAciertos = mVar(0)
                            mCategoriaApuesta = Especial
                        Else
                            mAciertos = mVar(2)
                            mCategoriaApuesta = Primera
                        End If
                    Else
                        If bReintegroAcertado Then
                            mAciertos = mVar(1)
                            mCategoriaApuesta = Especial
                        Else
                            mAciertos = mVar(3)
                            mCategoriaApuesta = Primera
                        End If
                    End If
                
                Case 6:
                    If bComplementarioAcertado Then
                        If bReintegroAcertado Then
                            mAciertos = mVar(0)
                            mCategoriaApuesta = Especial
                        Else
                            mAciertos = mVar(2)
                            mCategoriaApuesta = Primera
                        End If
                    Else
                        If bReintegroAcertado Then
                            mAciertos = mVar(1)
                            mCategoriaApuesta = Especial
                        Else
                            mAciertos = mVar(3)
                            mCategoriaApuesta = Primera
                        End If
                    End If
                    
                Case 5:
                    If bComplementarioAcertado Then
                        mAciertos = mVar(4)
                        mCategoriaApuesta = Segunda
                    Else
                        mAciertos = mVar(5)
                        mCategoriaApuesta = Tercera
                    End If
                    
                Case 4:
                    If bComplementarioAcertado Then
                        mAciertos = mVar(7)
                        mCategoriaApuesta = Quinta
                    Else
                        mAciertos = mVar(6)
                        mCategoriaApuesta = Cuarta
                    End If
                    
                Case 3:
                    If bComplementarioAcertado Then
                        mAciertos = mVar(8)
                        mCategoriaApuesta = Ninguna
                    Else
                        mAciertos = mVar(7)
                        mCategoriaApuesta = Quinta
                    End If
                    
                Case 2:
                    mAciertos = mVar(8)
                    mCategoriaApuesta = Ninguna
                    
                Case 1:
                    mAciertos = mVar(9)
                    mCategoriaApuesta = Ninguna
                Case Else:
                    mAciertos = mVar(10)
                    mCategoriaApuesta = Ninguna
            End Select
        
        
        
        Case GordoPrimitiva:
            mVar = Split(LT_ACIERTOSGP, ";")
            Select Case nBolasAcertadas
                Case 5:
                    If bReintegroAcertado Then
                        mAciertos = mVar(0)
                        mCategoriaApuesta = Primera
                    Else
                        mAciertos = mVar(1)
                        mCategoriaApuesta = Segunda
                    End If
                    
                Case 4:
                    If bReintegroAcertado Then
                        mAciertos = mVar(2)
                        mCategoriaApuesta = Tercera
                    Else
                        mAciertos = mVar(3)
                        mCategoriaApuesta = Cuarta
                    End If
                    
                Case 3:
                    If bReintegroAcertado Then
                        mAciertos = mVar(4)
                        mCategoriaApuesta = Quinta
                    Else
                        mAciertos = mVar(5)
                        mCategoriaApuesta = Sexta
                    End If
                    
                Case 2:
                    If bReintegroAcertado Then
                        mAciertos = mVar(6)
                        mCategoriaApuesta = Septima
                    Else
                        mAciertos = mVar(7)
                        mCategoriaApuesta = Octava
                    End If
                    
                Case 1:
                    mAciertos = mVar(8)
                    mCategoriaApuesta = Ninguna
                Case Else:
                    mAciertos = mVar(9)
                    mCategoriaApuesta = Ninguna
            End Select
        
        
        Case Euromillones:
            mVar = Split(LT_ACIERTOSEU, ";")
            
            Select Case nBolasAcertadas
                Case 5:
                    Select Case nEstrellasAcertadas
                        Case 2:
                            mAciertos = mVar(0)
                            mCategoriaApuesta = Primera
                        Case 1:
                            mAciertos = mVar(1)
                            mCategoriaApuesta = Segunda
                        Case Else:
                            mAciertos = mVar(2)
                            mCategoriaApuesta = Tercera
                    End Select
                        
                Case 4:
                    Select Case nEstrellasAcertadas
                        Case 2:
                            mAciertos = mVar(3)
                            mCategoriaApuesta = Cuarta
                        Case 1:
                            mAciertos = mVar(4)
                            mCategoriaApuesta = Quinta
                        Case Else:
                            mAciertos = mVar(5)
                            mCategoriaApuesta = Sexta
                    End Select
                
                Case 3:
                    Select Case nEstrellasAcertadas
                        Case 2:
                            mAciertos = mVar(6)
                            mCategoriaApuesta = Septima
                        Case 1:
                            mAciertos = mVar(8)
                            mCategoriaApuesta = Novena
                        Case Else:
                            mAciertos = mVar(9)
                            mCategoriaApuesta = Decima
                    End Select
                
                Case 2:
                    Select Case nEstrellasAcertadas
                        Case 2:
                            mAciertos = mVar(7)
                            mCategoriaApuesta = Octava
                        Case 1:
                            mAciertos = mVar(11)
                            mCategoriaApuesta = Duodecima
                        Case Else:
                            mAciertos = mVar(12)
                            mCategoriaApuesta = Trigesimotercera
                    End Select
                
                Case 1:
                    Select Case nEstrellasAcertadas
                        Case 2:
                            mAciertos = mVar(10)
                            mCategoriaApuesta = Undecima
                        Case 1:
                            mAciertos = mVar(13)
                            mCategoriaApuesta = Ninguna
                        Case Else:
                            mAciertos = mVar(14)
                            mCategoriaApuesta = Ninguna
                    End Select
                Case Else:
                    mCategoriaApuesta = Ninguna
                    Select Case nEstrellasAcertadas
                        Case 2:
                            mAciertos = mVar(15)
                        Case 1:
                            mAciertos = mVar(16)
                        Case Else:
                            mAciertos = mVar(17)
                    End Select
            End Select
            
            
    End Select
  
  On Error GoTo 0
    Exit Sub
SetCategoriaPremio_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ComprobarSugerencia.SetCategoriaPremio", ErrSource)
    Err.Raise ErrNumber, "ComprobarSugerencia.SetCategoriaPremio", ErrDescription
End Sub


'------------------------------------------------------------------------------*
' Procedimiento  : SetPuntuacion
' Fecha          : vi., 05/feb/2021 19:30:40
' Propósito      : Obtiene la puntuación para la sugerencia con respecto al
'                  sorteo
'------------------------------------------------------------------------------*
'
Private Sub SetPuntuacion()
    Dim iResult As Integer
  On Error GoTo SetPuntuacion_Error
    
    Select Case oSorteo.Juego
        Case Bonoloto
            Select Case nBolasAcertadas
                Case 0: iResult = 0
                Case 1: iResult = 200
                Case 2: iResult = 400
                Case 3:
                    If mCategoriaApuesta = Quinta Then
                        iResult = 900
                    Else
                        iResult = 800
                    End If
                Case 4: iResult = 1000
                Case 5: iResult = 1100
                Case 6: iResult = 1500
                Case 7: iResult = 2000
            End Select
        
        Case LoteriaPrimitiva
            Select Case nBolasAcertadas
                Case 0: iResult = 0
                Case 1: iResult = 200
                Case 2: iResult = 400
                Case 3:
                    If mCategoriaApuesta = Quinta Then
                        iResult = 900
                    Else
                        iResult = 800
                    End If
                Case 4: iResult = 1000
                Case 5: iResult = 1100
                Case 6: iResult = 1500
                Case 7: iResult = 2000
            End Select
            
        Case Euromillones
            Select Case mCategoriaApuesta
                Case Primera: iResult = 2000
                Case Segunda: iResult = 1900
                Case Tercera: iResult = 1800
                Case Cuarta: iResult = 1700
                Case Quinta: iResult = 1600
                Case Sexta: iResult = 1500
                Case Septima: iResult = 1000
                Case Octava: iResult = 900
                Case Novena: iResult = 800
                Case Decima: iResult = 700
                Case Undecima: iResult = 600
                Case Duodecima: iResult = 500
                Case Trigesimotercera: iResult = 400
                Case Ninguna:
                    If nBolasAcertadas = 1 Then
                        If nEstrellasAcertadas = 1 Then
                            iResult = 300
                        Else
                            iResult = 200
                        End If
                    Else
                        If nEstrellasAcertadas = 1 Then
                            iResult = 100
                        Else
                            iResult = 0
                        End If
                    End If
            End Select
            
        Case GordoPrimitiva
            Select Case mCategoriaApuesta
                Case Primera: iResult = 2000
                Case Segunda: iResult = 1900
                Case Tercera: iResult = 1800
                Case Cuarta: iResult = 1700
                Case Quinta: iResult = 1600
                Case Sexta: iResult = 1500
                Case Septima: iResult = 1000
                Case Octava: iResult = 500
                Case Ninguna:
                    If nBolasAcertadas = 1 Then
                        If nEstrellasAcertadas = 1 Then
                            iResult = 300
                        Else
                            iResult = 200
                        End If
                    Else
                        If nEstrellasAcertadas = 1 Then
                            iResult = 100
                        Else
                            iResult = 0
                        End If
                    End If
            End Select
    End Select
    
    mPuntuacion = iResult
    
  On Error GoTo 0
    Exit Sub
SetPuntuacion_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ComprobarSugerencia.SetPuntuacion", ErrSource)
    Err.Raise ErrNumber, "ComprobarSugerencia.SetPuntuacion", ErrDescription
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
            mMatriz = Split(LT_MATRIZMULTIEU001 & LT_MATRIZMULTIEU002 & _
                            LT_MATRIZMULTIEU003a & LT_MATRIZMULTIEU003b, ";")
            
    End Select
    '
    '   Para cada elemento en la matriz
    '
   
    For i = 0 To UBound(mMatriz) - 1
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


'------------------------------------------------------------------------------*
' Procedimiento  : CalImporteSugerencia
' Fecha          : do., 07/feb/2021 18:53:32
' Propósito      : Obtiene el importe de una sugerencia si es premiada
'------------------------------------------------------------------------------*
'
Private Function CalImporteSugerencia() As Currency
    Dim mMtzApuestas    As Variant
    Dim mDupla          As Dupla
    Dim mStrKey         As String
    Dim mApuestas       As Integer
    Dim mPremio         As Premio
    
  On Error GoTo CalImporteSugerencia_Error
    '
    '   Inicializamos la función
    '
    CalImporteSugerencia = 0
    '
    '   Creamos el premio
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
            If oSugerencia.EsMultiple Then
                '
                '   Si es multiple cargamos la matriz de premios
                '
                Set mMtzApuestas = GetMatrizApuestas(oSugerencia.Combinacion.Count)
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
                    mStrKey = CStr(oSugerencia.Combinacion.Count) & mAciertos & mPremio.CategoriaPremio
                    On Error Resume Next
                    Set mDupla = mMtzApuestas(mStrKey)
                    If Err.Number = 5 Then
                        mApuestas = 0
                    Else
                        mApuestas = Val(mDupla.Value)
                    End If
                    If mPremio.CategoriaPremio >= mCategoriaApuesta _
                    And mPremio.CategoriaPremio < Reintegro Then
                        CalImporteSugerencia = CalImporteSugerencia + (mPremio.Importe * mApuestas)
                    End If
                Next mPremio
            Else
                '
                '   Obtenemos el premio correspondiente a la categoria
                '
                Set oPremio = oSorteo.Premios.GetPremioByCategoria(mCategoriaApuesta)
                CalImporteSugerencia = oPremio.Importe
            End If
        
        Case GordoPrimitiva:
            '
            '   Si es apuesta múltiple
            '
            If oSugerencia.EsMultiple Then
                '
                '   Si es multiple cargamos la matriz de premios
                '
                Set mMtzApuestas = GetMatrizApuestas(oSugerencia.Combinacion.Count)
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
                    '   Clave = Pronosticos + Numero Apuestas + Aciertos + Categoria Premio
                    '
                    mStrKey = CStr(oSugerencia.Combinacion.Count) & mAciertos & mPremio.CategoriaPremio
                    On Error Resume Next
                    Set mDupla = mMtzApuestas(mStrKey)
                    If Err.Number = 5 Then
                        mApuestas = 0
                    Else
                        mApuestas = Val(mDupla.Value)
                    End If
                    If mPremio.CategoriaPremio >= mCategoriaApuesta _
                    And mPremio.CategoriaPremio < Reintegro Then
                        CalImporteSugerencia = CalImporteSugerencia + (mPremio.Importe * mApuestas)
                    End If
                Next mPremio
            Else
                '
                '   Obtenemos el premio correspondiente a la categoria
                '
                Set oPremio = oSorteo.Premios.GetPremioByCategoria(mCategoriaApuesta)
                CalImporteSugerencia = oPremio.Importe
            End If
        
        Case Euromillones:
            '
            '   Si es apuesta múltiple
            '
            If oSugerencia.EsMultiple Then
                '
                '   Si es multiple cargamos la matriz de premios
                '
                Set mMtzApuestas = GetMatrizApuestas(oSugerencia.Combinacion.Count)
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
                    '   Clave = Numero Apuestas + Estrellas + Numeros Acertados + Categoria Premio
                    '
                    mStrKey = CStr(oSugerencia.Combinacion.Count) & CStr(oSugerencia.Combinacion.Count) & _
                              NumerosAcertados & mPremio.CategoriaPremio
                    
                    On Error Resume Next
                    Set mDupla = mMtzApuestas(mStrKey)
                    If Err.Number = 5 Then
                        mApuestas = 0
                    Else
                        mApuestas = Val(mDupla.Value)
                    End If
                    If mPremio.CategoriaPremio >= mCategoriaApuesta _
                    And mPremio.CategoriaPremio < Reintegro Then
                        CalImporteSugerencia = CalImporteSugerencia + (mPremio.Importe * mApuestas)
                    End If
                Next mPremio
            Else
                '
                '   Obtenemos el premio correspondiente a la categoria
                '
                Set oPremio = oSorteo.Premios.GetPremioByCategoria(mCategoriaApuesta)
                CalImporteSugerencia = oPremio.Importe
            End If
    End Select
  
  On Error GoTo 0
    Exit Function
CalImporteSugerencia_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ComprobarSugerencia.CalImporteSugerencia", ErrSource)
    Err.Raise ErrNumber, "ComprobarSugerencia.CalImporteSugerencia", ErrDescription
End Function




'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ComprobarSugerencia
' Fecha          : do., 07/feb/2021 18:42:17
' Propósito      : Comprueba los numeros de la sugerencia contra un sorteo
' Parámetros     : Sugerencia
'------------------------------------------------------------------------------*
'
Public Sub ComprobarSugerencia(vNewValue As Sugerencia)
    Dim mNum As Numero
    
  On Error GoTo ComprobarSugerencia_Error
    '
    '   Si no hay sorteo para comprobar
    '
    If oSorteo Is Nothing Then
        Exit Sub
    End If
    '
    '   Actualizamos la sugerencia
    '
    Set oSugerencia = vNewValue
    '
    '   Inicialización de contadores
    '
    nBolasAcertadas = 0
    nEstrellasAcertadas = 0
    bComplementarioAcertado = False
    bReintegroAcertado = False
    mCategoriaApuesta = Ninguna
    '
    '   Segun el juego comprobamos
    '
    Select Case vNewValue.Juego
    
        Case Bonoloto, LoteriaPrimitiva:
            '
            '   Para cada número en la apuesta
            '
            For Each mNum In vNewValue.Combinacion.Numeros
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
            For Each mNum In vNewValue.Combinacion.Numeros
                If oSorteo.Combinacion.Contiene(mNum.Valor) Then
                    nBolasAcertadas = nBolasAcertadas + 1
                End If
            Next mNum
        
        Case Euromillones:
            '
            '   Para cada número en la combinación
            '
            For Each mNum In vNewValue.Combinacion.Numeros
                If oSorteo.Combinacion.Contiene(mNum.Valor) Then
                    nBolasAcertadas = nBolasAcertadas + 1
                End If
            Next mNum
            '
            '   Para cada número en las estrellas
            '
            For Each mNum In vNewValue.Estrellas.Numeros
                If oSorteo.Estrellas.Contiene(mNum.Valor) Then
                    nEstrellasAcertadas = nEstrellasAcertadas + 1
                End If
            Next mNum
    End Select
    '
    '   Definir la categoría del premio
    '
    SetCategoriaPremio
    '
    '   Evaluar sugerencia
    '
    SetPuntuacion
    '
    '   Calcular importes premio
    '
    If mCategoriaApuesta <> Ninguna Then
        '
        '   Calculamos el importe de la Sugerencia
        '
        mImporteApta = CalImporteSugerencia
    Else
        mImporteApta = 0
    End If
        
  On Error GoTo 0
ComprobarSugerencia__CleanExit:
    Exit Sub
ComprobarSugerencia_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "ComprobarSugerencia.ComprobarSugerencia", ErrSource)
    Err.Raise ErrNumber, "ComprobarSugerencia.ComprobarSugerencia", ErrDescription
End Sub




'------------------------------------------------------------------------------*
' Procedimiento  : EstaPremiada
' Fecha          : do., 07/feb/2021 18:42:17
' Propósito      : Si la combinación está premiada devuelve true
'------------------------------------------------------------------------------*
'
Public Function EstaPremiada()
    If Me.CatPremioApuesta <> Ninguna Then
        EstaPremiada = True
    Else
        EstaPremiada = False
    End If
End Function


' *===========(EOF): ComprobarSugerencia.cls



