VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SorteoEditarView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : SorteoEditarView.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : mi., 17/abr/2019 22:35:26
' *     Versión    : 1.0
' *     Propósito  : Interfaz de Usuario Vista Editar Sorteo MVC
' *
' *============================================================================*
Option Explicit
Option Base 0
'
'
'--- Variables Privadas -------------------------------------------------------*
Private mwsUserInterface As Worksheet

'--- Constantes ---------------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : GetDisplaySorteo
' Fecha          : do., 12/may/2019 21:06:51
' Propósito      : Configura la vista para el juego
'------------------------------------------------------------------------------*
Public Function GetDisplaySorteo() As SorteoModel
    Dim oModel As SorteoModel
    Dim mPremios() As String
    
    Set oModel = New SorteoModel
    With oModel
        .IdSelected = mwsUserInterface.Range("C4").value
        .Juego = mwsUserInterface.Range("C5").value
        .NumSorteo = mwsUserInterface.Range("C6").value
        .FechaSorteo = mwsUserInterface.Range("C7").value
        .DiaSemana = mwsUserInterface.Range("C8").value
        .Semana = mwsUserInterface.Range("C9").value
        .OrdenAparicion = mwsUserInterface.Range("C10").value
        '
        '   TODO capturar los premios
        '
        Select Case .Juego
            Case LT_EUROMILLON
                .N1 = mwsUserInterface.Range("C12").value
                .N2 = mwsUserInterface.Range("D12").value
                .N3 = mwsUserInterface.Range("E12").value
                .N4 = mwsUserInterface.Range("F12").value
                .N5 = mwsUserInterface.Range("G12").value
                .E1 = mwsUserInterface.Range("C13").value
                .E2 = mwsUserInterface.Range("D13").value
                ReDim mPremios(12)
                mPremios(0) = CStr(mwsUserInterface.Range("C18").value)
                mPremios(1) = CStr(mwsUserInterface.Range("C19").value)
                mPremios(2) = CStr(mwsUserInterface.Range("C20").value)
                mPremios(3) = CStr(mwsUserInterface.Range("C21").value)
                mPremios(4) = CStr(mwsUserInterface.Range("C22").value)
                mPremios(5) = CStr(mwsUserInterface.Range("C23").value)
                mPremios(6) = CStr(mwsUserInterface.Range("C24").value)
                mPremios(7) = CStr(mwsUserInterface.Range("C25").value)
                mPremios(8) = CStr(mwsUserInterface.Range("C26").value)
                mPremios(9) = CStr(mwsUserInterface.Range("C27").value)
                mPremios(10) = CStr(mwsUserInterface.Range("C28").value)
                mPremios(11) = CStr(mwsUserInterface.Range("C29").value)
                mPremios(12) = CStr(mwsUserInterface.Range("C30").value)
                .MatrizPremios = mPremios
                
            Case LT_GORDO
                .N1 = mwsUserInterface.Range("C12").value
                .N2 = mwsUserInterface.Range("D12").value
                .N3 = mwsUserInterface.Range("E12").value
                .N4 = mwsUserInterface.Range("F12").value
                .N5 = mwsUserInterface.Range("G12").value
                .Reintegro = mwsUserInterface.Range("C14").value
                ReDim mPremios(8)
                mPremios(0) = CStr(mwsUserInterface.Range("C18").value)
                mPremios(1) = CStr(mwsUserInterface.Range("C19").value)
                mPremios(2) = CStr(mwsUserInterface.Range("C20").value)
                mPremios(3) = CStr(mwsUserInterface.Range("C21").value)
                mPremios(4) = CStr(mwsUserInterface.Range("C22").value)
                mPremios(5) = CStr(mwsUserInterface.Range("C23").value)
                mPremios(6) = CStr(mwsUserInterface.Range("C24").value)
                mPremios(7) = CStr(mwsUserInterface.Range("C25").value)
                mPremios(8) = CStr(mwsUserInterface.Range("C26").value)
                .MatrizPremios = mPremios
            
            Case Else
                .N1 = mwsUserInterface.Range("C12").value
                .N2 = mwsUserInterface.Range("D12").value
                .N3 = mwsUserInterface.Range("E12").value
                .N4 = mwsUserInterface.Range("F12").value
                .N5 = mwsUserInterface.Range("G12").value
                .N6 = mwsUserInterface.Range("H12").value
                .Complementario = mwsUserInterface.Range("C13").value
                .Reintegro = mwsUserInterface.Range("C14").value
                ReDim mPremios(5)
                mPremios(0) = CStr(mwsUserInterface.Range("C18").value)
                mPremios(1) = CStr(mwsUserInterface.Range("C19").value)
                mPremios(2) = CStr(mwsUserInterface.Range("C20").value)
                mPremios(3) = CStr(mwsUserInterface.Range("C21").value)
                mPremios(4) = CStr(mwsUserInterface.Range("C22").value)
                mPremios(5) = CStr(mwsUserInterface.Range("C23").value)
                .MatrizPremios = mPremios
        End Select
    End With
    Set GetDisplaySorteo = oModel
End Function

'--- Métodos Privados ---------------------------------------------------------*
Private Sub Class_Initialize()
    Set mwsUserInterface = ThisWorkbook.Sheets("Editar")
End Sub
Private Sub Class_Terminate()
    Set mwsUserInterface = Nothing
End Sub

'--- Métodos Públicos ---------------------------------------------------------*
'------------------------------------------------------------------------------*
' Procedimiento  : ClearSorteoDisplay
' Fecha          : do., 12/may/2019 21:06:51
' Propósito      : Limpia los datos del registro del formulario
'------------------------------------------------------------------------------*
Public Sub ClearSorteoDisplay(EsNuevo As Boolean)
    Application.ScreenUpdating = False
    '
    '   Inicializa areas de datos
    '
    If EsNuevo Then
        mwsUserInterface.Range("C6:C10").ClearContents
    Else
        mwsUserInterface.Range("C4:C10").ClearContents
    End If
    mwsUserInterface.Range("C12:H14").ClearContents
    mwsUserInterface.Range("B18:C30").ClearContents
    Application.ScreenUpdating = True
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : DisplaySorteoRecord
' Fecha          : do., 12/may/2019 21:06:51
' Propósito      : Visualiza las propiedades del Sorteo
'------------------------------------------------------------------------------*
Public Sub DisplaySorteoRecord(oModel As SorteoModel)
    Application.ScreenUpdating = False
    '
    '   Establecemos los literales para el juego
    '
    Me.SetDisplayJuego oModel.Juego
    '
    '   Movemos los datos a la Hoja
    '
    With mwsUserInterface
        .Range("C5").value = oModel.Juego
        .Range("C4").value = oModel.IdSelected
        .Range("C6").value = oModel.NumSorteo
        .Range("C7").value = oModel.FechaSorteo
        .Range("C8").value = oModel.DiaSemana
        .Range("C9").value = oModel.Semana
        .Range("C10").value = oModel.OrdenAparicion
        If oModel.Juego = LT_EUROMILLON Then
            .Range("C12").value = oModel.N1
            .Range("D12").value = oModel.N2
            .Range("E12").value = oModel.N3
            .Range("F12").value = oModel.N4
            .Range("G12").value = oModel.N5
            .Range("C13").value = oModel.E1
            .Range("D13").value = oModel.E2
            '
            '   TODO: Establecer los premios
            '
            .Range("C18").value = CCur(oModel.MatrizPremios(0))
            .Range("C19").value = CCur(oModel.MatrizPremios(1))
            .Range("C20").value = CCur(oModel.MatrizPremios(2))
            .Range("C21").value = CCur(oModel.MatrizPremios(3))
            .Range("C22").value = CCur(oModel.MatrizPremios(4))
            .Range("C23").value = CCur(oModel.MatrizPremios(5))
            .Range("C24").value = CCur(oModel.MatrizPremios(6))
            .Range("C25").value = CCur(oModel.MatrizPremios(7))
            .Range("C26").value = CCur(oModel.MatrizPremios(8))
            .Range("C27").value = CCur(oModel.MatrizPremios(9))
            .Range("C28").value = CCur(oModel.MatrizPremios(10))
            .Range("C29").value = CCur(oModel.MatrizPremios(11))
            .Range("C30").value = CCur(oModel.MatrizPremios(12))
        
        ElseIf oModel.Juego = LT_GORDO Then
            .Range("C12").value = oModel.N1
            .Range("D12").value = oModel.N2
            .Range("E12").value = oModel.N3
            .Range("F12").value = oModel.N4
            .Range("G12").value = oModel.N5
            .Range("H12").value = Empty
            .Range("B13").value = Empty
            .Range("B14").value = LT_CLAVE
            .Range("C14").value = oModel.Reintegro
            .Range("C18").value = CCur(oModel.MatrizPremios(0))
            .Range("C19").value = CCur(oModel.MatrizPremios(1))
            .Range("C20").value = CCur(oModel.MatrizPremios(2))
            .Range("C21").value = CCur(oModel.MatrizPremios(3))
            .Range("C22").value = CCur(oModel.MatrizPremios(4))
            .Range("C23").value = CCur(oModel.MatrizPremios(5))
            .Range("C24").value = CCur(oModel.MatrizPremios(6))
            .Range("C25").value = CCur(oModel.MatrizPremios(7))
            .Range("C26").value = CCur(oModel.MatrizPremios(8))
            
        ElseIf oModel.Juego = LT_PRIMITIVA Then
            .Range("C12").value = oModel.N1
            .Range("D12").value = oModel.N2
            .Range("E12").value = oModel.N3
            .Range("F12").value = oModel.N4
            .Range("G12").value = oModel.N5
            .Range("H12").value = oModel.N6
            .Range("B13").value = LT_COMPLEMENTARIO
            .Range("C13").value = oModel.Complementario
            .Range("D13").value = Empty
            .Range("B14").value = LT_REINTEGRO
            .Range("C14").value = oModel.Reintegro
            .Range("C18").value = CCur(oModel.MatrizPremios(0))
            .Range("C19").value = CCur(oModel.MatrizPremios(1))
            .Range("C20").value = CCur(oModel.MatrizPremios(2))
            .Range("C21").value = CCur(oModel.MatrizPremios(3))
            .Range("C22").value = CCur(oModel.MatrizPremios(4))
            .Range("C23").value = CCur(oModel.MatrizPremios(5))
            .Range("C24").value = CCur(oModel.MatrizPremios(6))
        
        ElseIf oModel.Juego = LT_BONOLOTO Then
            .Range("C12").value = oModel.N1
            .Range("D12").value = oModel.N2
            .Range("E12").value = oModel.N3
            .Range("F12").value = oModel.N4
            .Range("G12").value = oModel.N5
            .Range("H12").value = oModel.N6
            .Range("B13").value = LT_COMPLEMENTARIO
            .Range("C13").value = oModel.Complementario
            .Range("D13").value = Empty
            .Range("B14").value = LT_REINTEGRO
            .Range("C14").value = oModel.Reintegro
            .Range("C18").value = CCur(oModel.MatrizPremios(0))
            .Range("C19").value = CCur(oModel.MatrizPremios(1))
            .Range("C20").value = CCur(oModel.MatrizPremios(2))
            .Range("C21").value = CCur(oModel.MatrizPremios(3))
            .Range("C22").value = CCur(oModel.MatrizPremios(4))
            .Range("C23").value = CCur(oModel.MatrizPremios(5))
        End If
    End With
    Application.ScreenUpdating = True
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : SetDisplayJuego
' Fecha          : do., 12/may/2019 21:09:15
' Propósito      : Establece los literales según el juego
'------------------------------------------------------------------------------*
Public Sub SetDisplayJuego(ByVal mJuego As String)
    Application.ScreenUpdating = False
    '
    '   Inicializamos áreas
    '
    Me.ClearSorteoDisplay True
    '
    '   Inicializamos áreas de color
    '
    With mwsUserInterface.Range("C12:H14").Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    '
    '   Area de premios
    '
    With mwsUserInterface.Range("C18:C30").Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    '
    '
    '
    Select Case mJuego
        Case LT_EUROMILLON
            '
            '   Area de Combinación
            '
            With mwsUserInterface.Range("C12:G12")
                '
                '   Fondo
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorDark1
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
                '   Letra
                .Font.Name = "Calibri"
                .Font.Size = 18
                .Font.Strikethrough = False
                .Font.Superscript = False
                .Font.Subscript = False
                .Font.OutlineFont = False
                .Font.Shadow = False
                .Font.Underline = xlUnderlineStyleNone
                .Font.ThemeColor = xlThemeColorLight1
                .Font.TintAndShade = 0
                .Font.ThemeFont = xlThemeFontMinor
            End With
            '
            '   Area de Estrellas
            '
            With mwsUserInterface.Range("C13:D13")
                '
                '   Fondo
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorDark1
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
                '   Letra
                .Font.Name = "Calibri"
                .Font.Size = 18
                .Font.Strikethrough = False
                .Font.Superscript = False
                .Font.Subscript = False
                .Font.OutlineFont = False
                .Font.Shadow = False
                .Font.Underline = xlUnderlineStyleNone
                .Font.ThemeColor = xlThemeColorLight1
                .Font.TintAndShade = 0
                .Font.ThemeFont = xlThemeFontMinor
            End With
            '
            '   Area de premios
            '
            With mwsUserInterface
                .Range("B13").value = LT_ESTRELLAS
                .Range("B14").value = Empty
                .Range("B18").value = "1ª 5 + 2"
                .Range("B19").value = "2ª 5 + 1"
                .Range("B20").value = "3ª 5 + 0"
                .Range("B21").value = "4ª 4 + 2"
                .Range("B22").value = "5ª 4 + 1"
                .Range("B23").value = "6ª 3 + 2"
                .Range("B24").value = "7ª 4 + 0"
                .Range("B25").value = "8ª 2 + 2"
                .Range("B26").value = "9ª 3 + 1"
                .Range("B27").value = "10ª 3 + 0"
                .Range("B28").value = "11ª 1 + 2"
                .Range("B29").value = "12ª 2 + 1"
                .Range("B30").value = "13ª 2 + 0"
            End With
            With mwsUserInterface.Range("C18:C30").Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
'-------
        Case LT_GORDO
            '
            '   Area de Combinación
            '
            With mwsUserInterface.Range("C12:G12")
                '
                '   Fondo
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorDark1
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
                '   Letra
                .Font.Name = "Calibri"
                .Font.Size = 18
                .Font.Strikethrough = False
                .Font.Superscript = False
                .Font.Subscript = False
                .Font.OutlineFont = False
                .Font.Shadow = False
                .Font.Underline = xlUnderlineStyleNone
                .Font.ThemeColor = xlThemeColorLight1
                .Font.TintAndShade = 0
                .Font.ThemeFont = xlThemeFontMinor
            End With
            '
            '   Clave
            '
            With mwsUserInterface.Range("C14")
                '
                '   Fondo
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorDark1
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
            End With
            '
            '
            With mwsUserInterface
                .Range("B13").value = Empty
                .Range("B14").value = LT_CLAVE
                .Range("B18").value = "1ª (5 + 1)"
                .Range("B19").value = "2ª (5 + 0)"
                .Range("B20").value = "3ª (4 + 1)"
                .Range("B21").value = "4ª (4 + 0)"
                .Range("B22").value = "5ª (3 + 1)"
                .Range("B23").value = "6ª (3 + 0)"
                .Range("B24").value = "7ª (2 + 1)"
                .Range("B25").value = "8ª (2 + 0)"
                .Range("B26").value = "Reintegro"
            End With
            '
            '   Area de premios
            '
            With mwsUserInterface.Range("C18:C26").Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
'---------
        Case LT_PRIMITIVA
            '
            '   Area de Combinación
            '
            With mwsUserInterface.Range("C12:H12")
                '
                '   Fondo
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorDark1
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
                '   Letra
                .Font.Name = "Calibri"
                .Font.Size = 18
                .Font.Strikethrough = False
                .Font.Superscript = False
                .Font.Subscript = False
                .Font.OutlineFont = False
                .Font.Shadow = False
                .Font.Underline = xlUnderlineStyleNone
                .Font.ThemeColor = xlThemeColorLight1
                .Font.TintAndShade = 0
                .Font.ThemeFont = xlThemeFontMinor
            End With
            '
            '   Area de Complementario
            '
            With mwsUserInterface.Range("C13")
                '
                '   Fondo
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorDark1
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
                '   Letra
                .Font.Name = "Calibri"
                .Font.Size = 18
                .Font.Strikethrough = False
                .Font.Superscript = False
                .Font.Subscript = False
                .Font.OutlineFont = False
                .Font.Shadow = False
                .Font.Underline = xlUnderlineStyleNone
                .Font.ThemeColor = xlThemeColorLight1
                .Font.TintAndShade = 0
                .Font.ThemeFont = xlThemeFontMinor
            End With
            '
            '   Reintegro
            '
            With mwsUserInterface.Range("C14")
                '
                '   Fondo
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorDark1
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
            End With
            With mwsUserInterface
                .Range("B13").value = LT_COMPLEMENTARIO
                .Range("B14").value = LT_REINTEGRO
                .Range("B18").value = "Especial (6 Aciertos + R)"
                .Range("B19").value = "1ª (6 Aciertos)"
                .Range("B20").value = "2ª (5 Aciertos + C)"
                .Range("B21").value = "3ª (5 Aciertos)"
                .Range("B22").value = "4ª (4 Aciertos)"
                .Range("B23").value = "5ª (3 Aciertos)"
                .Range("B24").value = "Reintegro"
            End With
            '
            '   Area de premios
            '
            With mwsUserInterface.Range("C18:C24").Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
'---------
        Case LT_BONOLOTO
            '
            '   Area de Combinación
            '
            With mwsUserInterface.Range("C12:H12")
                '
                '   Fondo
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorDark1
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
                '   Letra
                .Font.Name = "Calibri"
                .Font.Size = 18
                .Font.Strikethrough = False
                .Font.Superscript = False
                .Font.Subscript = False
                .Font.OutlineFont = False
                .Font.Shadow = False
                .Font.Underline = xlUnderlineStyleNone
                .Font.ThemeColor = xlThemeColorLight1
                .Font.TintAndShade = 0
                .Font.ThemeFont = xlThemeFontMinor
            End With
            '
            '   Area de Complementario
            '
            With mwsUserInterface.Range("C13")
                '
                '   Fondo
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorDark1
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
                '   Letra
                .Font.Name = "Calibri"
                .Font.Size = 18
                .Font.Strikethrough = False
                .Font.Superscript = False
                .Font.Subscript = False
                .Font.OutlineFont = False
                .Font.Shadow = False
                .Font.Underline = xlUnderlineStyleNone
                .Font.ThemeColor = xlThemeColorLight1
                .Font.TintAndShade = 0
                .Font.ThemeFont = xlThemeFontMinor
            End With
            '
            '   Reintegro
            '
            With mwsUserInterface.Range("C14")
                '
                '   Fondo
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.ThemeColor = xlThemeColorDark1
                .Interior.TintAndShade = 0
                .Interior.PatternTintAndShade = 0
            End With
            With mwsUserInterface
                .Range("B13").value = LT_COMPLEMENTARIO
                .Range("B14").value = LT_REINTEGRO
                .Range("B18").value = "1ª (6 Aciertos)"
                .Range("B19").value = "2ª (5 Aciertos + C)"
                .Range("B20").value = "3ª (5 Aciertos)"
                .Range("B21").value = "4ª (4 Aciertos)"
                .Range("B22").value = "5ª (3 Aciertos)"
                .Range("B23").value = "Reintegro"
            End With
            '
            '   Area de premios
            '
            With mwsUserInterface.Range("C18:C23").Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

    End Select
    Application.ScreenUpdating = True
End Sub

'' *===========(EOF): SorteoEditarView.cls

