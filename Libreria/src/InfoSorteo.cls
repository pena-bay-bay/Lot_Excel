VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfoSorteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *============================================================================*
' *
' *     Fichero    : InfoSorteo.cls
' *
' *     Autor      : Carlos Almela Baeza
' *     Creación   : 22/jul/2017
' *     Versión    : 1.0
' *     Propósito  : Proporciona información de fechas de cada sorteo
' *
' *============================================================================*
Option Explicit
Option Base 0
Private mFInicioBonoloto As Date
Private mFInicioPrimitiva As Date
Private mFInicioEuroMillon As Date
Private mFecCambioBonoloto As Date
Private mJuego As Juego
'------------------------------------------------------------------------------*
' Procedimiento  : Inicializacion
' Fecha          : 22/jul/2017
'------------------------------------------------------------------------------*
'
Private Sub Class_Initialize()
    mFInicioBonoloto = #10/5/1999#
    mFecCambioBonoloto = #3/6/2015#
    mJuego = JUEGO_DEFECTO
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Constructor
' Fecha          : 22/jul/2017
' Propósito      : Inicializar la clase con parametros
' Parámetros     : Juego
'------------------------------------------------------------------------------*
'
Public Sub Constructor(vNewJuego As Juego)
    mJuego = vNewJuego
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetProximoSorteo
' Fecha          : 22/jul/2017
' Propósito      : Devuelve la fecha del próximo Sorteo
' Parámetros     : Fecha de análisis y juego
' Retorno        : Fecha del próximo sorteo a la fecha analizada
'------------------------------------------------------------------------------*
'
Public Function GetProximoSorteo(vNewDate As Date) As Date
    Dim mDias As Integer
    Dim mDiaSemana As Integer
    '
    '  1 = Domingo
    '  D,L,M,X,J,V,S
    '  1,2,3,4,5,6,7
    '
    mDiaSemana = Weekday(vNewDate)
    '
    ' Asume la fecha de pase como fecha del próximo sorteo
    '
    GetProximoSorteo = vNewDate
    '
    ' Evalua el juego y calcula una fecha en función de cada juego
    '
    Select Case mJuego
        '
        '
        Case Juego.bonoloto
            '
            ' Bonoloto tiene dos versiones de juego
            '
            If vNewDate > mFecCambioBonoloto Then
                mDias = 6   ' L,M,X,J,V,S
                If mDiaSemana = 7 Then            ' S
                    GetProximoSorteo = GetProximoSorteo + 2
                Else
                    GetProximoSorteo = GetProximoSorteo + 1
                End If
            Else
                mDias = 4   ' L,M,X,V
                Select Case mDiaSemana
                    Case 5:  ' J
                        GetProximoSorteo = GetProximoSorteo + 1
                    Case 7:  ' S
                        GetProximoSorteo = GetProximoSorteo + 2
                    Case 1:  ' D
                        GetProximoSorteo = GetProximoSorteo + 1
                End Select
            End If
            
        Case Juego.Euromillones
            mDias = 2       ' M, V
            Select Case mDiaSemana
                Case 2, 5:  ' L J
                    GetProximoSorteo = GetProximoSorteo + 1
                Case 7:     ' S
                    GetProximoSorteo = GetProximoSorteo + 3
                Case 1, 4:  ' D, X
                    GetProximoSorteo = GetProximoSorteo + 2
            End Select
            
        Case Juego.gordoPrimitiva
            mDias = 1       ' D
            If mDiaSemana > 1 Then
                GetProximoSorteo = GetProximoSorteo + (7 - mDiaSemana)
            End If
            
        Case Juego.LoteriaPrimitiva
            mDias = 2
             Select Case mDiaSemana
                    Case 4, 6:  '  X y V
                        GetProximoSorteo = GetProximoSorteo + 1
                    Case 2:     ' L
                        GetProximoSorteo = GetProximoSorteo + 3
                    Case 1:     ' D
                        GetProximoSorteo = GetProximoSorteo + 4
                    Case 3:     ' M
                        GetProximoSorteo = GetProximoSorteo + 2
                End Select
    End Select

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetSorteosEntreFechas
' Fecha          : 22/jul/2017
' Propósito      : Devuelve el número de sorteos entre dos fechas
' Parámetros     :
' Retorno        :
'------------------------------------------------------------------------------*
'
Public Function GetSorteosEntreFechas(vNewDateIni As Date, vNewDateFin As Date) As Integer
    Dim mSorteos              As Integer
    Dim mDiaSem               As Integer
    Dim mCurrentFecha         As Date
    '
    ' Si la fecha inicial es superior a la final devuelve 0 sorteos
    '
    If vNewDateIni > vNewDateFin Then
        GetSorteosEntreFechas = 0
        Exit Function
    End If
    '
    ' Si la fecha inicial es igual a la final devuelve 0 sorteos
    '
    If (vNewDateFin - vNewDateIni) = 0 Then
        GetSorteosEntreFechas = 0
        Exit Function
    End If
    '
    ' Inicializa el número de sorteos
    '
    mSorteos = 0
    '
    ' Para cada fecha en el intervalo analiza la celebración del sorteo
    '
    For mCurrentFecha = vNewDateIni To vNewDateFin - 1 Step 1
        '
        ' Calcula el dia de la semana 1 = Domingo
        '
        mDiaSem = Weekday(mCurrentFecha)
        '
        ' Analiza el juego del sorteo
        '
        Select Case mJuego
            Case Juego.bonoloto:
                '
                ' Si es bonoloto y no es domingo suma 1 sorteo
                '
                If mDiaSem > 1 Then
                    mSorteos = mSorteos + 1
                End If
                
            Case Juego.LoteriaPrimitiva
                '
                ' Si es primitiva y es jueves o sabado suma 1 sorteo
                '
                If (mDiaSem = 5) Or (mDiaSem = 7) Then
                    mSorteos = mSorteos + 1
                End If
                
            Case Juego.Euromillones:
                '
                ' Si es euromillon y es martes o viernes suma 1 sorteo
                '
                If (mDiaSem = 3) Or (mDiaSem = 6) Then
                    mSorteos = mSorteos + 1
                End If
                
            Case Juego.gordoPrimitiva:
                '
                ' Si es el Gordo de la primitiva y es domingo suma 1 sorteo
                '
                If mDiaSem = 1 Then
                    mSorteos = mSorteos + 1
                End If
        End Select
    Next
    '
    ' Devuelve el Numero de sorteos
    '
    GetSorteosEntreFechas = mSorteos
    
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : EsFechaSorteo
' Fecha          : 22/jul/2017
' Propósito      : Comprueba que la fecha corresponde a un sorteo
' Parámetros     : Fecha a comprobar
' Retorno        : True si es una fecha de sorteo False si no lo es
'------------------------------------------------------------------------------*
'
Public Function EsFechaSorteo(vNewDate As Date) As Boolean
    Dim mDiaSem               As Integer
    '
    '  1 = Domingo
    '  D,L,M,X,J,V,S
    '  1,2,3,4,5,6,7
    '
    EsFechaSorteo = False
    mDiaSem = Weekday(vNewDate)
    
    Select Case mJuego
            Case Juego.bonoloto:
                If mDiaSem <> 1 Then               ' Si es distinto a Domingo es fecha de sorteo
                   EsFechaSorteo = True
                End If
                
            Case Juego.LoteriaPrimitiva
                If (mDiaSem = 5) Or (mDiaSem = 7) Then
                     EsFechaSorteo = True
                End If
                
            Case Juego.Euromillones:
                If (mDiaSem = 3) Or (mDiaSem = 6) Then
                    EsFechaSorteo = True
                End If
                
            Case Juego.gordoPrimitiva:
                If mDiaSem = 1 Then
                    EsFechaSorteo = True
                End If
    End Select
    
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetAnteriorSorteo
' Fecha          : 11/02/2018 19:44
' Propósito      : Devuelve la fecha del anterior Sorteo
' Parámetros     : Fecha de análisis
' Retorno        : Fecha del anterior sorteo a la fecha analizada
'------------------------------------------------------------------------------*
'
Public Function GetAnteriorSorteo(vNewDate As Date) As Date
    Dim mDias As Integer
    Dim mDiaSemana As Integer
    '
    '  1 = Domingo
    '  D,L,M,X,J,V,S
    '  1,2,3,4,5,6,7
    '
    mDiaSemana = Weekday(vNewDate)
    '
    ' Asume la fecha de pase como fecha del anterior sorteo
    '
    GetAnteriorSorteo = vNewDate
    '
    ' Evalua el juego y calcula una fecha en función de cada juego
    '
    Select Case mJuego
        '
        '
        Case Juego.bonoloto
            '
            ' Bonoloto tiene dos versiones de juego
            '
            If vNewDate > mFecCambioBonoloto Then
                mDias = 6   ' L,M,X,J,V,S
                If mDiaSemana = 2 Then            ' L
                    GetAnteriorSorteo = GetAnteriorSorteo - 2
                Else
                    GetAnteriorSorteo = GetAnteriorSorteo - 1
                End If
            Else
                mDias = 4   ' L,M,X,V
                Select Case mDiaSemana
                    Case 6, 1:  ' D y V
                        GetAnteriorSorteo = GetAnteriorSorteo - 2
                    Case 2:  ' L
                        GetAnteriorSorteo = GetAnteriorSorteo - 3
                    Case Else
                        GetAnteriorSorteo = GetAnteriorSorteo - 1
                End Select
            End If
            
        Case Juego.Euromillones
            mDias = 2       ' M, V
            Select Case mDiaSemana
                Case 2, 5:  ' L J
                    GetAnteriorSorteo = GetAnteriorSorteo - 1
                Case 7:     ' S
                    GetAnteriorSorteo = GetAnteriorSorteo - 3
                Case 1, 4:  ' D, X
                    GetAnteriorSorteo = GetAnteriorSorteo - 2
            End Select
            
        Case Juego.gordoPrimitiva
            mDias = 1       ' D
            If mDiaSemana > 1 Then
                GetAnteriorSorteo = GetAnteriorSorteo - (7 - mDiaSemana)
            End If
            
        Case Juego.LoteriaPrimitiva
            mDias = 2
             Select Case mDiaSemana
                    Case 4, 6:  '  X y V
                        GetAnteriorSorteo = GetAnteriorSorteo - 1
                    Case 2:     ' L
                        GetAnteriorSorteo = GetAnteriorSorteo - 3
                    Case 1:     ' D
                        GetAnteriorSorteo = GetAnteriorSorteo - 4
                    Case 3:     ' M
                        GetAnteriorSorteo = GetAnteriorSorteo - 2
                End Select
    End Select
End Function
