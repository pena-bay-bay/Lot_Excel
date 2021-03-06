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
' *     Creaci�n   : 22/jul/2017
' *     Versi�n    : 1.0
' *     Prop�sito  : Proporciona informaci�n de fechas de cada sorteo
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
    mFInicioEuroMillon = #2/13/2004#
    mJuego = JUEGO_DEFECTO
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : Constructor
' Fecha          : 22/jul/2017
' Prop�sito      : Inicializar la clase con parametros
' Par�metros     : Juego
'------------------------------------------------------------------------------*
'
Public Sub Constructor(vNewJuego As Juego)
    mJuego = vNewJuego
End Sub
'------------------------------------------------------------------------------*
' Procedimiento  : GetProximoSorteo
' Fecha          : 22/jul/2017
' Prop�sito      : Devuelve la fecha del pr�ximo Sorteo
' Par�metros     : Fecha de an�lisis y juego
' Retorno        : Fecha del pr�ximo sorteo a la fecha analizada
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
    ' Asume la fecha de pase como fecha del pr�ximo sorteo
    '
    GetProximoSorteo = vNewDate
    '
    ' Evalua el juego y calcula una fecha en funci�n de cada juego
    '
    Select Case mJuego
        '
        '
        Case Juego.Bonoloto
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
                    Case 6: 'V
                        GetProximoSorteo = GetProximoSorteo + 3
                    Case 4, 7:  '  X y S
                        GetProximoSorteo = GetProximoSorteo + 2
                    Case Else:
                        GetProximoSorteo = GetProximoSorteo + 1
                End Select
            End If
            
        Case Juego.Euromillones
            mDias = 2       ' M, V
            Select Case mDiaSemana
                Case 2, 5:  ' L J
                    GetProximoSorteo = GetProximoSorteo + 1
                Case 1, 4:  ' D X
                    GetProximoSorteo = GetProximoSorteo + 2
                Case 7, 3:  ' S M
                    GetProximoSorteo = GetProximoSorteo + 3
                Case 6:     ' V
                    GetProximoSorteo = GetProximoSorteo + 4
            End Select
            
        Case Juego.GordoPrimitiva
            mDias = 1       ' D
            If mDiaSemana > 1 Then
                GetProximoSorteo = GetProximoSorteo + (7 - mDiaSemana) + 1
            Else
                GetProximoSorteo = GetProximoSorteo + 7
            End If
            
        Case Juego.LoteriaPrimitiva
            mDias = 2
            Select Case mDiaSemana
                Case 1:     ' D
                    GetProximoSorteo = GetProximoSorteo + 4
                Case 2:     ' L
                    GetProximoSorteo = GetProximoSorteo + 3
                Case 3, 5:  ' M y J
                    GetProximoSorteo = GetProximoSorteo + 2
                Case 4, 6:  ' X y V
                    GetProximoSorteo = GetProximoSorteo + 1
                Case 7:     ' S
                    GetProximoSorteo = GetProximoSorteo + 5
            End Select
    End Select

End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetSorteosEntreFechas
' Fecha          : 22/jul/2017
' Prop�sito      : Devuelve el n�mero de sorteos entre dos fechas
' Par�metros     :
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
    ' Inicializa el n�mero de sorteos
    '
    mSorteos = 0
    '
    ' Para cada fecha en el intervalo analiza la celebraci�n del sorteo
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
            Case Juego.Bonoloto:
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
                
            Case Juego.GordoPrimitiva:
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
' Prop�sito      : Comprueba que la fecha corresponde a un sorteo
' Par�metros     : Fecha a comprobar
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
            Case Juego.Bonoloto:
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
                
            Case Juego.GordoPrimitiva:
                If mDiaSem = 1 Then
                    EsFechaSorteo = True
                End If
    End Select
    
End Function
'------------------------------------------------------------------------------*
' Procedimiento  : GetAnteriorSorteo
' Fecha          : 11/02/2018 19:44
' Prop�sito      : Devuelve la fecha del anterior Sorteo
' Par�metros     : Fecha de an�lisis
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
    ' Evalua el juego y calcula una fecha en funci�n de cada juego
    '
    Select Case mJuego
        '
        '
        Case Juego.Bonoloto
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
                Case 4, 7:     ' X y S
                    GetAnteriorSorteo = GetAnteriorSorteo - 1
                Case 1, 5:     ' D y J
                    GetAnteriorSorteo = GetAnteriorSorteo - 2
                Case 2, 6:     ' L y V
                    GetAnteriorSorteo = GetAnteriorSorteo - 3
                Case 3:     ' M
                    GetAnteriorSorteo = GetAnteriorSorteo - 4
            End Select
            
        Case Juego.GordoPrimitiva
            mDias = 1       ' D
            If mDiaSemana > 1 Then
                GetAnteriorSorteo = (GetAnteriorSorteo - mDiaSemana) + 1
            Else
                GetAnteriorSorteo = GetAnteriorSorteo - 7
            End If
            
        Case Juego.LoteriaPrimitiva
            mDias = 2
             Select Case mDiaSemana
                    Case 1, 6: ' D y V
                        GetAnteriorSorteo = GetAnteriorSorteo - 1
                    Case 2, 7: ' L y S
                        GetAnteriorSorteo = GetAnteriorSorteo - 2
                    Case 3: ' M
                        GetAnteriorSorteo = GetAnteriorSorteo - 3
                    Case 4: ' X
                        GetAnteriorSorteo = GetAnteriorSorteo - 4
                    Case 5: ' J
                        GetAnteriorSorteo = GetAnteriorSorteo - 5
                End Select
    End Select
End Function

'------------------------------------------------------------------------------*
' Procedimiento  : AddDiasSorteo
' Fecha          : 24/01/2020 09:59
' Prop�sito      : Devuelve la fecha resultante de sumar ndias a un sorteo
' Par�metros     : Fecha de an�lisis, Numero dias
' Retorno        : Fecha correspondiente a la suma de sorteos a una fecha
'------------------------------------------------------------------------------*
'
Public Function AddDiasSorteo(vNewDate As Date, nNewDias As Integer) As Date
    Dim i     As Integer
    Dim mDate As Date
    
    mDate = vNewDate
    For i = 1 To nNewDias
        mDate = Me.GetProximoSorteo(mDate)
    Next i
'
'   #TODO: Optimizar con un algoritmo de calculo segun el juego
'
'    '
'    '  1 = Domingo
'    '  D,L,M,X,J,V,S
'    '  1,2,3,4,5,6,7
'    '
'    mDiaSemana = Weekday(vNewDate)
'    '
'    '   Cuantas semanas hay en los dias suministrados
'    '   depende del juego, si es bonoloto hay 6 juegos por semana
'
'
'    mDate = vNewDate
'    For i = 0 To nNewDias
'        mDate = mDate + i
'
'    Next i
'
    AddDiasSorteo = mDate
End Function


'------------------------------------------------------------------------------*
' Procedimiento  : SubstractDiasSorteo
' Fecha          : ma., 07/jul/2020 16:20:31
' Prop�sito      : Devuelve la fecha resultante de restar ndias a un sorteo
' Par�metros     : Fecha de an�lisis, Numero dias
' Retorno        : Fecha correspondiente a la resta de sorteos a una fecha
'------------------------------------------------------------------------------*
'
Public Function SubstractDiasSorteo(vNewDate As Date, nNewDias As Integer) As Date
    Dim i     As Integer
    Dim mDate As Date
    
    mDate = vNewDate
    For i = 1 To nNewDias
        mDate = Me.GetAnteriorSorteo(mDate)
    Next i
'
'   #TODO: Optimizar con un algoritmo de calculo segun el juego
'
'    '
'    '  1 = Domingo
'    '  D,L,M,X,J,V,S
'    '  1,2,3,4,5,6,7
'    '
'    mDiaSemana = Weekday(vNewDate)
'    '
'    '   Cuantas semanas hay en los dias suministrados
'    '   depende del juego, si es bonoloto hay 6 juegos por semana
'
'
'    mDate = vNewDate
'    For i = 0 To nNewDias
'        mDate = mDate + i
'
'    Next i
'
    SubstractDiasSorteo = mDate
End Function
