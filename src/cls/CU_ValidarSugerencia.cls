VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CU_ValidarSugerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : GetAciertos
' Author    : Charly
' Date      : 15/04/2012
' Purpose   : Devuelve el Numero de aciertos de una sugerencia
'---------------------------------------------------------------------------------------
'
Public Function GetAciertos(datSugerencia As Sugerencia, datSorteo As Sorteo) As Integer
    Dim m_iAciertos         As Integer  ' Numero de aciertos
    Dim m_oNum              As Numero   ' Numero de control
  
  On Error GoTo GetAciertos_Error
    '
    '   Se inicializa los aciertos
    '
    m_iAciertos = 0
    '
    '   Para cada Numero en la combinación
    '
    For Each m_oNum In datSugerencia.Combinacion.Numeros
        '
        '   Se comprueba si está en la combinación del sorteo
        '
        If (datSorteo.Combinacion.Contiene(m_oNum.Valor)) Then
            '
            '  Se incrementa el contador
            '
            m_iAciertos = m_iAciertos + 1
        End If
    Next m_oNum
    '
    '   Devuelve el Numero de aciertos
    '
    GetAciertos = m_iAciertos

   On Error GoTo 0
   Exit Function

GetAciertos_Error:
     Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "CU_ValidarSugerencia.GetAciertos")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

Public Function GetPremio(datSugerencia As Sugerencia, datSorteo As Sorteo) As String
    Dim m_iAciertos     As Integer      ' Numero de aciertos
    Dim m_bComple       As Integer      ' Un acierto es del complementario
    Dim m_iPronosticos  As Integer
    Dim m_sPremio       As String
    
    m_sPremio = ""
    m_iAciertos = GetAciertos(datSugerencia, datSorteo)
    m_bComple = EsComplementario(datSugerencia, datSorteo)
    If m_bComple Then
        m_iAciertos = m_iAciertos - 1
    End If
    m_iPronosticos = datSugerencia.Combinacion.Count
    '
    '   COmprobar el juego del sorteo
    '
    If m_iPronosticos = 5 Then
        Select Case m_iAciertos
            Case 2: m_sPremio = "5ª"
            Case 3: m_sPremio = "4ª"
            Case 4:
                If (m_bComple) Then
                    m_sPremio = "3ª"
                Else
                    m_sPremio = "2ª"
                End If
            Case 5: m_sPremio = "1ª"
        End Select
    Else
        Select Case m_iAciertos
            Case 3: m_sPremio = "5ª"
            Case 4: m_sPremio = "4ª"
            Case 5:
                If (m_bComple) Then
                    m_sPremio = "3ª"
                Else
                    m_sPremio = "2ª"
                End If
            Case 6: m_sPremio = "1ª"
        End Select
    End If
    '
    '   Devuelve el premio
    '
    GetPremio = m_sPremio
End Function


Private Function EsComplementario(datSugerencia As Sugerencia, datSorteo As Sorteo) As Boolean
    Dim m_oNum              As Numero   ' Numero de control
    EsComplementario = False
    For Each m_oNum In datSugerencia.Combinacion.Numeros
        '
        '   Se comprueba si está en la combinación del sorteo
        '
        If (datSorteo.Complementario = m_oNum.Valor) Then
            EsComplementario = True
            Exit Function
        End If
    Next m_oNum

End Function
