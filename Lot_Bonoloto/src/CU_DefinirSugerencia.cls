VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CU_DefinirSugerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : GetSugerencia
' Author    : Charly
' Date      : 15/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetSugerencia(datMetodo As metodo, datMuestra As Muestra)
    Dim oSug        As New Sugerencia
    Dim oOrden      As TipoOrdenacion
    Dim oAgrup      As TipoAgrupacion
    Dim oBola       As bola
    Dim bAscendente As Boolean
    Dim vArray      As Variant
    Dim vArray2     As Variant
    'Dim Onum        As Numero
    Dim m_inum      As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim iPronosticos As Integer
    
  On Error GoTo GetSugerencia_Error

    i = 0
    j = 0
    '
    '   Obtenemos los parametros del metodo
    '
    With datMetodo.Parametros
        oOrden = .CriteriosOrdenacion
        oAgrup = .CriteriosAgrupacion
        bAscendente = .SentidoOrdenacion
        iPronosticos = .Pronosticos
    End With
    '
    '   Obtenemos la matriz de Numeros ordenada por el valor
    '
    vArray = datMuestra.GetMatrizBolas(oOrden, bAscendente)
    '
    '
    '
    If oAgrup <> grpSinDefinir Then
        '
        ' Atender a la agrupación
        '   Crear matriz de 2 *2
        '
        vArray2 = AgrupaMatriz(vArray, oAgrup)
        '
        '
        '
        Do Until oSug.Bolas.Count = iPronosticos
            '
            '   Obtiene el Numero de la matriz
            '
            m_inum = vArray2(i, j)
            '
            '   Si el  Numero no existe
            '
            If Not oSug.Combinacion.Contiene(m_inum) Then
                Set oBola = datMuestra.Get_Bola(m_inum)
                oSug.Combinacion.Add oBola.Numero
                oSug.Bolas.Add oBola
            End If
            '
            '   Cambia de grupo
            '
            If i >= UBound(vArray2, 1) Then
                i = 0
                j = j + 1
            Else
                i = i + 1
            End If
        Loop
        
    Else
        Do Until oSug.Bolas.Count = iPronosticos
            '
            '   Obtiene el Numero de la matriz
            '
            m_inum = vArray(i, 0)
            '
            '   Si el  Numero no existe
            '
            If Not oSug.Combinacion.Contiene(m_inum) Then
                Set oBola = datMuestra.Get_Bola(m_inum)
                oSug.Combinacion.Add oBola.Numero
                oSug.Bolas.Add oBola
            End If
            i = i + 1
        Loop
    End If
    '
    '   Inicializa atributos del metodo
    '
    Set oSug.metodo = datMetodo
    Set oSug.Parametros = datMuestra.ParametrosMuestra
    Set GetSugerencia = oSug

   On Error GoTo 0
   Exit Function

GetSugerencia_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "CU_DefinirSugerencia.GetSugerencia")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function


'---------------------------------------------------------------------------------------
' Procedure : AgrupaMatriz
' Author    : Charly
' Date      : 15/04/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function AgrupaMatriz(vMat As Variant, vAgrup As TipoAgrupacion) As Variant
    Dim a_Matriz()          As Variant
    Dim n                   As New Numero
    Dim i                   As Integer
    Dim a_cont()            As Variant
    Dim j                   As Integer
    
  On Error GoTo AgrupaMatriz_Error

    Select Case vAgrup
        
        Case grpDecenas:
            ReDim a_Matriz(4, 10)
            ReDim a_cont(4, 1)
            
        Case grpSeptenas:
            ReDim a_Matriz(7, 7)
            ReDim a_cont(7, 1)
            
        Case grpParidad:
            ReDim a_Matriz(1, 25)
            ReDim a_cont(1, 1)
            
        Case grpPeso:
            ReDim a_Matriz(1, 25)
            ReDim a_cont(1, 1)
            
        Case grpTerminacion:
            ReDim a_Matriz(9, 4)
            ReDim a_cont(9, 1)

    End Select
    
    '
    '   Redimimensionar contador y Probabilidad
    '
    
    For i = 0 To UBound(vMat)
        '
        '   Concretamos el Numero
        '
        n.Valor = vMat(i, 0)
        '
        '
        '
        Select Case vAgrup
            Case grpDecenas:
            
                j = n.Decena - 1
                a_Matriz(j, a_cont(j, 0)) = n.Valor
                a_cont(j, 0) = a_cont(j, 0) + 1
                a_cont(j, 1) = a_cont(j, 1) + vMat(i, 1)
                
            Case grpSeptenas:
                j = n.Septena - 1
                a_Matriz(j, a_cont(j, 0)) = n.Valor
                a_cont(j, 0) = a_cont(j, 0) + 1
                a_cont(j, 1) = a_cont(j, 1) + vMat(i, 1)
    
            Case grpParidad:
                j = IIf(n.Paridad = LT_PAR, 0, 1)
                a_Matriz(j, a_cont(j, 0)) = n.Valor
                a_cont(j, 0) = a_cont(j, 0) + 1
                a_cont(j, 1) = a_cont(j, 1) + vMat(i, 1)
    
            Case grpPeso:
                j = IIf(n.Peso = LT_BAJO, 0, 1)
                a_Matriz(j, a_cont(j, 0)) = n.Valor
                a_cont(j, 0) = a_cont(j, 0) + 1
                a_cont(j, 1) = a_cont(j, 1) + vMat(i, 1)
    
            Case grpTerminacion:
                j = n.Terminacion
                a_Matriz(j, a_cont(j, 0)) = n.Valor
                a_cont(j, 0) = a_cont(j, 0) + 1
                a_cont(j, 1) = a_cont(j, 1) + vMat(i, 1)
        
        End Select
    Next i
    
    AgrupaMatriz = a_Matriz

   On Error GoTo 0
   Exit Function

AgrupaMatriz_Error:
   Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
   ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
   '   Audita el error
   Call HandleException(ErrNumber, ErrDescription, "CU_DefinirSugerencia.AgrupaMatriz")
   '   Lanza el Error
   Err.Raise ErrNumber, ErrSource, ErrDescription
End Function

