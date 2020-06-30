Attribute VB_Name = "Lot_fncollections"
'---------------------------------------------------------------------------------------
' Modulo    : fn_collections
' Creado    : 25/05/2004 23:04
' Autor     : Carlos Almela Baeza
' Version   : 1.0.0 08-dic-2006 20:55
' Objeto    : Módulo de manejo de colecciones y arrays
'---------------------------------------------------------------------------------------
'
Option Explicit

'---------------------------------------------------------------------------------------
' Procedimiento : ExistenElem
' Creación      : 12-nov-2006 00:38
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Function ExistenElem(col As Collection, Clave As String) As Boolean
    Dim prueba As Variant
    On Error Resume Next
    prueba = col.Item(Clave)
    ExistenElem = (Err <> 5)
End Function
'---------------------------------------------------------------------------------------
' Procedimiento : EliminarTodosElementos
' Creación      : 12-nov-2006 00:38
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Sub EliminarTodosElementos(col As Collection)
    Do While col.Count
        col.Remove 1
    Loop
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : SustituirElem
' Creación      : 12-nov-2006 00:37
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
Public Sub SustituirElem(col As Collection, indice As Variant, nuevoValor As Variant)
    col.Remove indice
    If VarType(indice) = vbString Then
        col.Add nuevoValor, indice
    Else
        If indice > col.Count Then
            col.Add nuevoValor, , col.Count
        Else
            col.Add nuevoValor, , indice
        End If
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedimiento : ShellSortAny
' Creación      : 12-nov-2006 00:37
' Autor         : Carlos Almela Baeza
' Objeto        :
'---------------------------------------------------------------------------------------
'
'Public Sub ShellSortAny(arr As Variant, numEls As Long, descendente As Boolean)
'    Dim indice As Long, indice2 As Long, primerElem As Long
'    Dim distancia As Long, Valor As Variant
'
'   On Error GoTo ShellSortAny_Error
'
'    ' salir si no es un array
'    If VarType(arr) < vbArray Then Exit Sub
'
'    primerElem = LBound(arr)
'
'    ' encontrar el mejor valor para distancia
'    Do
'        distancia = distancia * 3 + 1
'    Loop Until distancia > numEls
'
'    ' ordenar el array
'    Do
'        distancia = distancia / 3
'        For indice = distancia + primerElem To numEls + primerElem - 1
'            Valor = arr(indice)
'            indice2 = indice
'            Do While (arr(indice2 - distancia) > Valor) Xor descendente
'                arr(indice2) = arr(indice2 - distancia)
'                indice2 = indice2 - distancia
'                If indice2 - distancia < primerElem Then Exit Do
'            Loop
'            arr(indice2) = Valor
'        Next
'    Loop Until distancia = 3
'
'   On Error GoTo 0
'   Exit Sub
'
'ShellSortAny_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & _
'        ") in procedure ShellSortAny of Módulo fn_collections"
'End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Ordenar
' Creación      : 30-Oct-2002 00:40
' Autor         : Carlos Almela Baeza
' Objeto        : Ordenar una matriz de una sola dimensión
'---------------------------------------------------------------------------------------
'
Public Sub Ordenar(ByRef matriz As Variant, Ascendente As Boolean)
    Dim tsOrdenado  As Boolean          'indicador de matriz ordenada
    Dim MxLimite    As Integer          'Limite máximo del bucle
    Dim TmpDato     As Variant          'Dato temporal para el intercambio
    Dim i           As Integer          'Indice
    
  On Error GoTo Ordenar_Error:
  
    'Si no es un elemento array sale de la rutina
    If Not IsArray(matriz) Then Exit Sub
    
    'obtiene el número máximo de elementos de la matriz
    MxLimite = UBound(matriz)
    
    'Bucle de ordenación, se realiza hasta que esté ordenada
    Do
        tsOrdenado = True               'Se parte de matriz ordenada
        For i = 0 To MxLimite - 1       'Se revisa cada elemento con el siguiente
            If Ascendente Then
                If matriz(i) > matriz(i + 1) _
                And (matriz(i + 1) <> 0) Then
                    tsOrdenado = False
                    TmpDato = matriz(i)         'Guardamos la posicion iesima
                    matriz(i) = matriz(i + 1)   'pasamos la posision siguiente a la
                                                'iesima
                    matriz(i + 1) = TmpDato     'pasamos el dato guardado a la siguiente
                End If
            Else
                If matriz(i) < matriz(i + 1) _
                And (matriz(i + 1) <> 0) Then
                    tsOrdenado = False
                    TmpDato = matriz(i)         'Guardamos la posicion iesima
                    matriz(i) = matriz(i + 1)   'pasamos la posision siguiente a la
                                                'iesima
                    matriz(i + 1) = TmpDato     'pasamos el dato guardado a la siguiente
                End If
            End If
        Next i
    Loop Until tsOrdenado

  On Error GoTo 0
    Exit Sub
Ordenar_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_fncollections.Ordenar", ErrSource)
    Err.Raise ErrNumber, "Lot_fncollections.Ordenar", ErrDescription
End Sub



'------------------------------------------------------------------------------*
'Función     : Ordenar2
'Fecha       : 28-Nov-1999
'Parametros  : Matriz de dos dimensiones
'Descripción : Ordena la matriz de paso
'                     x ->
'              matriz (0, 0) (0, 1)
'             y  |    (1, 0) (1, 1)
'                v    (2, 0) (2, 1)
'------------------------------------------------------------------------------*
Public Sub Ordenar2(ByRef matriz As Variant, _
                    Optional columna As Integer = 2, _
                    Optional Ascendente As Boolean = True)
                    
    Dim tsOrdenado          As Boolean
    Dim limiteY             As Integer
    Dim limiteX             As Integer
    Dim TmpDato()           As Variant
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    
   On Error GoTo Ordenar2_Error

    limiteY = UBound(matriz, 1)
    limiteX = UBound(matriz, 2)
    ReDim TmpDato(limiteX)
    
    j = IIf(columna = 1, 0, 1)
    
    If Not IsArray(matriz) Then Exit Sub
    
    Do
        tsOrdenado = True
        For i = 0 To limiteY - 1
            If Ascendente Then
                If (matriz(i, j) > matriz(i + 1, j)) Then
                    tsOrdenado = False
                    For k = 0 To limiteX
                        TmpDato(k) = matriz(i, k)
                        matriz(i, k) = matriz(i + 1, k)
                        matriz(i + 1, k) = TmpDato(k)
                    Next k
                End If
            Else
                If (matriz(i, j) < matriz(i + 1, j)) Then
                    tsOrdenado = False
                    For k = 0 To limiteX
                        TmpDato(k) = matriz(i, k)
                        matriz(i, k) = matriz(i + 1, k)
                        matriz(i + 1, k) = TmpDato(k)
                    Next k
                End If
            End If
        Next i
    Loop Until tsOrdenado

   On Error GoTo 0
   Exit Sub

Ordenar2_Error:
    Dim ErrNumber As Long: Dim ErrDescription As String: Dim ErrSource As String
    ErrNumber = Err.Number: ErrDescription = Err.Description: ErrSource = Err.Source
    Call HandleException(ErrNumber, ErrDescription, "Lot_fncollections.Ordenar2", ErrSource)
    Err.Raise ErrNumber, "Lot_fncollections.Ordenar2", ErrDescription
End Sub



