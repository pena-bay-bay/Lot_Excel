Attribute VB_Name = "Lot_PqtNucleoTesting"
'---------------------------------------------------------------------------------------
' Module    : Lot_PqtNucleoTesting
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:24
' Purpose   :
'---------------------------------------------------------------------------------------
'
Option Explicit
Option Base 0



'---------------------------------------------------------------------------------------
' Procedure : CreatePeriodo
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:47
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub CreatePeriodo()
    Dim Obj As Periodo
    Dim cboPrueba As ComboBox
    Dim frm As frmSelPeriodo
    Dim mLista As Variant
    
    Set Obj = New Periodo
    Set frm = New frmSelPeriodo
    Set cboPrueba = frm.cboPerMuestra
    
    mLista = Array(ctPersonalizadas, ctSemanaPasada, ctSemanaActual, ctMesActual, ctHoy, ctAyer, ctLoQueVadeMes, _
                                     ctLoQueVadeSemana)
    
    Obj.CargaCombo cboPrueba, mLista
    
    PintarPeriodo Obj
    Obj.Tipo_Fecha = ctAñoAnterior
    

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ParametrosTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:01:53
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ParametrosTest()
    Dim Obj As Parametros
    
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParametroTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:02
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ParametroTest()
    Dim oPar As Parametro
    Dim iPar As Integer
    Dim dPar As Date
    Dim bPar As Boolean
    Dim pPar As Double
    
    
    Set oPar = New Parametro
    With oPar
        .Descripcion = "Esta es una variable de prueba"
        .Nombre = "VARIABLE"
        .Tipo = parTexto
        .Valor = "Ejemplo"
    End With
    PrintParametro oPar
    '
    ' Prueba entero
    '
    iPar = 3294
    oPar.Valor = iPar
    PrintParametro oPar
    '
    ' prueba fecha
    '
    dPar = #1/1/2014#
    oPar.Valor = dPar
    PrintParametro oPar
    '
    ' prueba Doble
    '
    pPar = 12536.254
    oPar.Valor = pPar
    PrintParametro oPar
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParametrosEngineTest
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:04
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ParametrosEngineTest()

End Sub



'---------------------------------------------------------------------------------------
' Procedure : PintarPeriodo
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PintarPeriodo(datPeriodo As Periodo)
    Debug.Print "==> Periodo "
    Debug.Print "Dias          = " & datPeriodo.Dias
    Debug.Print "Fecha Final   = " & datPeriodo.FechaFinal
    Debug.Print "Fecha Inicial = " & datPeriodo.FechaInicial
    Debug.Print "Texto         = " & datPeriodo.Texto
    Debug.Print "Tipo Fecha    = " & datPeriodo.Tipo_Fecha
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrintParametro
' Author    : CHARLY
' Date      : sáb, 01/nov/2014 21:02:19
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub PrintParametro(Obj As Parametro)
    Debug.Print "==> Parametro "
    Debug.Print "Descripcion     = " & Obj.Descripcion
    Debug.Print "EntidadNegocio  = " & Obj.EntidadNegocio.ClassStorage
    Debug.Print "Fecha Alta      = " & Obj.FechaAlta
    Debug.Print "Fecha Modif.    = " & Obj.FechaModificacion
    Debug.Print "Id              = " & Obj.Id
    Debug.Print "Nombre          = " & Obj.Nombre
    Debug.Print "Orden           = " & Obj.Orden
    Debug.Print "Tipo            = " & Obj.Tipo
    Debug.Print "Valor           = " & Obj.Valor
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ParametrosMetodoTest
' Author    : Charly
' Date      : 19/03/2012
' Purpose   : Probar la clase ParametrosMetodo
'---------------------------------------------------------------------------------------
'
Private Sub ParametrosMetodoTest()
    Dim m_objParMetodo As Metodo
    
    Set m_objParMetodo = New Metodo
    
    
    With m_objParMetodo
        .CriteriosAgrupacion = grpDecenas
        .CriteriosOrdenacion = ordProbabilidad
        .DiasAnalisis = 45
        .Id = 1
        .ModalidadJuego = LP_LB_6_49
        .NumeroSorteos = 40
        .Pronosticos = 6
        .SentidoOrdenacion = True
    End With

    Debug.Print "==> Pruebas ParametrosMetodoTest"
    Debug.Print "Id                       = " & m_objParMetodo.Id
    Debug.Print "Juego                    = " & m_objParMetodo.ModalidadJuego
    Debug.Print "Criterio Ordenación      = " & m_objParMetodo.CriteriosOrdenacion
    Debug.Print "Criterio Agrupación      = " & m_objParMetodo.CriteriosAgrupacion
    Debug.Print "Dias de Analisis         = " & m_objParMetodo.DiasAnalisis
    Debug.Print "Numero de Sorteos        = " & m_objParMetodo.NumeroSorteos
    Debug.Print "Pronosticos              = " & m_objParMetodo.Pronosticos
    Debug.Print "Sentido de la Ordenación = " & m_objParMetodo.SentidoOrdenacion
    Debug.Print "OrdenacionToString()     = " & m_objParMetodo.OrdenacionToString()
    Debug.Print "AgrupacionToString()     = " & m_objParMetodo.AgrupacionToString()
    Debug.Print "ToString()               = " & m_objParMetodo.ToString()

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MetodoTest
' Author    : Charly
' Date      : 19/03/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub MetodoTest()
    Dim m_objMetodo As Metodo
    
    Set m_objMetodo = New Metodo
    
    With m_objMetodo
        .TipoProcedimiento = mtdEstadistico
        .EntidadNegocio.FechaModificacion = Date
'        .EsMultiple = False
        .CriteriosAgrupacion = grpParidad
        .CriteriosOrdenacion = ordDesviacion
        .DiasAnalisis = 42
        .SentidoOrdenacion = True
        .TipoMuestra = True
    End With
    
    
    Debug.Print "==> Pruebas Metodo"
    Debug.Print "ClassStorage           =" & m_objMetodo.EntidadNegocio.ClassStorage
    Debug.Print "FechaAlta              =" & m_objMetodo.EntidadNegocio.FechaAlta
    Debug.Print "FechaBaja              =" & m_objMetodo.EntidadNegocio.FechaBaja
    Debug.Print "FechaModificacion      =" & m_objMetodo.EntidadNegocio.FechaModificacion
    Debug.Print "Id                     =" & m_objMetodo.EntidadNegocio.Id
    Debug.Print "IsDirty                =" & m_objMetodo.EntidadNegocio.IsDirty
    Debug.Print "IsNew                  =" & m_objMetodo.EntidadNegocio.IsNew
    Debug.Print "MarkForDelete          =" & m_objMetodo.EntidadNegocio.MarkForDelete
    Debug.Print "Origen                 =" & m_objMetodo.EntidadNegocio.Origen
    Debug.Print "Situacion              =" & m_objMetodo.EntidadNegocio.Situacion
'    Debug.Print "EsMultiple             =" & m_objMetodo.EsMultiple
    Debug.Print "AgrupacionToString     =" & m_objMetodo.AgrupacionToString
    Debug.Print "CriteriosAgrupacion    =" & m_objMetodo.CriteriosAgrupacion
    Debug.Print "CriteriosOrdenacion    =" & m_objMetodo.CriteriosOrdenacion
    Debug.Print "DiasAnalisis           =" & m_objMetodo.DiasAnalisis
    Debug.Print "Id                     =" & m_objMetodo.Id
    Debug.Print "ModalidadJuego         =" & m_objMetodo.ModalidadJuego
    Debug.Print "NumeroSorteos          =" & m_objMetodo.NumeroSorteos
    Debug.Print "OrdenacionToString     =" & m_objMetodo.OrdenacionToString
    Debug.Print "Pronosticos            =" & m_objMetodo.Pronosticos
    Debug.Print "SentidoOrdenacion      =" & m_objMetodo.SentidoOrdenacion
    Debug.Print "ToString               =" & m_objMetodo.ToString
    Debug.Print "TipoMuestra            =" & m_objMetodo.TipoMuestra
    Debug.Print "TipoProcedimiento      =" & m_objMetodo.TipoProcedimiento
       
End Sub



