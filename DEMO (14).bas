 

Option Explicit

' ============================================================================
' FUNCION PRINCIPAL
' ============================================================================

Public Sub CrearExcelesSeparados()
    Dim wsConfig As Worksheet
    Dim configuraciones As Collection
    Dim config As Variant
    Dim rutaBase As String
    Dim nombreOriginal As String
    Dim totalExcels As Integer
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("columnas")
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        MsgBox "ERROR: No se encuentra la hoja 'columnas'", vbCritical
        Exit Sub
    End If
    
    rutaBase = "C:\CLIENTES\PRUEBAS\BP\"
    
    If Not CrearCarpeta(rutaBase) Then
        MsgBox "ERROR: No se pudo crear la carpeta: " & rutaBase, vbCritical
        Exit Sub
    End If
    
    nombreOriginal = Replace(ThisWorkbook.Name, ".xlsm", "")
    nombreOriginal = Replace(nombreOriginal, ".xlsx", "")
    nombreOriginal = Replace(nombreOriginal, ".xls", "")
    
    Set configuraciones = DetectarConfiguraciones(wsConfig)
    
    If configuraciones.Count = 0 Then
        MsgBox "No se encontraron configuraciones en hoja 'columnas'", vbInformation
        Exit Sub
    End If
   
    ' SILENCIAR ALERTAS PARA EVITAR EL MENSAJE DE "YA EXISTE EL ARCHIVO"
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each config In configuraciones
        CrearExcelParaConfiguracion wsConfig, CStr(config), rutaBase, nombreOriginal
        totalExcels = totalExcels + 1
    Next config
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
  
    MsgBox "Proceso completado! Excels creados: " & totalExcels
End Sub

' ============================================================================
' MODULO 1: CREAR CARPETA
' ============================================================================

Private Function CrearCarpeta(ByVal ruta As String) As Boolean
    On Error Resume Next
    MkDir "C:\CLIENTES"
    MkDir "C:\CLIENTES\PRUEBAS"
    MkDir "C:\CLIENTES\PRUEBAS\BP"
    CrearCarpeta = (Dir(ruta, vbDirectory) <> "")
    On Error GoTo 0
End Function

' ============================================================================
' MODULO 2: DETECTAR CONFIGURACIONES
' ============================================================================

Private Function DetectarConfiguraciones(ByVal wsConfig As Worksheet) As Collection
    Dim configs As Collection, col As Long, nombreConfig As String, ultimaColumna As Long
    Set configs = New Collection
    ultimaColumna = wsConfig.Cells(3, wsConfig.Columns.Count).End(xlToLeft).Column
    For col = 3 To ultimaColumna
        nombreConfig = Trim(wsConfig.Cells(3, col).Value)
        If nombreConfig <> "" Then configs.Add nombreConfig
    Next col
    Set DetectarConfiguraciones = configs
End Function

' ============================================================================
' MODULO 3: CREAR EXCEL PARA CONFIGURACION (CORREGIDO)
' ============================================================================

Private Sub CrearExcelParaConfiguracion(ByVal wsConfig As Worksheet, _
                                         ByVal nombreConfig As String, _
                                         ByVal rutaBase As String, _
                                         ByVal nombreOriginal As String)
    
    Dim wbNuevo As Workbook
    Dim rutaCompleta As String
    Dim rutaTempXLSM As String
    Dim colConfig As Long
    Dim columnasABorrar As Collection
    Dim filasABorrar As Collection
    Dim wsColumnas As Worksheet
    Dim wsFilas As Worksheet
    
    rutaCompleta = rutaBase & nombreOriginal & "_" & nombreConfig & ".xlsx"
    rutaTempXLSM = rutaBase & "temp_" & nombreConfig & ".xlsm"
    
    ' CORRECCIÓN: Guardamos copia como XLSM temporal para que Excel no se queje al abrirlo
    ThisWorkbook.SaveCopyAs rutaTempXLSM
    Set wbNuevo = Workbooks.Open(rutaTempXLSM)
    
    ' PARTE 1: COLUMNAS
    On Error Resume Next
    Set wsColumnas = wbNuevo.Worksheets("columnas")
    On Error GoTo 0
    
    If Not wsColumnas Is Nothing Then
        colConfig = BuscarColumnaConfiguracion(wsColumnas, nombreConfig)
        If colConfig > 0 Then
            Set columnasABorrar = LeerColumnasABorrar(wsColumnas, wbNuevo.Worksheets("FuncionFiltar"), colConfig)
            If columnasABorrar.Count > 0 Then BorrarColumnas wbNuevo.Worksheets("FuncionFiltar"), columnasABorrar
        End If
    End If
    
    ' PARTE 2: FILAS
    On Error Resume Next
    Set wsFilas = wbNuevo.Worksheets("filas")
    On Error GoTo 0
    
    If Not wsFilas Is Nothing Then
        colConfig = BuscarColumnaConfiguracion(wsFilas, nombreConfig)
        If colConfig > 0 Then
            Dim wsTEXOENFILADOS As Worksheet
            On Error Resume Next
            Set wsTEXOENFILADOS = wbNuevo.Worksheets("TEXOENFILADOS")
            On Error GoTo 0
            If Not wsTEXOENFILADOS Is Nothing Then
                Set filasABorrar = LeerFilasABorrar(wsFilas, wsTEXOENFILADOS, colConfig)
                If filasABorrar.Count > 0 Then BorrarFilas wsTEXOENFILADOS, filasABorrar
            End If
        End If
    End If
    
    ' PARTE 3: LIMPIEZA HOJAS
    Application.DisplayAlerts = False
    On Error Resume Next
    wbNuevo.Worksheets("columnas").Delete
    wbNuevo.Worksheets("filas").Delete
    On Error GoTo 0
    
    ' FINALIZAR: GUARDAR COMO XLSX Y ELIMINAR TEMPORAL
    ' Al usar FileFormat 51 se eliminan las macros automáticamente
    wbNuevo.SaveAs Filename:=rutaCompleta, FileFormat:=51
    wbNuevo.Close SaveChanges:=False
    
    ' Borrar el temporal que usamos para trabajar
    If Dir(rutaTempXLSM) <> "" Then Kill rutaTempXLSM
    Application.DisplayAlerts = True
End Sub

' ============================================================================
' MODULOS DE APOYO (SIN CAMBIOS RESPECTO AL ORIGINAL)
' ============================================================================

Private Function BuscarColumnaConfiguracion(ByVal ws As Worksheet, ByVal nombreConfig As String) As Long
    Dim col As Long, ultimaColumna As Long, filaEncabezado As Long
    For filaEncabezado = 1 To 5
        If InStr(1, ws.Cells(filaEncabezado, 2).Value, nombreConfig, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(filaEncabezado, 3).Value, nombreConfig, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(filaEncabezado, 4).Value, nombreConfig, vbTextCompare) > 0 Then
            Exit For
        End If
    Next filaEncabezado
    If filaEncabezado > 5 Then filaEncabezado = 2
    ultimaColumna = ws.Cells(filaEncabezado, ws.Columns.Count).End(xlToLeft).Column
    For col = 1 To ultimaColumna
        If UCase(Trim(ws.Cells(filaEncabezado, col).Value)) = UCase(nombreConfig) Then
            BuscarColumnaConfiguracion = col
            Exit Function
        End If
    Next col
End Function

Private Function LeerColumnasABorrar(ByVal wsConfig As Worksheet, ByVal wsOrigen As Worksheet, ByVal colConfig As Long) As Collection
    Dim columnas As Collection, fila As Long, ultimaFila As Long, nombreColumna As String, valor As String, numColOrigen As Long
    Set columnas = New Collection
    ultimaFila = wsConfig.Cells(wsConfig.Rows.Count, 2).End(xlUp).Row
    For fila = 4 To ultimaFila
        nombreColumna = Trim(wsConfig.Cells(fila, 2).Value)
        If nombreColumna <> "" Then
            valor = Trim(wsConfig.Cells(fila, colConfig).Value)
            If UCase(valor) = "NO" Then
                numColOrigen = BuscarColumnaEnOrigen(wsOrigen, nombreColumna)
                If numColOrigen > 0 Then columnas.Add numColOrigen
            End If
        End If
    Next fila
    Set LeerColumnasABorrar = columnas
End Function

Private Function BuscarColumnaEnOrigen(ByVal ws As Worksheet, ByVal nombreBuscado As String) As Long
    Dim col As Long, ultimaCol As Long, filaEncabezado As Long
    For filaEncabezado = 1 To 10
        If ws.Cells(filaEncabezado, 1).Value <> "" Then Exit For
    Next filaEncabezado
    ultimaCol = ws.Cells(filaEncabezado, ws.Columns.Count).End(xlToLeft).Column
    For col = 1 To ultimaCol
        If Trim(ws.Cells(filaEncabezado, col).Value) = nombreBuscado Then
            BuscarColumnaEnOrigen = col
            Exit Function
        End If
    Next col
End Function

Private Sub BorrarColumnas(ByVal ws As Worksheet, ByVal columnasABorrar As Collection)
    Dim arr() As Long, i As Long, j As Long, temp As Long, numCol As Variant
    If columnasABorrar.Count = 0 Then Exit Sub
    ReDim arr(1 To columnasABorrar.Count)
    i = 1: For Each numCol In columnasABorrar: arr(i) = CLng(numCol): i = i + 1: Next numCol
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then temp = arr(i): arr(i) = arr(j): arr(j) = temp
        Next j
    Next i
    For i = 1 To UBound(arr): ws.Columns(arr(i)).Delete: Next i
End Sub

Private Function LeerFilasABorrar(ByVal wsConfig As Worksheet, ByVal wsOrigen As Worksheet, ByVal colConfig As Long) As Collection
    Dim filas As Collection, fila As Long, ultimaFila As Long, textoLinea As String, valor As String, numFilaOrigen As Long
    Dim colTexto As Long, col As Long, filaInicio As Long, colExtra As Long, textoExtra As String
    Set filas = New Collection
    colExtra = colConfig + 5
    filaInicio = 3
    For fila = 2 To 10
        valor = UCase(Trim(wsConfig.Cells(fila, colConfig).Value))
        If valor = "NO" Or valor = "SI" Then filaInicio = fila: Exit For
    Next fila
    colTexto = 0: Dim maxLen As Integer: maxLen = 0
    For col = 1 To 20
        valor = Trim(wsConfig.Cells(filaInicio, col).Value)
        If Len(valor) > maxLen And Len(valor) > 20 Then maxLen = Len(valor): colTexto = col
    Next col
    If colTexto = 0 Then Set LeerFilasABorrar = filas: Exit Function
    ultimaFila = wsConfig.Cells(wsConfig.Rows.Count, colTexto).End(xlUp).Row
    For fila = filaInicio To ultimaFila
        textoLinea = Trim(wsConfig.Cells(fila, colTexto).Value)
        If Len(textoLinea) > 5 Then
            valor = Trim(wsConfig.Cells(fila, colConfig).Value)
            textoExtra = Trim(wsConfig.Cells(fila, colExtra).Value)
            If UCase(valor) = "NO" Then
                numFilaOrigen = BuscarFilaPorTexto(wsOrigen, textoLinea)
                If numFilaOrigen > 0 Then filas.Add numFilaOrigen & "|" & textoExtra
            ElseIf textoExtra <> "" Then
                numFilaOrigen = BuscarFilaPorTexto(wsOrigen, textoLinea)
                If numFilaOrigen > 0 Then filas.Add numFilaOrigen & "|AÑADIR|" & textoExtra
            End If
        End If
    Next fila
    Set LeerFilasABorrar = filas
End Function

Private Function BuscarFilaPorTexto(ByVal ws As Worksheet, ByVal textoBuscado As String) As Long
    Dim fila As Long, ultimaFila As Long, textoFila As String, col As Long, textoBuscar As String
    ultimaFila = 1
    For col = 1 To 20
        If ws.Cells(ws.Rows.Count, col).End(xlUp).Row > ultimaFila Then ultimaFila = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    Next col
    textoBuscar = IIf(Len(Trim(textoBuscado)) > 20, Left(Trim(textoBuscado), 20), Trim(textoBuscado))
    textoBuscar = Replace(textoBuscar, "  ", " ")
    For fila = 1 To ultimaFila
        For col = 1 To 20
            textoFila = Replace(Trim(ws.Cells(fila, col).Value), "  ", " ")
            If Len(textoFila) > 10 And InStr(1, textoFila, textoBuscar, vbTextCompare) > 0 Then
                BuscarFilaPorTexto = fila
                Exit Function
            End If
        Next col
    Next fila
End Function

Private Sub BorrarFilas(ByVal ws As Worksheet, ByVal filasABorrar As Collection)
    Dim arr() As String, i As Long, j As Long, temp As String, accion As Variant, partes() As String
    Dim numFila As Long, textoExtra As String, colTexto As Long, col As Long
    If filasABorrar.Count = 0 Then Exit Sub
    For Each accion In filasABorrar
        partes = Split(CStr(accion), "|")
        If UBound(partes) >= 2 And partes(1) = "AÑADIR" Then
            numFila = CLng(partes(0)): textoExtra = partes(2): colTexto = 0
            For col = 1 To 20
                If Len(Trim(ws.Cells(numFila, col).Value)) > 20 Then colTexto = col: Exit For
            Next col
            If colTexto = 0 Then colTexto = 2
            ws.Cells(numFila, colTexto + 1).Value = textoExtra
        End If
    Next accion
    Dim filasBorrar As Collection: Set filasBorrar = New Collection
    For Each accion In filasABorrar
        partes = Split(CStr(accion), "|")
        If UBound(partes) = 0 Or (UBound(partes) >= 1 And partes(1) <> "AÑADIR") Then filasBorrar.Add CLng(partes(0))
    Next accion
    If filasBorrar.Count = 0 Then Exit Sub
    ReDim arr(1 To filasBorrar.Count)
    i = 1: For Each accion In filasBorrar: arr(i) = CStr(accion): i = i + 1: Next accion
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CLng(arr(i)) < CLng(arr(j)) Then temp = arr(i): arr(i) = arr(j): arr(j) = temp
        Next j
    Next i
    For i = 1 To UBound(arr): ws.Rows(CLng(arr(i))).Delete: Next i
End Sub

