Attribute VB_Name = "Modulo_Generador_Excel_Final"

Option Explicit

' ======================================================================================
' PROCEDIMIENTO PRINCIPAL
' ======================================================================================
Public Sub CrearExcelesSeparados()
    Dim wsConfiguracion As Worksheet
    Dim coleccionClientes As New Collection
    Dim clienteActual As Variant
    Dim rutaDestino As String
    Dim nombreLibroOriginal As String
    Dim i As Long
    Dim ultimaColumnaConfig As Long
    Dim contadorArchivos As Integer
    
    ' 1. Validar que la hoja 'columnas' existe
    On Error Resume Next
    Set wsConfiguracion = ThisWorkbook.Worksheets("columnas")
    On Error GoTo 0
    
    If wsConfiguracion Is Nothing Then
        MsgBox "Error Crítico: No se encontró la hoja 'columnas'.", vbCritical
        Exit Sub
    End If
    
    ' 2. Definir ruta y asegurar que las carpetas existan
    rutaDestino = "C:\CLIENTES\PRUEBAS\BP\"
    If Not AsegurarRutaLocal(rutaDestino) Then
        MsgBox "Error: No se puede acceder a la ruta " & rutaDestino, vbCritical
        Exit Sub
    End If
    
    ' 3. Obtener nombre del archivo sin extensiones
    nombreLibroOriginal = ThisWorkbook.Name
    nombreLibroOriginal = Replace(nombreLibroOriginal, ".xlsm", "")
    nombreLibroOriginal = Replace(nombreLibroOriginal, ".xlsx", "")
    nombreLibroOriginal = Replace(nombreLibroOriginal, ".xls", "")
    
    ' 4. Leer los nombres de las configuraciones (BOB, BING, etc.) de la fila 3
    ultimaColumnaConfig = wsConfiguracion.Cells(3, wsConfiguracion.Columns.Count).End(xlToLeft).Column
    For i = 3 To ultimaColumnaConfig
        If Trim(wsConfiguracion.Cells(3, i).Value) <> "" Then
            coleccionClientes.Add Trim(wsConfiguracion.Cells(3, i).Value)
        End If
    Next i
    
    If coleccionClientes.Count = 0 Then
        MsgBox "No hay configuraciones activas en la fila 3 de la hoja 'columnas'.", vbExclamation
        Exit Sub
    End If
    
    ' 5. Optimizar rendimiento
    GestionarEntorno True
    
    ' 6. BUCLE DE GENERACION
    contadorArchivos = 0
    For Each clienteActual In coleccionClientes
        Debug.Print "--------------------------------------------------------"
        Debug.Print "INICIANDO PROCESO PARA: " & clienteActual
        Debug.Print "--------------------------------------------------------"
        
        EjecutarGeneracionIndividual wsConfiguracion, CStr(clienteActual), rutaDestino, nombreLibroOriginal
        contadorArchivos = contadorArchivos + 1
    Next clienteActual
    
    ' 7. Restaurar Excel
    GestionarEntorno False
    
    MsgBox "Proceso finalizado correctamente." & vbCrLf & "Archivos generados: " & contadorArchivos, vbInformation
End Sub

' ======================================================================================
' PROCESO INDIVIDUAL POR ARCHIVO
' ======================================================================================
Private Sub EjecutarGeneracionIndividual(ByVal wsRef As Worksheet, ByVal idConfig As String, ByVal rutaCarpeta As String, ByVal nomBase As String)
    Dim wbCopia As Workbook
    Dim fFinal As String
    Dim fTemporal As String
    Dim indiceColumnaConfig As Long
    Dim seguridadOriginal As Long
    
    fFinal = rutaCarpeta & nomBase & "_" & idConfig & ".xlsx"
    ' El temporal se guarda en la misma carpeta que el original para mayor confianza de Excel
    fTemporal = ThisWorkbook.Path & "\~tmp_" & idConfig & ".xlsm"
    
    ' Crear copia temporal
    On Error Resume Next
    If Dir(fTemporal) <> "" Then Kill fTemporal
    On Error GoTo 0
    ThisWorkbook.SaveCopyAs fTemporal
    
    ' --- BYPASS DE SEGURIDAD PARA MACROS ---
    seguridadOriginal = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityLow
    
    ' Abrir el temporal sin actualizar links
    Set wbCopia = Workbooks.Open(fTemporal, UpdateLinks:=0)
    
    ' Restaurar seguridad inmediatamente
    Application.AutomationSecurity = seguridadOriginal
    
    ' Buscar en qué columna de la hoja de configuración está el cliente (BOB, etc.)
    indiceColumnaConfig = BuscarIndiceConfiguracion(wsRef, idConfig)
    
    If indiceColumnaConfig > 0 Then
        ' 1. Procesar Columnas (Hoja FuncionFiltar)
        ProcesarBorradoColumnas wbCopia, idConfig, indiceColumnaConfig
        
        ' 2. Procesar Filas (Hoja TEXOENFILADOS)
        ProcesarBorradoYModificadoFilas wbCopia, idConfig, indiceColumnaConfig
    End If
    
    ' --- LIMPIEZA FINAL DEL ARCHIVO ---
    Application.DisplayAlerts = False
    On Error Resume Next
    wbCopia.Worksheets("columnas").Delete
    wbCopia.Worksheets("filas").Delete
    On Error GoTo 0
    
    ' Guardar como XLSX (formato 51) para limpiar macros definitivamente
    wbCopia.SaveAs Filename:=fFinal, FileFormat:=51
    wbCopia.Close SaveChanges:=False
    
    ' Eliminar archivo temporal .xlsm
    If Dir(fTemporal) <> "" Then Kill fTemporal
    Application.DisplayAlerts = True
    
    Debug.Print "ARCHIVO FINALIZADO: " & fFinal
End Sub

' ======================================================================================
' GESTION DE COLUMNAS (CON LOGS)
' ======================================================================================
Private Sub ProcesarBorradoColumnas(ByRef wb As Workbook, ByVal configNombre As String, ByVal colIndex As Long)
    Dim wsConfigCol As Worksheet
    Dim wsDestinoCol As Worksheet
    Dim ultimaFilaConf As Long
    Dim i As Long
    Dim nombreColumnaABuscar As String
    Dim valorConfig As String
    Dim colIndexEnDestino As Long
    Dim listaBorrado As New Collection
    
    Set wsConfigCol = wb.Worksheets("columnas")
    Set wsDestinoCol = wb.Worksheets("FuncionFiltar")
    
    ultimaFilaConf = wsConfigCol.Cells(wsConfigCol.Rows.Count, 2).End(xlUp).Row
    
    Debug.Print "LOG COLUMNAS [" & configNombre & "]:"
    
    For i = 4 To ultimaFilaConf
        nombreColumnaABuscar = Trim(wsConfigCol.Cells(i, 2).Value)
        valorConfig = UCase(Trim(wsConfigCol.Cells(i, colIndex).Value))
        
        If nombreColumnaABuscar <> "" Then
            If valorConfig = "NO" Then
                colIndexEnDestino = EncontrarColumnaPorTexto(wsDestinoCol, nombreColumnaABuscar)
                If colIndexEnDestino > 0 Then
                    listaBorrado.Add colIndexEnDestino
                    Debug.Print "  - Marcada para borrar: '" & nombreColumnaABuscar & "' (Col " & colIndexEnDestino & ")"
                End If
            End If
        End If
    Next i
    
    ' Borrar columnas de derecha a izquierda
    BorrarElementosEstructurales wsDestinoCol, listaBorrado, "COLUMNA"
End Sub

' ======================================================================================
' GESTION DE FILAS (CON LOGS)
' ======================================================================================
Private Sub ProcesarBorradoYModificadoFilas(ByRef wb As Workbook, ByVal configNombre As String, ByVal colIndex As Long)
    Dim wsConfigFilas As Worksheet
    Dim wsDestinoFilas As Worksheet
    Dim ultimaFilaConf As Long
    Dim i As Long
    Dim textoABuscar As String
    Dim valorConfig As String
    Dim filaEncontrada As Long
    Dim listaBorradoFilas As New Collection
    Dim textoExtra As String
    
    Set wsConfigFilas = wb.Worksheets("filas")
    Set wsDestinoFilas = wb.Worksheets("TEXOENFILADOS")
    
    ultimaFilaConf = wsConfigFilas.Cells(wsConfigFilas.Rows.Count, 6).End(xlUp).Row
    
    Debug.Print "LOG FILAS [" & configNombre & "]:"
    
    For i = 3 To ultimaFilaConf
        textoABuscar = Trim(wsConfigFilas.Cells(i, 6).Value)
        valorConfig = UCase(Trim(wsConfigFilas.Cells(i, colIndex).Value))
        
        If textoABuscar <> "" Then
            filaEncontrada = BuscarFilaPorTexto(wsDestinoFilas, textoABuscar)
            
            If filaEncontrada > 0 Then
                If valorConfig = "NO" Then
                    listaBorradoFilas.Add filaEncontrada
                    Debug.Print "  - Fila marcada para BORRAR: '" & Left(textoABuscar, 30) & "...' (Fila " & filaEncontrada & ")"
                Else
                    ' Si no se borra, comprobamos si hay que añadir texto (columna del cliente + 5)
                    textoExtra = Trim(wsConfigFilas.Cells(i, colIndex + 5).Value)
                    If textoExtra <> "" Then
                        wsDestinoFilas.Cells(filaEncontrada, 3).Value = textoExtra
                        Debug.Print "  - Fila MODIFICADA (Texto añadido): '" & textoExtra & "' en fila " & filaEncontrada
                    End If
                End If
            End If
        End If
    Next i
    
    ' Borrar filas de abajo hacia arriba
    BorrarElementosEstructurales wsDestinoFilas, listaBorradoFilas, "FILA"
End Sub

' ======================================================================================
' FUNCIONES AUXILIARES DE BUSQUEDA Y BORRADO
' ======================================================================================

Private Function BuscarIndiceConfiguracion(ByVal ws As Worksheet, ByVal nombre As String) As Long
    Dim c As Long
    Dim ultimaCol As Long
    ultimaCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    For c = 3 To ultimaCol
        If UCase(Trim(ws.Cells(3, c).Value)) = UCase(nombre) Then
            BuscarIndiceConfiguracion = c
            Exit Function
        End If
    Next c
End Function

Private Function EncontrarColumnaPorTexto(ByVal ws As Worksheet, ByVal txt As String) As Long
    Dim c As Long
    Dim r As Long
    ' Busca en las primeras 5 filas por si la cabecera no está en la 1
    For r = 1 To 5
        For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            If UCase(Trim(ws.Cells(r, c).Value)) = UCase(txt) Then
                EncontrarColumnaPorTexto = c
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function BuscarFilaPorTexto(ByVal ws As Worksheet, ByVal txt As String) As Long
    Dim r As Long
    Dim ultimaFila As Long
    Dim fragmentoBuscado As String
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' Usamos un fragmento para evitar problemas con textos demasiado largos
    fragmentoBuscado = IIf(Len(txt) > 50, Left(txt, 50), txt)
    For r = 1 To ultimaFila
        If InStr(1, ws.Cells(r, 1).Value, fragmentoBuscado, vbTextCompare) > 0 Then
            BuscarFilaPorTexto = r
            Exit Function
        End If
    Next r
End Function

Private Sub BorrarElementosEstructurales(ByVal ws As Worksheet, ByVal coleccionIndices As Collection, ByVal tipo As String)
    If coleccionIndices.Count = 0 Then Exit Sub
    
    ' Convertir a array para ordenar de mayor a menor
    Dim arr() As Long
    ReDim arr(1 To coleccionIndices.Count)
    Dim i As Long, j As Long, temp As Long
    
    For i = 1 To coleccionIndices.Count
        arr(i) = coleccionIndices(i)
    Next i
    
    ' Ordenar descendente (Bubble Sort)
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    ' Ejecutar borrado
    For i = 1 To UBound(arr)
        If tipo = "COLUMNA" Then
            ws.Columns(arr(i)).Delete
        Else
            ws.Rows(arr(i)).Delete
        End If
    Next i
End Sub

' ======================================================================================
' UTILIDADES DE SISTEMA
' ======================================================================================

Private Function AsegurarRutaLocal(ByVal ruta As String) As Boolean
    On Error Resume Next
    MkDir "C:\CLIENTES"
    MkDir "C:\CLIENTES\PRUEBAS"
    MkDir "C:\CLIENTES\PRUEBAS\BP"
    AsegurarRutaLocal = (Dir(ruta, vbDirectory) <> "")
    On Error GoTo 0
End Function

Private Sub GestionarEntorno(ByVal activar As Boolean)
    Application.ScreenUpdating = Not activar
    Application.DisplayAlerts = Not activar
    Application.EnableEvents = Not activar
    Application.Calculation = IIf(activar, xlCalculationManual, xlCalculationAutomatic)
End Sub
