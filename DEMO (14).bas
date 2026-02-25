Attribute VB_Name = "Modulo_Generador_Excel_Pro"

Option Explicit

' ======================================================================================
' PROCEDIMIENTO PRINCIPAL: CrearExcelesSeparados
' ======================================================================================
Public Sub CrearExcelesSeparados()
    Dim wsConfig As Worksheet
    Dim colConfiguraciones As New Collection
    Dim configActual As Variant
    Dim rutaDestino As String
    Dim nombreBase As String
    Dim i As Long
    Dim ultimaCol As Long
    Dim contadorProcesados As Integer
    
    ' 1. Configuración inicial y rutas
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("columnas")
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        MsgBox "Error: No se encontró la hoja 'columnas'.", vbCritical
        Exit Sub
    End If
    
    rutaDestino = "C:\CLIENTES\PRUEBAS\BP\"
    If Not AsegurarCarpeta(rutaDestino) Then
        MsgBox "Error: No se pudo crear o acceder a: " & rutaDestino, vbCritical
        Exit Sub
    End If
    
    ' Limpiar el nombre del archivo (quitar la extensión)
    nombreBase = ThisWorkbook.Name
    nombreBase = Left(nombreBase, InStrRev(nombreBase, ".") - 1)
    
    ' 2. Identificar qué configuraciones hay (BOB, BING, etc.) en la fila 3
    ultimaCol = wsConfig.Cells(3, wsConfig.Columns.Count).End(xlToLeft).Column
    For i = 3 To ultimaCol
        If Trim(wsConfig.Cells(3, i).Value) <> "" Then
            colConfiguraciones.Add Trim(wsConfig.Cells(3, i).Value)
        End If
    Next i
    
    If colConfiguraciones.Count = 0 Then
        MsgBox "No se detectaron nombres de configuración en la fila 3.", vbExclamation
        Exit Sub
    End If
    
    ' 3. Optimización de Excel (Silenciar pantalla y alertas)
    IniciarOptimizacion True
    
    ' 4. Bucle principal de creación
    contadorProcesados = 0
    For Each configActual In colConfiguraciones
        Debug.Print "--- PROCESANDO CONFIGURACIÓN: " & configActual & " ---"
        ProcesarUnicoArchivo wsConfig, CStr(configActual), rutaDestino, nombreBase
        contadorProcesados = contadorProcesados + 1
    Next configActual
    
    ' 5. Finalización
    IniciarOptimizacion False
    MsgBox "Proceso completado con éxito." & vbCrLf & "Archivos generados: " & contadorProcesados, vbInformation
End Sub

' ======================================================================================
' SUBPROCESO: Generar cada archivo individual
' ======================================================================================
Private Sub ProcesarUnicoArchivo(ByVal wsRef As Worksheet, ByVal nombreConf As String, ByVal rutaDir As String, ByVal nomArchivo As String)
    Dim wbNuevo As Workbook
    Dim rutaFinal As String
    Dim rutaTemp As String
    Dim colIndice As Long
    Dim seguridadOriginal As Long
    
    rutaFinal = rutaDir & nomArchivo & "_" & nombreConf & ".xlsx"
    rutaTemp = ThisWorkbook.Path & "\~temp_" & nombreConf & ".xlsm"
    
    ' Guardar copia temporal
    ThisWorkbook.SaveCopyAs rutaTemp
    
    ' Bypass de seguridad para evitar el aviso de "Habilitar Macros" al abrir el temporal
    seguridadOriginal = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityLow
    
    Set wbNuevo = Workbooks.Open(rutaTemp, UpdateLinks:=0)
    
    Application.AutomationSecurity = seguridadOriginal
    
    ' --- FASE 1: GESTIÓN DE COLUMNAS ---
    colIndice = BuscarColumnaEnHoja(wsRef, nombreConf)
    
    If colIndice > 0 Then
        EliminarColumnasSegunConfig wbNuevo.Worksheets("FuncionFiltar"), wsRef, colIndice, nombreConf
    End If
    
    ' --- FASE 2: GESTIÓN DE FILAS ---
    ' (Aquí podrías añadir la lógica de filas siguiendo el mismo patrón de limpieza)
    ' ...
    
    ' --- FASE 3: LIMPIEZA Y GUARDADO ---
    Application.DisplayAlerts = False
    On Error Resume Next
    wbNuevo.Worksheets("columnas").Delete
    wbNuevo.Worksheets("filas").Delete
    On Error GoTo 0
    
    ' Guardar como XLSX (formato 51) elimina las macros automáticamente
    wbNuevo.SaveAs Filename:=rutaFinal, FileFormat:=51
    wbNuevo.Close SaveChanges:=False
    
    ' Borrar rastro temporal
    If Dir(rutaTemp) <> "" Then Kill rutaTemp
    Application.DisplayAlerts = True
End Sub

' ======================================================================================
' LÓGICA DE FILTRADO DE COLUMNAS (CON LOGS)
' ======================================================================================
Private Sub EliminarColumnasSegunConfig(ByVal wsDestino As Worksheet, ByVal wsConfig As Worksheet, ByVal colActiva As Long, ByVal nombreC As String)
    Dim ultimaFila As Long
    Dim i As Long
    Dim nombreColumnaInterna As String
    Dim decision As String
    Dim colAEliminar As New Collection
    Dim numColDestino As Long
    Dim item As Variant
    
    ultimaFila = wsConfig.Cells(wsConfig.Rows.Count, 2).End(xlUp).Row
    
    ' Primero identificamos qué columnas hay que borrar
    For i = 4 To ultimaFila
        nombreColumnaInterna = Trim(wsConfig.Cells(i, 2).Value)
        decision = UCase(Trim(wsConfig.Cells(i, colActiva).Value))
        
        ' LOG DE COMPARACIÓN
        Debug.Print "LOG [Col]: '" & nombreColumnaInterna & "' para " & nombreC & " -> Valor: [" & decision & "]"
        
        If decision = "NO" Then
            numColDestino = EncontrarColumnaPorNombre(wsDestino, nombreColumnaInterna)
            If numColDestino > 0 Then
                colAEliminar.Add numColDestino
                Debug.Print "    -> MARCADA PARA ELIMINAR (Col index: " & numColDestino & ")"
            End If
        End If
    Next i
    
    ' Borrar columnas de derecha a izquierda para no perder el índice
    If colAEliminar.Count > 0 Then
        ' Ordenamos la colección de mayor a menor (simple bubble sort para indices)
        Dim arr() As Long: ReDim arr(1 To colAEliminar.Count)
        For i = 1 To colAEliminar.Count: arr(i) = colAEliminar(i): Next i
        
        Dim j As Long, temp As Long
        For i = 1 To UBound(arr) - 1
            For j = i + 1 To UBound(arr)
                If arr(i) < arr(j) Then
                    temp = arr(i): arr(i) = arr(j): arr(j) = temp
                End If
            Next j
        Next i
        
        For i = 1 To UBound(arr)
            wsDestino.Columns(arr(i)).Delete
        Next i
    End If
End Sub

' ======================================================================================
' FUNCIONES DE APOYO
' ======================================================================================

Private Function EncontrarColumnaPorNombre(ByVal ws As Worksheet, ByVal nombreCol As String) As Long
    Dim ultimaCol As Long
    Dim i As Long
    ultimaCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Asumimos fila 1 para cabeceras
    
    ' Si la fila 1 está vacía, buscamos en las primeras 5 filas
    Dim r As Long
    For r = 1 To 5
        For i = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            If Trim(ws.Cells(r, i).Value) = nombreCol Then
                EncontrarColumnaPorNombre = i
                Exit Function
            End If
        Next i
    Next r
End Function

Private Function BuscarColumnaEnHoja(ByVal ws As Worksheet, ByVal texto As String) As Long
    ' Busca en qué columna de la hoja de configuración está el nombre (BOB, BING...)
    Dim i As Long
    Dim ultimaCol As Long
    ultimaCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    For i = 3 To ultimaCol
        If UCase(Trim(ws.Cells(3, i).Value)) = UCase(texto) Then
            BuscarColumnaEnHoja = i
            Exit Function
        End If
    Next i
End Function

Private Function AsegurarCarpeta(ByVal ruta As String) As Boolean
    On Error Resume Next
    ' Crear niveles de carpeta
    MkDir "C:\CLIENTES"
    MkDir "C:\CLIENTES\PRUEBAS"
    MkDir "C:\CLIENTES\PRUEBAS\BP"
    AsegurarCarpeta = (Dir(ruta, vbDirectory) <> "")
    On Error GoTo 0
End Function

Private Sub IniciarOptimizacion(ByVal activar As Boolean)
    Application.ScreenUpdating = Not activar
    Application.DisplayAlerts = Not activar
    Application.Calculation = IIf(activar, xlCalculationManual, xlCalculationAutomatic)
End Sub
