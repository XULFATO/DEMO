Attribute VB_Name = "Módulo2"

Option Explicit

' ======================================================================================
' CONSTANTES GLOBALES DE MÓDULO
' ======================================================================================

' Contraseña única para desproteger hojas — cambiar aquí si se actualiza
Private Const PASSWORD_HOJAS   As String = "ADP"

' Valores reconocidos en las columnas de configuración
Private Const CONFIG_MANTENER  As String = "MANTENER"
Private Const CONFIG_QUITAR    As String = "QUITAR"

' Ruta de red por defecto
Private Const UNIDAD_RED       As String = "O:\"
Private Const RUTA_POR_DEFECTO As String = "O:\CLIENTES\PRUEBAS\BP\"

' Nombre de la hoja de datos principal (la que se recorta según cliente)
' *** AJUSTAR AQUÍ si el nombre de la hoja cambia ***
Private Const HOJA_DATOS       As String = "Analisis conceptos BOB"
Private Const HOJA_PREGUNTAS   As String = "Preguntas generales"

' ======================================================================================
' BOTONES PRINCIPALES
' ======================================================================================

Public Sub GenerarBOB()
    GenerarArchivoIndividual "BOB"
End Sub

Public Sub GenerarCELERGO()
    GenerarArchivoIndividual "CELERGO"
End Sub

Public Sub GenerarBOByCELERGO()
    Dim clientes(1) As String
    clientes(0) = "BOB"
    clientes(1) = "CELERGO"
    
    Dim i As Integer
    For i = LBound(clientes) To UBound(clientes)
        GenerarArchivoIndividual clientes(i)
    Next i
End Sub

' ======================================================================================
' FUNCIÓN CENTRALIZADA
' ======================================================================================

Private Sub GenerarArchivoIndividual(ByVal clienteNombre As String)
    Dim wsConfigCol As Worksheet
    Dim rutaDestino As String
    Dim nombreBase  As String
    Dim indiceCol   As Long
    
    Debug.Print String(80, "=")
    Debug.Print "INICIO PROCESO PARA: " & clienteNombre
    Debug.Print String(80, "=")
    
    ' 1. Validar hoja de configuración de columnas
    On Error Resume Next
    Set wsConfigCol = ThisWorkbook.Worksheets("columnas")
    On Error GoTo 0
    
    If wsConfigCol Is Nothing Then
        MsgBox "Error crítico: no se encontró la hoja 'columnas'.", vbCritical
        Exit Sub
    End If
    
    ' 2. Validar que el cliente existe en la configuración
    indiceCol = BuscarIndiceConfiguracion(wsConfigCol, clienteNombre)
    If indiceCol = 0 Then
        MsgBox "Error: no se encontró la configuración para '" & clienteNombre & "' en la hoja 'columnas'.", vbExclamation
        Exit Sub
    End If
    Debug.Print "[OK] '" & clienteNombre & "' encontrado en columna: " & indiceCol
    
    ' 3. Obtener ruta de destino (con control de unidad de red)
    rutaDestino = ObtenerRutaDestino()
    If rutaDestino = "" Then
        Debug.Print "[CANCELADO] El usuario canceló la operación."
        Exit Sub
    End If
    Debug.Print "[OK] Ruta seleccionada: " & rutaDestino
    
    ' 4. Nombre base del archivo sin extensión
    nombreBase = LimpiarExtension(ThisWorkbook.Name)
    Debug.Print "[OK] Nombre base: " & nombreBase
    
    ' 5. Optimizar rendimiento
    GestionarEntorno True
    
    ' 6. Generar archivo — pasamos indiceCol ya calculado para no recalcular dentro
    EjecutarGeneracion clienteNombre, indiceCol, rutaDestino, nombreBase
    
    ' 7. Restaurar entorno (siempre, aunque haya error dentro de EjecutarGeneracion)
    GestionarEntorno False
    
    Debug.Print String(80, "=")
    Debug.Print "PROCESO COMPLETADO PARA: " & clienteNombre
    Debug.Print String(80, "=")
    
    MsgBox "Proceso finalizado correctamente para " & clienteNombre & "." & vbCrLf & _
           "Archivo generado en:" & vbCrLf & rutaDestino & clienteNombre & "_" & nombreBase & ".xlsx", _
           vbInformation
End Sub

' ======================================================================================
' UTILIDAD: LIMPIAR EXTENSIÓN DEL NOMBRE DE ARCHIVO
' ======================================================================================

Private Function LimpiarExtension(ByVal nombre As String) As String
    Dim extensiones As Variant
    Dim ext         As Variant
    extensiones = Array(".xlsm", ".xlsx", ".xls")
    For Each ext In extensiones
        nombre = Replace(nombre, ext, "", Compare:=vbTextCompare)
    Next ext
    LimpiarExtension = nombre
End Function

' ======================================================================================
' CONTROL DE RUTA DE DESTINO (RED O LOCAL)
' ======================================================================================

Private Function ObtenerRutaDestino() As String
    Dim rutaSeleccionada As String
    Dim respuesta        As VbMsgBoxResult
    
    If Not UnidadRedConectada(UNIDAD_RED) Then
        respuesta = MsgBox( _
            "La unidad de red " & UNIDAD_RED & " no está conectada." & vbCrLf & vbCrLf & _
            "Debe iniciar sesión en la red para acceder a " & UNIDAD_RED & vbCrLf & vbCrLf & _
            "¿Desea seleccionar una carpeta local como alternativa?", _
            vbExclamation + vbYesNo, "Unidad de red no disponible")
        
        If respuesta = vbNo Then
            ObtenerRutaDestino = ""
            Exit Function
        End If
        
        rutaSeleccionada = SeleccionarCarpeta("C:\")
    Else
        rutaSeleccionada = SeleccionarCarpeta(RUTA_POR_DEFECTO)
    End If
    
    If rutaSeleccionada = "" Then
        ObtenerRutaDestino = ""
        Exit Function
    End If
    
    If Right(rutaSeleccionada, 1) <> "\" Then rutaSeleccionada = rutaSeleccionada & "\"
    ObtenerRutaDestino = rutaSeleccionada
End Function

' ======================================================================================
' VERIFICAR UNIDAD DE RED
' ======================================================================================

Private Function UnidadRedConectada(ByVal unidad As String) As Boolean
    Dim fso   As Object
    Dim drive As Object
    
    UnidadRedConectada = False  ' Valor por defecto seguro
    
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function  ' No se pudo crear FSO
    
    If fso.DriveExists(Left(unidad, 1)) Then
        Set drive = fso.GetDrive(Left(unidad, 1))
        If Not drive Is Nothing Then
            UnidadRedConectada = drive.IsReady
        End If
    End If
    
    Set drive = Nothing
    Set fso   = Nothing
    On Error GoTo 0
End Function

' ======================================================================================
' DIÁLOGO DE SELECCIÓN DE CARPETA
' ======================================================================================

Private Function SeleccionarCarpeta(Optional ByVal rutaInicial As String = "C:\") As String
    Dim shellApp    As Object
    Dim shellFolder As Object
    
    On Error Resume Next
    Set shellApp    = CreateObject("Shell.Application")
    Set shellFolder = shellApp.BrowseForFolder(0, "Seleccione la carpeta donde guardar los archivos:", 0, rutaInicial)
    
    If shellFolder Is Nothing Then
        SeleccionarCarpeta = ""
    Else
        SeleccionarCarpeta = shellFolder.Self.Path
    End If
    
    Set shellFolder = Nothing
    Set shellApp    = Nothing
    On Error GoTo 0
End Function

' ======================================================================================
' CREAR RUTA SI NO EXISTE (COMPATIBLE CON RED)
' ======================================================================================

Private Function AsegurarRuta(ByVal ruta As String) As Boolean
    Dim fso       As Object
    Dim carpetas() As String
    Dim rutaAcum  As String
    Dim i         As Integer
    
    On Error Resume Next
    Set fso    = CreateObject("Scripting.FileSystemObject")
    carpetas   = Split(ruta, "\")
    rutaAcum   = carpetas(0)
    
    For i = 1 To UBound(carpetas)
        If carpetas(i) <> "" Then
            rutaAcum = rutaAcum & "\" & carpetas(i)
            If Not fso.FolderExists(rutaAcum) Then
                fso.CreateFolder rutaAcum
                Debug.Print "[CREADA] " & rutaAcum
            End If
        End If
    Next i
    
    AsegurarRuta = fso.FolderExists(ruta)
    Set fso = Nothing
    On Error GoTo 0
End Function

' ======================================================================================
' MANEJO DE PROTECCIÓN DE HOJAS
' ======================================================================================

' Desprotege una hoja: primero sin contraseña (hojas sin clave), luego con ella
Private Sub DesprotegerHoja(ByVal ws As Worksheet, ByVal pwd As String)
    If Not ws.ProtectContents Then Exit Sub
    
    Debug.Print "  [PASS] Desprotegiendo: " & ws.Name
    
    On Error Resume Next
    ws.Unprotect               ' Intento sin contraseña
    If ws.ProtectContents Then _
        ws.Unprotect pwd       ' Si sigue protegida, usar contraseña
    On Error GoTo 0
    
    If ws.ProtectContents Then
        Debug.Print "    [AVISO] No se pudo desproteger: " & ws.Name
    Else
        Debug.Print "    [OK] Desprotegida: " & ws.Name
    End If
End Sub

' Desprotege todas las hojas de un libro
Private Sub DesprotegerTodasHojas(ByVal wb As Workbook, ByVal pwd As String)
    Dim ws As Worksheet
    Debug.Print "[PASS] Desprotegiendo todas las hojas..."
    For Each ws In wb.Worksheets
        DesprotegerHoja ws, pwd
    Next ws
End Sub

' Verificación final: ninguna hoja debe quedar protegida
Private Sub AsegurarHojasDesprotegidas(ByVal wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.ProtectContents Then
            Debug.Print "  [AVISO] Hoja aún protegida tras proceso: " & ws.Name
            DesprotegerHoja ws, PASSWORD_HOJAS
        End If
    Next ws
End Sub

' ======================================================================================
' PROCESO PRINCIPAL DE GENERACIÓN
' BUG CORREGIDO: ya no recalcula indiceCol internamente — llega como parámetro
' BUG CORREGIDO: DisplayAlerts se restaura siempre, incluso si hay error al guardar
' ======================================================================================

Private Sub EjecutarGeneracion(ByVal idConfig As String, ByVal indiceCol As Long, _
                                ByVal rutaCarpeta As String, ByVal nomBase As String)
    Dim wbCopia   As Workbook
    Dim fFinal    As String
    Dim fTemporal As String
    Dim seguridad As Long
    
    fFinal    = rutaCarpeta & idConfig & "_" & nomBase & ".xlsx"
    fTemporal = ThisWorkbook.Path & "\~tmp_" & idConfig & ".xlsm"
    
    Debug.Print "[1] Destino final: " & fFinal
    Debug.Print "[2] Creando copia temporal..."
    
    On Error Resume Next
    If Dir(fTemporal) <> "" Then Kill fTemporal
    On Error GoTo 0
    
    ThisWorkbook.SaveCopyAs fTemporal
    Debug.Print "  [OK] Temporal creado: " & fTemporal
    
    ' Abrir temporal sin ejecutar macros automáticas
    seguridad = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityLow
    Set wbCopia = Workbooks.Open(fTemporal, UpdateLinks:=0)
    Application.AutomationSecurity = seguridad
    Debug.Print "[3] Temporal abierto"
    
    ' Desproteger todas las hojas del libro copia
    Debug.Print "[4] Desprotegiendo hojas..."
    DesprotegerTodasHojas wbCopia, PASSWORD_HOJAS
    
    ' Procesar columnas y filas según configuración
    Debug.Print "[5] Procesando columnas..."
    ProcesarColumnas wbCopia, idConfig, indiceCol
    
    Debug.Print "[6] Procesando filas..."
    ProcesarFilas wbCopia, idConfig
    
    ' Eliminar hojas de configuración del archivo de destino
    Debug.Print "[7] Eliminando hojas de configuración..."
    Application.DisplayAlerts = False
    On Error Resume Next
    wbCopia.Worksheets("columnas").Delete
    wbCopia.Worksheets("filas").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True   ' Restaurar siempre tras el bloque de borrado
    
    ' Verificación final de protección
    Debug.Print "[8] Verificación final de protección..."
    AsegurarHojasDesprotegidas wbCopia
    
    ' Crear ruta si no existe
    If Not AsegurarRuta(rutaCarpeta) Then
        MsgBox "Error: no se puede acceder a la ruta " & rutaCarpeta, vbCritical
        wbCopia.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Guardar como xlsx
    Debug.Print "[9] Guardando como .xlsx..."
    wbCopia.SaveAs Filename:=fFinal, FileFormat:=51
    wbCopia.Close SaveChanges:=False
    Debug.Print "  [OK] Guardado: " & fFinal
    
    ' Eliminar temporal
    On Error Resume Next
    If Dir(fTemporal) <> "" Then Kill fTemporal
    On Error GoTo 0
    Debug.Print "[10] Temporal eliminado"
    Debug.Print "[COMPLETADO] " & idConfig
End Sub

' ======================================================================================
' PROCESADO DE COLUMNAS
' BUG CORREGIDO: nombre de hoja destino via constante, no hardcodeado en el cuerpo
' ======================================================================================

Private Sub ProcesarColumnas(ByRef wb As Workbook, ByVal configNombre As String, ByVal colIndex As Long)
    Dim wsConfig   As Worksheet
    Dim wsDest     As Worksheet
    Dim ultimaFila As Long
    Dim i          As Long
    Dim nombreCol  As String
    Dim valorConf  As String
    Dim colDest    As Long
    Dim listaBorrar As New Collection
    
    Set wsConfig = ThisWorkbook.Worksheets("columnas")
    
    ' Verificar que la hoja de destino existe en el libro copia
    On Error Resume Next
    Set wsDest = wb.Worksheets(HOJA_DATOS)
    On Error GoTo 0
    
    If wsDest Is Nothing Then
        Debug.Print "  [AVISO] Hoja '" & HOJA_DATOS & "' no encontrada en el libro copia. Saltando columnas."
        Exit Sub
    End If
    
    DesprotegerHoja wsDest, PASSWORD_HOJAS
    
    ultimaFila = UltimaFilaReal(wsConfig, 2)
    Debug.Print "  [COLUMNAS] Última fila config: " & ultimaFila & " | Columna config: " & colIndex
    
    For i = 4 To ultimaFila
        nombreCol = Trim(wsConfig.Cells(i, 2).Value)
        valorConf = UCase(Trim(wsConfig.Cells(i, colIndex).Value))
        
        If nombreCol <> "" Then
            Select Case valorConf
                Case CONFIG_QUITAR
                    colDest = EncontrarColumnaPorTexto(wsDest, nombreCol)
                    If colDest > 0 Then
                        listaBorrar.Add colDest
                        Debug.Print "    [-] QUITAR columna '" & nombreCol & "' (pos " & colDest & ")"
                    Else
                        Debug.Print "    [!] No encontrada en destino: '" & nombreCol & "'"
                    End If
                Case CONFIG_MANTENER
                    ' No se hace nada, la columna se conserva tal cual
                Case ""
                    ' Celda vacía, ignorar
                Case Else
                    Debug.Print "    [?] Valor no reconocido '" & valorConf & "' en fila " & i & " para '" & nombreCol & "'"
            End Select
        End If
    Next i
    
    Debug.Print "    Total columnas a quitar: " & listaBorrar.Count
    BorrarElementos wsDest, listaBorrar, "COLUMNA"
End Sub

' ======================================================================================
' PROCESADO DE FILAS
' BUG CORREGIDO: colTextoDestino declarada al inicio del Sub, no dentro del If
' BUG CORREGIDO: colIndex no se reutiliza de columnas, siempre se busca en hoja filas
' ======================================================================================

Private Sub ProcesarFilas(ByRef wb As Workbook, ByVal configNombre As String)
    Dim wsConfig      As Worksheet
    Dim wsDest        As Worksheet
    Dim ultimaFila    As Long
    Dim i             As Long
    Dim textoFila     As String
    Dim valorConf     As String
    Dim filaDestino   As Long
    Dim listaBorrar   As New Collection
    Dim textoExtra    As String
    Dim colTexto      As Long
    Dim colExtra      As Long
    Dim colIndexFilas As Long
    Dim colTextoDestino As Long   ' Declarada aquí, no dentro del bucle
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("filas")
    Set wsDest   = wb.Worksheets(HOJA_PREGUNTAS)
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        Debug.Print "  [AVISO] Hoja 'filas' no encontrada en el libro original. Saltando."
        Exit Sub
    End If
    
    If wsDest Is Nothing Then
        Debug.Print "  [AVISO] Hoja '" & HOJA_PREGUNTAS & "' no encontrada en el libro copia. Saltando."
        Exit Sub
    End If
    
    DesprotegerHoja wsDest, PASSWORD_HOJAS
    
    ' Buscar columna del cliente en la hoja filas (independiente de la hoja columnas)
    colIndexFilas = BuscarIndiceConfiguracion(wsConfig, configNombre)
    If colIndexFilas = 0 Then
        Debug.Print "  [ERROR] '" & configNombre & "' no encontrado en hoja 'filas'. Saltando."
        Exit Sub
    End If
    
    colTexto = DetectarColumnaTextos(wsConfig, 3)
    colExtra = colIndexFilas + 5
    ultimaFila = UltimaFilaReal(wsConfig, colTexto)
    
    Debug.Print "  [FILAS] Última fila: " & ultimaFila & " | ColTexto: " & colTexto & _
                " | ColConf: " & colIndexFilas & " | ColExtra: " & colExtra
    
    For i = 3 To ultimaFila
        textoFila  = Trim(wsConfig.Cells(i, colTexto).Value)
        valorConf  = UCase(Trim(wsConfig.Cells(i, colIndexFilas).Value))
        textoExtra = Trim(wsConfig.Cells(i, colExtra).Value)
        
        If Len(textoFila) > 5 Then
            filaDestino = BuscarFilaPorTexto(wsDest, textoFila)
            
            If filaDestino > 0 Then
                Select Case valorConf
                    Case CONFIG_QUITAR
                        listaBorrar.Add filaDestino
                        Debug.Print "    [-] QUITAR fila " & filaDestino & " ('" & Left(textoFila, 35) & "...')"
                    
                    Case CONFIG_MANTENER
                        If textoExtra <> "" Then
                            colTextoDestino = EncontrarColumnaConTextoLargo(wsDest, filaDestino)
                            If colTextoDestino > 0 Then
                                wsDest.Cells(filaDestino, colTextoDestino + 1).Value = textoExtra
                                Debug.Print "    [+] Texto extra añadido en col " & (colTextoDestino + 1) & _
                                            " para fila " & filaDestino
                            Else
                                Debug.Print "    [!] No se encontró columna de texto largo en fila " & filaDestino
                            End If
                        Else
                            Debug.Print "    [=] MANTENER sin cambios, fila " & filaDestino
                        End If
                    
                    Case ""
                        ' Celda vacía, ignorar
                    
                    Case Else
                        Debug.Print "    [?] Valor no reconocido '" & valorConf & "' en fila config " & i
                End Select
            Else
                Debug.Print "    [!] No encontrado en destino: '" & Left(textoFila, 40) & "'"
            End If
        End If
    Next i
    
    Debug.Print "    Total filas a quitar: " & listaBorrar.Count
    BorrarElementos wsDest, listaBorrar, "FILA"
End Sub

' ======================================================================================
' FUNCIONES AUXILIARES
' ======================================================================================

' Busca en qué columna de la hoja está el nombre del cliente (fila 2 o fila 3)
Private Function BuscarIndiceConfiguracion(ByVal ws As Worksheet, ByVal nombre As String) As Long
    Dim filaEnc   As Long
    Dim c         As Long
    Dim ultimaCol As Long
    Dim hayDatos  As Boolean
    
    ' Detectar si el cliente está en fila 2; si no, probar fila 3
    hayDatos = False
    For c = 1 To 15
        If UCase(Trim(ws.Cells(2, c).Value)) = UCase(nombre) Then
            hayDatos = True
            Exit For
        End If
    Next c
    filaEnc = IIf(hayDatos, 2, 3)
    
    ultimaCol = ws.Cells(filaEnc, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ultimaCol
        If UCase(Trim(ws.Cells(filaEnc, c).Value)) = UCase(nombre) Then
            BuscarIndiceConfiguracion = c
            Exit Function
        End If
    Next c
    
    BuscarIndiceConfiguracion = 0
End Function

' Busca en las primeras 5 filas de la hoja en qué columna está un encabezado
Private Function EncontrarColumnaPorTexto(ByVal ws As Worksheet, ByVal txt As String) As Long
    Dim r As Long
    Dim c As Long
    For r = 1 To 5
        For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            If UCase(Trim(ws.Cells(r, c).Value)) = UCase(txt) Then
                EncontrarColumnaPorTexto = c
                Exit Function
            End If
        Next c
    Next r
    EncontrarColumnaPorTexto = 0
End Function

' Detecta la columna de la fila dada que tiene el texto más largo (referencia de filas)
Private Function DetectarColumnaTextos(ByVal ws As Worksheet, ByVal fila As Long) As Long
    Dim col       As Long
    Dim maxLen    As Long
    Dim colMax    As Long
    Dim lenActual As Long
    Dim ultimaCol As Long
    
    maxLen    = 0
    colMax    = 6                ' Valor por defecto si no se encuentra nada mejor
    ultimaCol = ws.Cells(fila, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To ultimaCol
        lenActual = Len(Trim(ws.Cells(fila, col).Value))
        If lenActual > maxLen And lenActual > 20 Then
            maxLen = lenActual
            colMax = col
        End If
    Next col
    
    DetectarColumnaTextos = colMax
End Function

' Devuelve la última fila con contenido en una columna dada (usa toda la hoja)
Private Function UltimaFilaReal(ByVal ws As Worksheet, ByVal col As Long) As Long
    UltimaFilaReal = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' Busca en qué fila de la hoja aparece un fragmento de texto, recorriendo todas las columnas reales
Private Function BuscarFilaPorTexto(ByVal ws As Worksheet, ByVal txt As String) As Long
    Dim r          As Long
    Dim c          As Long
    Dim ultimaFila As Long
    Dim ultimaCol  As Long
    Dim fragmento  As String
    Dim textoCelda As String
    
    ' Rango real de la hoja
    ultimaFila = 1
    ultimaCol  = ws.UsedRange.Columns.Count
    
    Dim tempFila As Long
    For c = 1 To ultimaCol
        tempFila = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If tempFila > ultimaFila Then ultimaFila = tempFila
    Next c
    
    fragmento = Trim(IIf(Len(txt) > 20, Left(txt, 20), txt))
    
    For r = 1 To ultimaFila
        For c = 1 To ultimaCol
            On Error Resume Next
            textoCelda = Trim(ws.Cells(r, c).Value)
            On Error GoTo 0
            If Len(textoCelda) > 10 And InStr(1, textoCelda, fragmento, vbTextCompare) > 0 Then
                BuscarFilaPorTexto = r
                Exit Function
            End If
        Next c
    Next r
    
    BuscarFilaPorTexto = 0
End Function

' Devuelve la columna con el texto más largo en una fila (para saber dónde añadir texto extra)
Private Function EncontrarColumnaConTextoLargo(ByVal ws As Worksheet, ByVal fila As Long) As Long
    Dim col       As Long
    Dim maxLen    As Long
    Dim colMax    As Long
    Dim lenActual As Long
    Dim ultimaCol As Long
    
    maxLen    = 0
    colMax    = 0
    ultimaCol = ws.Cells(fila, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To ultimaCol
        On Error Resume Next
        lenActual = Len(Trim(ws.Cells(fila, col).Value))
        On Error GoTo 0
        If lenActual > maxLen And lenActual > 20 Then
            maxLen = lenActual
            colMax = col
        End If
    Next col
    
    EncontrarColumnaConTextoLargo = colMax
End Function

' Borra filas o columnas ordenando los índices de mayor a menor para no desplazar posiciones
Private Sub BorrarElementos(ByVal ws As Worksheet, ByVal indices As Collection, ByVal tipo As String)
    If indices.Count = 0 Then
        Debug.Print "      [INFO] Nada que quitar."
        Exit Sub
    End If
    
    Dim arr() As Long
    Dim i     As Long
    Dim j     As Long
    Dim temp  As Long
    ReDim arr(1 To indices.Count)
    
    For i = 1 To indices.Count
        arr(i) = indices(i)
    Next i
    
    ' Ordenar descendente (burbuja simple — válido para colecciones pequeñas)
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                temp   = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    For i = 1 To UBound(arr)
        If tipo = "COLUMNA" Then
            ws.Columns(arr(i)).Delete
        Else
            ws.Rows(arr(i)).Delete
        End If
        Debug.Print "        [X] " & tipo & " " & arr(i) & " eliminada"
    Next i
End Sub

' ======================================================================================
' UTILIDADES DE ENTORNO
' ======================================================================================

Private Sub GestionarEntorno(ByVal activar As Boolean)
    Application.ScreenUpdating = Not activar
    Application.DisplayAlerts  = Not activar
    Application.EnableEvents   = Not activar
    Application.Calculation    = IIf(activar, xlCalculationManual, xlCalculationAutomatic)
End Sub
