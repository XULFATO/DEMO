Attribute VB_Name = "Sistema_Crear_Excels_Separados"
'===============================================================================
' SISTEMA DE CREACION DE EXCELS SEPARADOS
' Crea un Excel completo por cada configuracion (BOB, BING, BANG)
' Borra las columnas marcadas con "no" en cada Excel
'===============================================================================

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
    
    ' Verificar hoja de configuracion (ahora se llama "columnas")
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("columnas")
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        MsgBox "ERROR: No se encuentra la hoja 'columnas'" & vbCrLf & _
               "(antes se llamaba Hoja1)", vbCritical
        Exit Sub
    End If
    
    Debug.Print String(70, "=")
    Debug.Print "INICIO CREACION DE EXCELS SEPARADOS"
    Debug.Print String(70, "=")
    
    ' Ruta de destino
    rutaBase = "C:\CLIENTES\PRUEBAS\BP\"
    
    ' Crear carpeta si no existe
    If Not CrearCarpeta(rutaBase) Then
        MsgBox "ERROR: No se pudo crear la carpeta: " & rutaBase, vbCritical
        Exit Sub
    End If
    
    ' Nombre del archivo original (sin extension)
    nombreOriginal = Replace(ThisWorkbook.Name, ".xlsm", "")
    nombreOriginal = Replace(nombreOriginal, ".xlsx", "")
    nombreOriginal = Replace(nombreOriginal, ".xls", "")
    
    Debug.Print "Nombre base: " & nombreOriginal
    Debug.Print "Ruta destino: " & rutaBase
    Debug.Print ""
    
    ' Detectar configuraciones
    Set configuraciones = DetectarConfiguraciones(wsConfig)
    
    If configuraciones.Count = 0 Then
        MsgBox "No se encontraron configuraciones en hoja 'columnas'", vbInformation
        Exit Sub
    End If
    
    Debug.Print "Configuraciones encontradas: " & configuraciones.Count
    Debug.Print ""
    
    ' Crear un Excel por cada configuracion
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each config In configuraciones
        CrearExcelParaConfiguracion wsConfig, CStr(config), rutaBase, nombreOriginal
        totalExcels = totalExcels + 1
    Next config
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Debug.Print ""
    Debug.Print String(70, "=")
    Debug.Print "PROCESO COMPLETADO"
    Debug.Print "Excels creados: " & totalExcels
    Debug.Print "Ubicacion: " & rutaBase
    Debug.Print String(70, "=")
    
    ' Abrir carpeta
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("Proceso completado!" & vbCrLf & vbCrLf & _
                       "Excels creados: " & totalExcels & vbCrLf & _
                       "Ubicacion: " & rutaBase & vbCrLf & vbCrLf & _
                       "Abrir carpeta?", _
                       vbYesNo + vbInformation, "Completado")
    
    If respuesta = vbYes Then
        Shell "explorer.exe " & rutaBase, vbNormalFocus
    End If
End Sub


' ============================================================================
' MODULO 1: CREAR CARPETA
' ============================================================================

Private Function CrearCarpeta(ByVal ruta As String) As Boolean
    On Error Resume Next
    
    ' Intentar crear carpeta (si ya existe, no da error)
    MkDir "C:\CLIENTES"
    MkDir "C:\CLIENTES\PRUEBAS"
    MkDir "C:\CLIENTES\PRUEBAS\BP"
    
    ' Verificar que existe
    If Dir(ruta, vbDirectory) <> "" Then
        CrearCarpeta = True
        Debug.Print "[OK] Carpeta verificada: " & ruta
    Else
        CrearCarpeta = False
        Debug.Print "[ERROR] No se pudo crear: " & ruta
    End If
    
    On Error GoTo 0
End Function


' ============================================================================
' MODULO 2: DETECTAR CONFIGURACIONES
' ============================================================================

Private Function DetectarConfiguraciones(ByVal wsConfig As Worksheet) As Collection
    Dim configs As Collection
    Dim col As Long
    Dim nombreConfig As String
    Dim filaConfig As Long
    Dim ultimaColumna As Long
    
    Set configs = New Collection
    filaConfig = 3
    
    ultimaColumna = wsConfig.Cells(filaConfig, wsConfig.Columns.Count).End(xlToLeft).Column
    
    Debug.Print "[DETECTANDO CONFIGURACIONES]"
    Debug.Print String(70, "-")
    
    For col = 3 To ultimaColumna
        nombreConfig = Trim(wsConfig.Cells(filaConfig, col).Value)
        
        If nombreConfig <> "" Then
            configs.Add nombreConfig
            Debug.Print "  [OK] " & nombreConfig & " (columna " & col & ")"
        End If
    Next col
    
    Debug.Print String(70, "-")
    
    Set DetectarConfiguraciones = configs
End Function


' ============================================================================
' MODULO 3: CREAR EXCEL PARA CONFIGURACION
' ============================================================================

Private Sub CrearExcelParaConfiguracion(ByVal wsConfig As Worksheet, _
                                         ByVal nombreConfig As String, _
                                         ByVal rutaBase As String, _
                                         ByVal nombreOriginal As String)
    
    Dim wbNuevo As Workbook
    Dim rutaCompleta As String
    Dim colConfig As Long
    Dim columnasABorrar As Collection
    Dim filasABorrar As Collection
    Dim wsColumnas As Worksheet
    Dim wsFilas As Worksheet
    
    Debug.Print ""
    Debug.Print "[PROCESANDO: " & nombreConfig & "]"
    Debug.Print String(70, "-")
    
    ' Nombre del nuevo Excel
    rutaCompleta = rutaBase & nombreOriginal & "_" & nombreConfig & ".xlsx"
    Debug.Print "  Archivo: " & nombreOriginal & "_" & nombreConfig & ".xlsx"
    
    ' 1. Copiar TODO el libro actual
    ThisWorkbook.SaveCopyAs rutaCompleta
    Debug.Print "  [OK] Copia creada"
    
    ' 2. Abrir la copia
    Set wbNuevo = Workbooks.Open(rutaCompleta)
    Debug.Print "  [OK] Copia abierta"
    
    ' ==================================================================
    ' PARTE 1: PROCESAR COLUMNAS (hoja "columnas")
    ' ==================================================================
    
    On Error Resume Next
    Set wsColumnas = wbNuevo.Worksheets("columnas")
    On Error GoTo 0
    
    If Not wsColumnas Is Nothing Then
        Debug.Print "  [COLUMNAS] Procesando..."
        
        colConfig = BuscarColumnaConfiguracion(wsColumnas, nombreConfig)
        
        If colConfig > 0 Then
            Debug.Print "    Columna configuracion: " & colConfig
            
            Set columnasABorrar = LeerColumnasABorrar(wsColumnas, _
                                                       wbNuevo.Worksheets("FuncionFiltar"), _
                                                       colConfig)
            
            Debug.Print "    Columnas a borrar: " & columnasABorrar.Count
            
            If columnasABorrar.Count > 0 Then
                BorrarColumnas wbNuevo.Worksheets("FuncionFiltar"), columnasABorrar
                Debug.Print "    [OK] Columnas borradas"
            End If
        End If
    Else
        Debug.Print "  [AVISO] No se encontro hoja 'columnas'"
    End If
    
    ' ==================================================================
    ' PARTE 2: PROCESAR FILAS (hoja "filas")
    ' ==================================================================
    
    On Error Resume Next
    Set wsFilas = wbNuevo.Worksheets("filas")
    On Error GoTo 0
    
    Debug.Print "  [FILAS] Procesando..."
    
    If wsFilas Is Nothing Then
        Debug.Print "    [AVISO] No se encontro hoja 'filas' en el Excel copiado"
    Else
        Debug.Print "    [OK] Hoja 'filas' encontrada"
        Debug.Print "    [DEBUG] Buscando configuracion '" & nombreConfig & "'..."
        
        colConfig = BuscarColumnaConfiguracion(wsFilas, nombreConfig)
        
        Debug.Print "    [DEBUG] Resultado busqueda columna: " & colConfig
        
        If colConfig > 0 Then
            Debug.Print "    [OK] Columna configuracion encontrada: " & colConfig
            
            ' Verificar si existe hoja TEXOENFILADOS
            Dim wsTEXOENFILADOS As Worksheet
            On Error Resume Next
            Set wsTEXOENFILADOS = wbNuevo.Worksheets("TEXOENFILADOS")
            On Error GoTo 0
            
            If wsTEXOENFILADOS Is Nothing Then
                Debug.Print "    [ERROR] No existe hoja 'TEXOENFILADOS' en el Excel copiado"
            Else
                Debug.Print "    [OK] Hoja 'TEXOENFILADOS' encontrada"
                Debug.Print "    [DEBUG] Llamando a LeerFilasABorrar..."
                
                Set filasABorrar = LeerFilasABorrar(wsFilas, wsTEXOENFILADOS, colConfig)
                
                Debug.Print "    [DEBUG] Filas a borrar: " & filasABorrar.Count
                
                If filasABorrar.Count > 0 Then
                    BorrarFilas wsTEXOENFILADOS, filasABorrar
                    Debug.Print "    [OK] Filas borradas"
                Else
                    Debug.Print "    [INFO] No hay filas para borrar (todas tienen SI)"
                End If
            End If
        Else
            Debug.Print "    [ERROR] No se encontro columna para configuracion: " & nombreConfig
        End If
    End If
    
    ' ==================================================================
    ' PARTE 3: ELIMINAR HOJAS DE CONFIGURACION
    ' ==================================================================
    
    Debug.Print "  [LIMPIEZA] Eliminando hojas de configuracion..."
    
    ' Eliminar hoja "columnas"
    On Error Resume Next
    Application.DisplayAlerts = False
    wbNuevo.Worksheets("columnas").Delete
    If Err.Number = 0 Then
        Debug.Print "    [X] Hoja 'columnas' eliminada"
    End If
    Err.Clear
    
    ' Eliminar hoja "filas"
    wbNuevo.Worksheets("filas").Delete
    If Err.Number = 0 Then
        Debug.Print "    [X] Hoja 'filas' eliminada"
    End If
    Err.Clear
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' ==================================================================
    ' FINALIZAR
    ' ==================================================================
    
    ' Guardar y cerrar
    wbNuevo.Save
    wbNuevo.Close SaveChanges:=True
    Debug.Print "  [OK] Excel guardado: " & nombreConfig
    Debug.Print String(70, "-")
End Sub


' ============================================================================
' MODULO 4: BUSCAR COLUMNA DE CONFIGURACION
' ============================================================================

Private Function BuscarColumnaConfiguracion(ByVal ws As Worksheet, _
                                             ByVal nombreConfig As String) As Long
    Dim col As Long
    Dim ultimaColumna As Long
    Dim filaEncabezado As Long
    
    ' DETECTAR en qué fila están los encabezados (BOB, BING, BANG)
    ' Buscar en las primeras 5 filas
    filaEncabezado = 0
    For filaEncabezado = 1 To 5
        If InStr(1, ws.Cells(filaEncabezado, 2).Value, nombreConfig, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(filaEncabezado, 3).Value, nombreConfig, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(filaEncabezado, 4).Value, nombreConfig, vbTextCompare) > 0 Then
            Exit For
        End If
    Next filaEncabezado
    
    If filaEncabezado > 5 Then filaEncabezado = 2  ' Por defecto fila 2
    
    Debug.Print "      [DEBUG] Buscando '" & nombreConfig & "' en fila " & filaEncabezado
    
    ' Detectar última columna con datos
    ultimaColumna = ws.Cells(filaEncabezado, ws.Columns.Count).End(xlToLeft).Column
    
    ' Buscar la configuración en esa fila
    For col = 1 To ultimaColumna
        Dim valorCelda As String
        valorCelda = Trim(ws.Cells(filaEncabezado, col).Value)
        
        Debug.Print "        Col " & col & " fila " & filaEncabezado & ": '" & valorCelda & "'"
        
        If UCase(valorCelda) = UCase(nombreConfig) Then
            Debug.Print "      [OK] '" & nombreConfig & "' encontrado en columna " & col
            BuscarColumnaConfiguracion = col
            Exit Function
        End If
    Next col
    
    Debug.Print "      [ERROR] '" & nombreConfig & "' NO encontrado"
    BuscarColumnaConfiguracion = 0
End Function


' ============================================================================
' MODULO 5: LEER COLUMNAS A BORRAR (las que tienen "no")
' ============================================================================

Private Function LeerColumnasABorrar(ByVal wsConfig As Worksheet, _
                                      ByVal wsOrigen As Worksheet, _
                                      ByVal colConfig As Long) As Collection
    
    Dim columnas As Collection
    Dim fila As Long
    Dim ultimaFila As Long
    Dim nombreColumna As String
    Dim valor As String
    Dim numColOrigen As Long
    
    Set columnas = New Collection
    
    ultimaFila = wsConfig.Cells(wsConfig.Rows.Count, 2).End(xlUp).Row
    
    For fila = 4 To ultimaFila
        nombreColumna = Trim(wsConfig.Cells(fila, 2).Value)
        
        If nombreColumna <> "" Then
            valor = Trim(wsConfig.Cells(fila, colConfig).Value)
            
            ' Si es "no" o "NO", borrar esta columna
            If UCase(valor) = "NO" Then
                numColOrigen = BuscarColumnaEnOrigen(wsOrigen, nombreColumna)
                
                If numColOrigen > 0 Then
                    columnas.Add numColOrigen
                    Debug.Print "    [-] Marcar para borrar: " & nombreColumna & " (col " & numColOrigen & ")"
                End If
            End If
        End If
    Next fila
    
    Set LeerColumnasABorrar = columnas
End Function


' ============================================================================
' MODULO 6: BUSCAR COLUMNA EN ORIGEN
' ============================================================================

Private Function BuscarColumnaEnOrigen(ByVal ws As Worksheet, _
                                        ByVal nombreBuscado As String) As Long
    
    Dim col As Long
    Dim ultimaCol As Long
    Dim filaEncabezado As Long
    
    ' Detectar fila de encabezado
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
    
    BuscarColumnaEnOrigen = 0
End Function


' ============================================================================
' MODULO 7: BORRAR COLUMNAS
' Borra columnas de mayor a menor para evitar desplazamientos
' ============================================================================

Private Sub BorrarColumnas(ByVal ws As Worksheet, ByVal columnasABorrar As Collection)
    Dim arrColumnas() As Long
    Dim i As Long
    Dim j As Long
    Dim temp As Long
    Dim numCol As Variant
    
    If columnasABorrar.Count = 0 Then Exit Sub
    
    ' Convertir collection a array
    ReDim arrColumnas(1 To columnasABorrar.Count)
    i = 1
    For Each numCol In columnasABorrar
        arrColumnas(i) = CLng(numCol)
        i = i + 1
    Next numCol
    
    ' Ordenar de MAYOR a MENOR (para borrar sin desplazar)
    For i = 1 To UBound(arrColumnas) - 1
        For j = i + 1 To UBound(arrColumnas)
            If arrColumnas(i) < arrColumnas(j) Then
                temp = arrColumnas(i)
                arrColumnas(i) = arrColumnas(j)
                arrColumnas(j) = temp
            End If
        Next j
    Next i
    
    ' Borrar columnas de mayor a menor
    For i = 1 To UBound(arrColumnas)
        ws.Columns(arrColumnas(i)).Delete
        Debug.Print "      [X] Columna " & arrColumnas(i) & " borrada"
    Next i
End Sub


' ============================================================================
' MODULO 8: LEER FILAS A BORRAR (las que tienen "no")
' ============================================================================

Private Function LeerFilasABorrar(ByVal wsConfig As Worksheet, _
                                   ByVal wsOrigen As Worksheet, _
                                   ByVal colConfig As Long) As Collection
    
    Dim filas As Collection
    Dim fila As Long
    Dim ultimaFila As Long
    Dim textoLinea As String
    Dim valor As String
    Dim numFilaOrigen As Long
    Dim colTexto As Long
    Dim col As Long
    Dim filaInicio As Long
    
    Set filas = New Collection
    
    Debug.Print "    [DEBUG-1] Iniciando lectura de filas a borrar"
    Debug.Print "    [DEBUG-2] Hoja config: " & wsConfig.Name
    Debug.Print "    [DEBUG-3] Hoja origen: " & wsOrigen.Name
    Debug.Print "    [DEBUG-4] Columna configuracion: " & colConfig
    
    ' DETECTAR fila de inicio (donde empiezan los datos, no los encabezados)
    ' Buscar primera fila con "NO" o "SI" en la columna de configuración
    filaInicio = 3  ' Por defecto fila 3
    For fila = 2 To 10
        Dim valorTest As String
        valorTest = UCase(Trim(wsConfig.Cells(fila, colConfig).Value))
        If valorTest = "NO" Or valorTest = "SI" Then
            filaInicio = fila
            Debug.Print "    [DEBUG-5] Primera fila de datos detectada: " & filaInicio
            Exit For
        End If
    Next fila
    
    ' DETECTAR columna con textos (la que tiene contenido largo, tipo F)
    Debug.Print "    [DEBUG-6] Buscando columna con textos largos en fila " & filaInicio & "..."
    
    colTexto = 0
    Dim maxLen As Integer
    maxLen = 0
    
    For col = 1 To 20
        Dim valorCelda As String
        valorCelda = Trim(wsConfig.Cells(filaInicio, col).Value)
        
        If Len(valorCelda) > maxLen And Len(valorCelda) > 20 Then
            maxLen = Len(valorCelda)
            colTexto = col
            Debug.Print "      Col " & col & " tiene " & Len(valorCelda) & " caracteres: '" & Left(valorCelda, 50) & "...'"
        End If
    Next col
    
    If colTexto = 0 Then
        Debug.Print "    [ERROR] NO se encontro columna con textos largos"
        Debug.Print "    [INFO] Mostrando todas las celdas de fila " & filaInicio & ":"
        For col = 1 To 10
            Debug.Print "      Col " & col & " (" & Len(wsConfig.Cells(filaInicio, col).Value) & " chars): '" & Left(wsConfig.Cells(filaInicio, col).Value, 50) & "'"
        Next col
        Set LeerFilasABorrar = filas
        Exit Function
    End If
    
    Debug.Print "    [DEBUG-7] Columna de textos detectada: " & colTexto & " (columna con mas caracteres)"
    
    ' Detectar ultima fila
    ultimaFila = wsConfig.Cells(wsConfig.Rows.Count, colTexto).End(xlUp).Row
    Debug.Print "    [DEBUG-8] Ultima fila en columna " & colTexto & ": " & ultimaFila
    
    ' Recorrer filas
    Debug.Print "    [DEBUG-9] Recorriendo desde fila " & filaInicio & " hasta " & ultimaFila
    
    Dim contadorProcesadas As Integer
    contadorProcesadas = 0
    
    For fila = filaInicio To ultimaFila
        textoLinea = Trim(wsConfig.Cells(fila, colTexto).Value)
        
        If Len(textoLinea) > 5 Then  ' Solo procesar si tiene contenido
            contadorProcesadas = contadorProcesadas + 1
            valor = Trim(wsConfig.Cells(fila, colConfig).Value)
            
            Debug.Print "    [DEBUG-10] Fila " & fila & ":"
            Debug.Print "      Texto (col " & colTexto & "): '" & Left(textoLinea, 60) & "...'"
            Debug.Print "      Valor config (col " & colConfig & "): '" & valor & "'"
            Debug.Print "      Es NO? " & (UCase(valor) = "NO")
            
            If UCase(valor) = "NO" Then
                Debug.Print "      --> SI es NO, buscando en origen..."
                
                numFilaOrigen = BuscarFilaPorTexto(wsOrigen, textoLinea)
                
                If numFilaOrigen > 0 Then
                    filas.Add numFilaOrigen
                    Debug.Print "      --> MARCADA para borrar fila " & numFilaOrigen & " de TEXOENFILADOS"
                Else
                    Debug.Print "      --> NO encontrada en TEXOENFILADOS"
                End If
            Else
                Debug.Print "      --> Es SI, se MANTIENE"
            End If
        End If
    Next fila
    
    Debug.Print "    [DEBUG-11] Filas procesadas: " & contadorProcesadas
    Debug.Print "    [DEBUG-12] Filas marcadas para borrar: " & filas.Count
    
    Set LeerFilasABorrar = filas
End Function


' ============================================================================
' MODULO 9: BUSCAR FILA POR TEXTO
' Busca una fila que contenga el texto especificado
' ============================================================================

Private Function BuscarFilaPorTexto(ByVal ws As Worksheet, _
                                     ByVal textoBuscado As String) As Long
    
    Dim fila As Long
    Dim ultimaFila As Long
    Dim textoFila As String
    Dim col As Long
    
    Debug.Print "        [INFO-BUSQUEDA] Buscando en:"
    Debug.Print "          Libro: " & ws.Parent.Name
    Debug.Print "          Hoja: " & ws.Name
    
    ' DETECTAR ULTIMA FILA CORRECTAMENTE (buscar en TODAS las columnas)
    ultimaFila = 1
    For col = 1 To 20
        Dim ultimaFilaCol As Long
        ultimaFilaCol = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If ultimaFilaCol > ultimaFila Then
            ultimaFila = ultimaFilaCol
        End If
    Next col
    
    Debug.Print "          Ultima fila (detectada): " & ultimaFila
    
    ' Mostrar primeras 5 filas para verificar contenido
    Debug.Print "          Muestra primeras 5 filas:"
    For fila = 1 To 5
        Dim textoMuestra As String
        textoMuestra = ""
        For col = 1 To 4
            textoMuestra = textoMuestra & Left(ws.Cells(fila, col).Value, 20) & " | "
        Next col
        Debug.Print "            Fila " & fila & ": " & textoMuestra
    Next fila
    
    ' Limpiar texto buscado
    textoBuscado = Trim(textoBuscado)
    
    ' Usar solo las primeras palabras significativas (15-20 caracteres)
    Dim textoBuscar As String
    If Len(textoBuscado) > 20 Then
        textoBuscar = Left(textoBuscado, 20)
    Else
        textoBuscar = textoBuscado
    End If
    
    ' Limpiar caracteres especiales para búsqueda más flexible
    textoBuscar = Replace(textoBuscar, "  ", " ")  ' Dobles espacios
    textoBuscar = Trim(textoBuscar)
    
    Debug.Print "        [BUSCAR] '" & textoBuscar & "...'"
    
    ' Buscar en todas las filas y columnas
    For fila = 1 To ultimaFila
        For col = 1 To 20
            On Error Resume Next
            textoFila = Trim(ws.Cells(fila, col).Value)
            On Error GoTo 0
            
            ' Limpiar también el texto de la celda
            textoFila = Replace(textoFila, "  ", " ")
            textoFila = Trim(textoFila)
            
            ' Buscar si contiene el texto (case insensitive)
            If Len(textoFila) > 10 And InStr(1, textoFila, textoBuscar, vbTextCompare) > 0 Then
                Debug.Print "        [OK] Encontrado en fila " & fila & ", col " & col
                BuscarFilaPorTexto = fila
                Exit Function
            End If
        Next col
    Next fila
    
    ' Si no encontró, intentar con menos caracteres (primeros 15)
    If Len(textoBuscado) > 15 Then
        textoBuscar = Left(textoBuscado, 15)
        textoBuscar = Replace(textoBuscar, "  ", " ")
        textoBuscar = Trim(textoBuscar)
        
        Debug.Print "        [REINTENTO] Buscando con menos texto: '" & textoBuscar & "...'"
        
        For fila = 1 To ultimaFila
            For col = 1 To 20
                On Error Resume Next
                textoFila = Trim(ws.Cells(fila, col).Value)
                On Error GoTo 0
                
                textoFila = Replace(textoFila, "  ", " ")
                textoFila = Trim(textoFila)
                
                If Len(textoFila) > 10 And InStr(1, textoFila, textoBuscar, vbTextCompare) > 0 Then
                    Debug.Print "        [OK] Encontrado en fila " & fila & ", col " & col & " (con texto corto)"
                    BuscarFilaPorTexto = fila
                    Exit Function
                End If
            Next col
        Next fila
    End If
    
    Debug.Print "        [!] NO encontrado: " & textoBuscar
    BuscarFilaPorTexto = 0
End Function


' ============================================================================
' MODULO 10: BORRAR FILAS
' Borra filas de mayor a menor para evitar desplazamientos
' ============================================================================

Private Sub BorrarFilas(ByVal ws As Worksheet, ByVal filasABorrar As Collection)
    Dim arrFilas() As Long
    Dim i As Long
    Dim j As Long
    Dim temp As Long
    Dim numFila As Variant
    
    If filasABorrar.Count = 0 Then Exit Sub
    
    ' Convertir collection a array
    ReDim arrFilas(1 To filasABorrar.Count)
    i = 1
    For Each numFila In filasABorrar
        arrFilas(i) = CLng(numFila)
        i = i + 1
    Next numFila
    
    ' Ordenar de MAYOR a MENOR
    For i = 1 To UBound(arrFilas) - 1
        For j = i + 1 To UBound(arrFilas)
            If arrFilas(i) < arrFilas(j) Then
                temp = arrFilas(i)
                arrFilas(i) = arrFilas(j)
                arrFilas(j) = temp
            End If
        Next j
    Next i
    
    ' Borrar filas de mayor a menor
    For i = 1 To UBound(arrFilas)
        ws.Rows(arrFilas(i)).Delete
        Debug.Print "      [X] Fila " & arrFilas(i) & " borrada"
    Next i
End Sub


' ============================================================================
' FUNCION EXTRA: LIMPIAR EXCELS CREADOS
' ============================================================================

Public Sub LimpiarExcelsCreados()
    Dim rutaBase As String
    Dim archivo As String
    Dim contador As Integer
    
    rutaBase = "C:\CLIENTES\PRUEBAS\BP\"
    
    If Dir(rutaBase, vbDirectory) = "" Then
        MsgBox "La carpeta no existe: " & rutaBase, vbInformation
        Exit Sub
    End If
    
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("Eliminar todos los archivos .xlsx de:" & vbCrLf & _
                       rutaBase & vbCrLf & vbCrLf & _
                       "Continuar?", vbYesNo + vbQuestion, "Confirmar")
    
    If respuesta <> vbYes Then Exit Sub
    
    archivo = Dir(rutaBase & "*.xlsx")
    
    Do While archivo <> ""
        On Error Resume Next
        Kill rutaBase & archivo
        If Err.Number = 0 Then
            contador = contador + 1
            Debug.Print "[X] Eliminado: " & archivo
        End If
        On Error GoTo 0
        archivo = Dir
    Loop
    
    MsgBox "Archivos eliminados: " & contador, vbInformation
End Sub


' ============================================================================
' FUNCION TEST: VERIFICAR LECTURA DE FILAS
' ============================================================================

Public Sub TestLecturaFilas()
    Dim wsFilas As Worksheet
    Dim wsTexto As Worksheet
    Dim fila As Long
    Dim ultimaFila As Long
    Dim textoLinea As String
    Dim numFilaEncontrada As Long
    Dim reporte As String
    
    ' Verificar hojas
    On Error Resume Next
    Set wsFilas = ThisWorkbook.Worksheets("filas")
    Set wsTexto = ThisWorkbook.Worksheets("TEXOENFILADOS")
    On Error GoTo 0
    
    If wsFilas Is Nothing Then
        MsgBox "No se encuentra hoja 'filas'", vbCritical
        Exit Sub
    End If
    
    If wsTexto Is Nothing Then
        MsgBox "No se encuentra hoja 'TEXOENFILADOS'", vbCritical
        Exit Sub
    End If
    
    reporte = "TEST BUSQUEDA DE FILAS" & vbCrLf
    reporte = reporte & String(70, "=") & vbCrLf & vbCrLf
    
    ultimaFila = wsFilas.Cells(wsFilas.Rows.Count, 6).End(xlUp).Row
    
    reporte = reporte & "Textos a buscar (columna F de 'filas'):" & vbCrLf
    reporte = reporte & String(70, "-") & vbCrLf
    
    For fila = 3 To ultimaFila
        textoLinea = Trim(wsFilas.Cells(fila, 6).Value)
        
        If textoLinea <> "" Then
            reporte = reporte & "Fila " & fila & ": " & Left(textoLinea, 50) & vbCrLf
            
            ' Intentar buscar
            numFilaEncontrada = BuscarFilaPorTexto(wsTexto, textoLinea)
            
            If numFilaEncontrada > 0 Then
                reporte = reporte & "  -> ENCONTRADO en fila " & numFilaEncontrada & " de TEXOENFILADOS" & vbCrLf
            Else
                reporte = reporte & "  -> NO ENCONTRADO" & vbCrLf
            End If
            reporte = reporte & vbCrLf
        End If
    Next fila
    
    Debug.Print reporte
    MsgBox "Test completado. Ver Ventana Inmediato (Ctrl+G)", vbInformation
End Sub
