Attribute VB_Name = "Módulo2"

Option Explicit

' ======================================================================================
' CONSTANTES GLOBALES DE MÓDULO
' ======================================================================================

Private Const PASSWORD_HOJAS   As String = "ADP"
Private Const CONFIG_MANTENER  As String = "MANTENER"
Private Const CONFIG_QUITAR    As String = "QUITAR"
Private Const UNIDAD_RED       As String = "O:\"
Private Const RUTA_POR_DEFECTO As String = "O:\CLIENTES\PRUEBAS\BP\"
Private Const HOJA_DATOS       As String = "Analisis conceptos BOB"
Private Const HOJA_PREGUNTAS   As String = "Preguntas generales"

' Nombres de hojas de configuración — TRIM+UCASE aplicado al comparar
Private Const HOJA_CONFIG_COL  As String = "columnas"
Private Const HOJA_CONFIG_FIL  As String = "filas"

' Nombres de clientes válidos — no se consideran literales extraños
Private Const CLIENTE_BOB      As String = "BOB"
Private Const CLIENTE_CELERGO  As String = "CELERGO"

' ======================================================================================
' BOTONES PRINCIPALES
' ======================================================================================

Public Sub GenerarBOB()
    GenerarArchivoIndividual CLIENTE_BOB
End Sub

Public Sub GenerarCELERGO()
    GenerarArchivoIndividual CLIENTE_CELERGO
End Sub

Public Sub GenerarBOByCELERGO()
    Dim clientes(1) As String
    clientes(0) = CLIENTE_BOB
    clientes(1) = CLIENTE_CELERGO

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
    Dim nombreFinal As String   ' Nombre completo del fichero resuelto por el diálogo
    Dim indiceCol   As Long

    Debug.Print String(80, "=")
    Debug.Print "INICIO PROCESO PARA: " & clienteNombre
    Debug.Print String(80, "=")

    ' 1. Validar hoja de configuración de columnas
    On Error Resume Next
    Set wsConfigCol = HojaConfiguracion(HOJA_CONFIG_COL)
    On Error GoTo 0

    If wsConfigCol Is Nothing Then
        MsgBox "Error crítico: no se encontró la hoja '" & HOJA_CONFIG_COL & "'.", vbCritical
        Exit Sub
    End If

    ' 2. Validar literales no reconocidos ANTES de continuar
    '    BOB y CELERGO son valores válidos (nombres de cliente cruzados)
    '    Todo lo demás distinto de MANTENER/QUITAR se tratará como QUITAR tras avisar
    On Error GoTo ErrorValidacion
    ValidarLiteralesConfiguracion clienteNombre
    On Error GoTo 0

    ' 3. Validar que el cliente existe en la configuración
    indiceCol = BuscarIndiceConfiguracion(wsConfigCol, clienteNombre)
    If indiceCol = 0 Then
        MsgBox "Error: no se encontró la configuración para '" & clienteNombre & _
               "' en la hoja '" & HOJA_CONFIG_COL & "'.", vbExclamation
        Exit Sub
    End If
    Debug.Print "[OK] '" & clienteNombre & "' encontrado en columna: " & indiceCol

    ' 4. Nombre base del archivo sin extensión
    nombreBase = LimpiarExtension(ThisWorkbook.Name)
    Debug.Print "[OK] Nombre base: " & nombreBase

    ' 5. Determinar ruta por defecto (red o local) y calcular nombre sugerido con versión
    '    Luego mostrar el diálogo Guardar Como con todo pre-informado
    Dim rutaBase As String
    If UnidadRedConectada(UNIDAD_RED) Then
        rutaBase = RUTA_POR_DEFECTO
    Else
        rutaBase = "C:\"
        MsgBox "La unidad de red " & UNIDAD_RED & " no está disponible." & vbCrLf & _
               "Se usará una ruta local.", vbExclamation, "Unidad de red no disponible"
    End If
    If Right(rutaBase, 1) <> "\" Then rutaBase = rutaBase & "\"

    ' Calcular versión sugerida
    Dim versionSugerida As String
    versionSugerida = CalcularVersionSugerida(rutaBase, clienteNombre, nombreBase)
    ' versionSugerida puede ser "V01" (nueva) o una versión existente si el usuario
    ' confirmó sobreescribir — aquí solo calculamos la sugerencia inicial

    Dim nombreSugerido As String
    nombreSugerido = clienteNombre & "_" & nombreBase & "_" & versionSugerida

    ' 6. Mostrar diálogo Guardar Como nativo: ruta pre-abierta + nombre pre-informado
    nombreFinal = MostrarDialogoGuardarComo(rutaBase, nombreSugerido)
    If nombreFinal = "" Then
        Debug.Print "[CANCELADO] El usuario canceló el diálogo de guardado."
        Exit Sub
    End If
    Debug.Print "[OK] Fichero destino: " & nombreFinal

    ' Extraer carpeta del fichero final seleccionado por el usuario
    rutaDestino = ExtraerCarpeta(nombreFinal)

    ' 7. Optimizar rendimiento
    GestionarEntorno True

    ' 8. Generar archivo
    EjecutarGeneracion clienteNombre, indiceCol, nombreFinal

    ' 9. Restaurar entorno
    GestionarEntorno False

    Debug.Print String(80, "=")
    Debug.Print "PROCESO COMPLETADO PARA: " & clienteNombre
    Debug.Print String(80, "=")

    MsgBox "Proceso finalizado correctamente para " & clienteNombre & "." & vbCrLf & _
           "Archivo generado en:" & vbCrLf & nombreFinal, vbInformation
    Exit Sub

ErrorValidacion:
    ' ValidarLiteralesConfiguracion elevó error por cancelación del usuario
    Debug.Print "[CANCELADO] El usuario canceló tras aviso de literales."
    Resume ExitSub
ExitSub:
End Sub

' ======================================================================================
' OBTENER HOJA DE CONFIGURACIÓN CON TRIM+UCASE EN EL NOMBRE
' Busca la hoja cuyo nombre, normalizado, coincide con el buscado.
' ======================================================================================

Private Function HojaConfiguracion(ByVal nombreHoja As String) As Worksheet
    Dim ws      As Worksheet
    Dim buscado As String
    buscado = UCase(Trim(nombreHoja))

    For Each ws In ThisWorkbook.Worksheets
        If UCase(Trim(ws.Name)) = buscado Then
            Set HojaConfiguracion = ws
            Exit Function
        End If
    Next ws
    Set HojaConfiguracion = Nothing
End Function

' ======================================================================================
' VALIDACIÓN DE LITERALES EN HOJAS DE CONFIGURACIÓN
' - Aplica TRIM+UCASE a todos los valores leídos
' - Valores válidos: MANTENER, QUITAR, "" (vacío), BOB, CELERGO
' - Cualquier otro valor genera aviso; si el usuario cancela se eleva error
' ======================================================================================

Private Sub ValidarLiteralesConfiguracion(ByVal clienteNombre As String)
    Dim nombresHojas(1) As String
    nombresHojas(0) = HOJA_CONFIG_COL
    nombresHojas(1) = HOJA_CONFIG_FIL

    Dim h            As Integer
    Dim ws           As Worksheet
    Dim colIdx       As Long
    Dim ultimaFila   As Long
    Dim filaInicio   As Long
    Dim i            As Long
    Dim valor        As String
    Dim extraños()   As String
    Dim numExtranios As Long
    Dim msg          As String

    ReDim extraños(0)
    numExtranios = 0

    For h = 0 To 1
        Set ws = HojaConfiguracion(nombresHojas(h))
        If ws Is Nothing Then GoTo SiguienteHoja

        colIdx = BuscarIndiceConfiguracion(ws, clienteNombre)
        If colIdx = 0 Then GoTo SiguienteHoja

        ' Fila de inicio de datos: 4 para "columnas", 3 para "filas"
        filaInicio = IIf(UCase(Trim(nombresHojas(h))) = UCase(Trim(HOJA_CONFIG_COL)), 4, 3)
        ultimaFila = UltimaFilaReal(ws, colIdx)

        For i = filaInicio To ultimaFila
            ' TRIM + UCASE aplicado al valor de celda
            valor = UCase(Trim(ws.Cells(i, colIdx).Value))

            ' Valores permitidos: vacío, MANTENER, QUITAR, BOB, CELERGO
            Select Case valor
                Case "", CONFIG_MANTENER, CONFIG_QUITAR, _
                     UCase(CLIENTE_BOB), UCase(CLIENTE_CELERGO)
                    ' OK — no hacer nada
                Case Else
                    ReDim Preserve extraños(numExtranios)
                    extraños(numExtranios) = "Hoja '" & ws.Name & "' fila " & i & _
                                             ": '" & ws.Cells(i, colIdx).Value & "'"
                    numExtranios = numExtranios + 1
            End Select
        Next i

SiguienteHoja:
    Next h

    If numExtranios > 0 Then
        msg = "Se han encontrado " & numExtranios & " valor(es) no reconocido(s) " & _
              "en la configuración de '" & clienteNombre & "':" & vbCrLf & vbCrLf
        Dim k As Long
        For k = 0 To numExtranios - 1
            msg = msg & "  " & Chr(149) & " " & extraños(k) & vbCrLf
        Next k
        msg = msg & vbCrLf & _
              "AVISO: cualquier valor distinto de MANTENER o QUITAR " & _
              "será tratado como QUITAR (la columna/fila se eliminará)." & _
              vbCrLf & vbCrLf & "¿Desea continuar de todas formas?"

        If MsgBox(msg, vbExclamation + vbYesNo, "Literales no reconocidos") = vbNo Then
            Err.Raise vbObjectError + 1, "ValidarLiteralesConfiguracion", "CANCELADO_POR_USUARIO"
        End If
    End If
End Sub

' ======================================================================================
' CALCULAR VERSIÓN SUGERIDA PARA EL NOMBRE DE FICHERO
' Devuelve "V01" si no existe ninguna versión previa, o "Vnn+1" si ya existe.
' No pregunta aquí — la confirmación la hace el propio diálogo de Guardar Como
' (el usuario puede cambiar el nombre libremente).
' ======================================================================================

Private Function CalcularVersionSugerida(ByVal rutaCarpeta As String, _
                                          ByVal cliente As String, _
                                          ByVal nomBase As String) As String
    Dim patronBase As String
    Dim versionNum As Integer
    Dim maxVersion As Integer

    patronBase = cliente & "_" & nomBase & "_"
    maxVersion = 0
    versionNum = 1

    Do While versionNum <= 99
        Dim rutaTest As String
        rutaTest = rutaCarpeta & patronBase & "V" & Format(versionNum, "00") & ".xlsx"
        If Dir(rutaTest) <> "" Then
            maxVersion = versionNum
            versionNum = versionNum + 1
        Else
            Exit Do
        End If
    Loop

    CalcularVersionSugerida = "V" & Format(maxVersion + 1, "00")
End Function

' ======================================================================================
' DIÁLOGO GUARDAR COMO NATIVO DE EXCEL
' - Abre en rutaInicial con nombreSugerido pre-informado en el campo Nombre
' - El usuario puede cambiar nombre, versión y carpeta libremente
' - Devuelve la ruta completa elegida, o "" si cancela
' ======================================================================================

Private Function MostrarDialogoGuardarComo(ByVal rutaInicial As String, _
                                            ByVal nombreSugerido As String) As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    With fd
        .Title           = "Guardar archivo de cliente"
        .InitialFileName = rutaInicial & nombreSugerido & ".xlsx"
        ' Filtro xlsx solamente
        On Error Resume Next
        .FilterIndex = 1   ' El índice exacto depende de la versión de Excel; no es crítico
        On Error GoTo 0

        If .Show = True Then
            Dim ruta As String
            ruta = .SelectedItems(1)
            ' Asegurar extensión .xlsx (el usuario podría haberla quitado)
            If UCase(Right(ruta, 5)) <> ".XLSX" Then ruta = ruta & ".xlsx"
            MostrarDialogoGuardarComo = ruta
        Else
            MostrarDialogoGuardarComo = ""
        End If
    End With

    Set fd = Nothing
End Function

' ======================================================================================
' EXTRAER CARPETA DE UNA RUTA COMPLETA DE FICHERO
' ======================================================================================

Private Function ExtraerCarpeta(ByVal rutaFichero As String) As String
    Dim pos As Long
    pos = 0
    Dim i As Long
    For i = Len(rutaFichero) To 1 Step -1
        If Mid(rutaFichero, i, 1) = "\" Then
            pos = i
            Exit For
        End If
    Next i
    If pos > 0 Then
        ExtraerCarpeta = Left(rutaFichero, pos)
    Else
        ExtraerCarpeta = ""
    End If
End Function

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
' VERIFICAR UNIDAD DE RED
' ======================================================================================

Private Function UnidadRedConectada(ByVal unidad As String) As Boolean
    Dim fso   As Object
    Dim drive As Object

    UnidadRedConectada = False

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function

    If fso.DriveExists(Left(Trim(unidad), 1)) Then
        Set drive = fso.GetDrive(Left(Trim(unidad), 1))
        If Not drive Is Nothing Then
            UnidadRedConectada = drive.IsReady
        End If
    End If

    Set drive = Nothing
    Set fso   = Nothing
    On Error GoTo 0
End Function

' ======================================================================================
' CREAR RUTA SI NO EXISTE (COMPATIBLE CON RED)
' ======================================================================================

Private Function AsegurarRuta(ByVal ruta As String) As Boolean
    Dim fso      As Object
    Dim carpetas() As String
    Dim rutaAcum As String
    Dim i        As Integer

    On Error Resume Next
    Set fso   = CreateObject("Scripting.FileSystemObject")
    carpetas  = Split(ruta, "\")
    rutaAcum  = carpetas(0)

    For i = 1 To UBound(carpetas)
        If Trim(carpetas(i)) <> "" Then
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

Private Sub DesprotegerHoja(ByVal ws As Worksheet, ByVal pwd As String)
    If Not ws.ProtectContents Then Exit Sub
    Debug.Print "  [PASS] Desprotegiendo: " & ws.Name
    On Error Resume Next
    ws.Unprotect
    If ws.ProtectContents Then ws.Unprotect pwd
    On Error GoTo 0
    If ws.ProtectContents Then
        Debug.Print "    [AVISO] No se pudo desproteger: " & ws.Name
    Else
        Debug.Print "    [OK] Desprotegida: " & ws.Name
    End If
End Sub

Private Sub DesprotegerTodasHojas(ByVal wb As Workbook, ByVal pwd As String)
    Dim ws As Worksheet
    Debug.Print "[PASS] Desprotegiendo todas las hojas..."
    For Each ws In wb.Worksheets
        DesprotegerHoja ws, pwd
    Next ws
End Sub

Private Sub AsegurarHojasDesprotegidas(ByVal wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.ProtectContents Then
            Debug.Print "  [AVISO] Hoja aún protegida: " & ws.Name
            DesprotegerHoja ws, PASSWORD_HOJAS
        End If
    Next ws
End Sub

' ======================================================================================
' PROCESO PRINCIPAL DE GENERACIÓN
' - Temporal en %TEMP%, borrado previo sin error
' - nombreFinal es la ruta completa elegida en el diálogo (ya incluye versión)
' - SaveAs con xlOpenXMLWorkbook + DisplayAlerts=False → sin diálogo de macros
' ======================================================================================

Private Sub EjecutarGeneracion(ByVal idConfig As String, ByVal indiceCol As Long, _
                                ByVal nombreFinal As String)
    Dim wbCopia   As Workbook
    Dim fTemporal As String
    Dim seguridad As Long
    Dim rutaTemp  As String

    rutaTemp  = Environ("TEMP")
    If Right(rutaTemp, 1) <> "\" Then rutaTemp = rutaTemp & "\"
    fTemporal = rutaTemp & "~tmp_" & idConfig & "_" & LimpiarExtension(ThisWorkbook.Name) & ".xlsm"

    Debug.Print "[1] Destino final:    " & nombreFinal
    Debug.Print "[2] Temporal en TEMP: " & fTemporal

    ' Borrar temporal previo sin error aunque no exista
    On Error Resume Next
    Kill fTemporal
    On Error GoTo 0
    Debug.Print "  [OK] Temporal previo purgado"

    ' Crear copia temporal
    Debug.Print "[3] Creando copia temporal..."
    ThisWorkbook.SaveCopyAs fTemporal
    Debug.Print "  [OK] Temporal creado"

    ' Abrir temporal sin disparar macros automáticas
    seguridad = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityLow
    Set wbCopia = Workbooks.Open(fTemporal, UpdateLinks:=0)
    Application.AutomationSecurity = seguridad
    Debug.Print "[4] Temporal abierto"

    ' Desproteger todas las hojas
    Debug.Print "[5] Desprotegiendo hojas..."
    DesprotegerTodasHojas wbCopia, PASSWORD_HOJAS

    ' Procesar columnas y filas
    Debug.Print "[6] Procesando columnas..."
    ProcesarColumnas wbCopia, idConfig, indiceCol

    Debug.Print "[7] Procesando filas..."
    ProcesarFilas wbCopia, idConfig

    ' Eliminar hojas de configuración
    ' DisplayAlerts ya es False gracias a GestionarEntorno True
    Debug.Print "[8] Eliminando hojas de configuración..."
    On Error Resume Next
    wbCopia.Worksheets(HOJA_CONFIG_COL).Delete
    wbCopia.Worksheets(HOJA_CONFIG_FIL).Delete
    On Error GoTo 0

    ' Verificación final de protección
    Debug.Print "[9] Verificación final de protección..."
    AsegurarHojasDesprotegidas wbCopia

    ' Asegurar que la carpeta destino existe
    Dim carpetaFinal As String
    carpetaFinal = ExtraerCarpeta(nombreFinal)
    If Not AsegurarRuta(carpetaFinal) Then
        MsgBox "Error: no se puede acceder a la ruta " & carpetaFinal, vbCritical
        wbCopia.Close SaveChanges:=False
        On Error Resume Next
        Kill fTemporal
        On Error GoTo 0
        Exit Sub
    End If

    ' Guardar como .xlsx — sin diálogo de macros (DisplayAlerts=False activo)
    Debug.Print "[10] Guardando como .xlsx..."
    wbCopia.SaveAs Filename:=nombreFinal, FileFormat:=xlOpenXMLWorkbook
    wbCopia.Close SaveChanges:=False
    Debug.Print "  [OK] Guardado: " & nombreFinal

    ' Eliminar temporal
    On Error Resume Next
    Kill fTemporal
    On Error GoTo 0
    Debug.Print "[11] Temporal eliminado"
    Debug.Print "[COMPLETADO] " & idConfig
End Sub

' ======================================================================================
' PROCESADO DE COLUMNAS
' TRIM+UCASE en todos los valores leídos de config.
' Literales no reconocidos → QUITAR (usuario ya fue avisado).
' ======================================================================================

Private Sub ProcesarColumnas(ByRef wb As Workbook, ByVal configNombre As String, ByVal colIndex As Long)
    Dim wsConfig    As Worksheet
    Dim wsDest      As Worksheet
    Dim ultimaFila  As Long
    Dim i           As Long
    Dim nombreCol   As String
    Dim valorConf   As String
    Dim colDest     As Long
    Dim listaBorrar As New Collection

    Set wsConfig = HojaConfiguracion(HOJA_CONFIG_COL)
    If wsConfig Is Nothing Then
        Debug.Print "  [AVISO] Hoja '" & HOJA_CONFIG_COL & "' no encontrada. Saltando columnas."
        Exit Sub
    End If

    On Error Resume Next
    Set wsDest = wb.Worksheets(HOJA_DATOS)
    On Error GoTo 0

    If wsDest Is Nothing Then
        Debug.Print "  [AVISO] Hoja '" & HOJA_DATOS & "' no encontrada en copia. Saltando columnas."
        Exit Sub
    End If

    DesprotegerHoja wsDest, PASSWORD_HOJAS

    ultimaFila = UltimaFilaReal(wsConfig, 2)
    Debug.Print "  [COLUMNAS] Última fila config: " & ultimaFila & " | Col config: " & colIndex

    For i = 4 To ultimaFila
        nombreCol = Trim(wsConfig.Cells(i, 2).Value)           ' TRIM en nombre de columna
        valorConf = UCase(Trim(wsConfig.Cells(i, colIndex).Value))  ' TRIM+UCASE en valor

        If nombreCol <> "" Then
            Select Case valorConf
                Case CONFIG_QUITAR
                    colDest = EncontrarColumnaPorTexto(wsDest, nombreCol)
                    If colDest > 0 Then
                        listaBorrar.Add colDest
                        Debug.Print "    [-] QUITAR col '" & nombreCol & "' (pos " & colDest & ")"
                    Else
                        Debug.Print "    [!] No encontrada: '" & nombreCol & "'"
                    End If

                Case CONFIG_MANTENER
                    ' Conservar tal cual

                Case "", UCase(CLIENTE_BOB), UCase(CLIENTE_CELERGO)
                    ' Vacío o nombre de cliente cruzado → ignorar

                Case Else
                    ' Literal desconocido → QUITAR (ya avisado)
                    colDest = EncontrarColumnaPorTexto(wsDest, nombreCol)
                    If colDest > 0 Then
                        listaBorrar.Add colDest
                        Debug.Print "    [-] QUITAR (literal '" & valorConf & "') col '" & nombreCol & "'"
                    Else
                        Debug.Print "    [!] No encontrada (literal '" & valorConf & "'): '" & nombreCol & "'"
                    End If
            End Select
        End If
    Next i

    Debug.Print "    Total columnas a quitar: " & listaBorrar.Count
    BorrarElementos wsDest, listaBorrar, "COLUMNA"
End Sub

' ======================================================================================
' PROCESADO DE FILAS
' TRIM+UCASE en todos los valores leídos de config.
' Literales no reconocidos → QUITAR (usuario ya fue avisado).
' ======================================================================================

Private Sub ProcesarFilas(ByRef wb As Workbook, ByVal configNombre As String)
    Dim wsConfig        As Worksheet
    Dim wsDest          As Worksheet
    Dim ultimaFila      As Long
    Dim i               As Long
    Dim textoFila       As String
    Dim valorConf       As String
    Dim filaDestino     As Long
    Dim listaBorrar     As New Collection
    Dim textoExtra      As String
    Dim colTexto        As Long
    Dim colExtra        As Long
    Dim colIndexFilas   As Long
    Dim colTextoDestino As Long

    Set wsConfig = HojaConfiguracion(HOJA_CONFIG_FIL)
    If wsConfig Is Nothing Then
        Debug.Print "  [AVISO] Hoja '" & HOJA_CONFIG_FIL & "' no encontrada. Saltando filas."
        Exit Sub
    End If

    On Error Resume Next
    Set wsDest = wb.Worksheets(HOJA_PREGUNTAS)
    On Error GoTo 0

    If wsDest Is Nothing Then
        Debug.Print "  [AVISO] Hoja '" & HOJA_PREGUNTAS & "' no encontrada en copia. Saltando filas."
        Exit Sub
    End If

    DesprotegerHoja wsDest, PASSWORD_HOJAS

    colIndexFilas = BuscarIndiceConfiguracion(wsConfig, configNombre)
    If colIndexFilas = 0 Then
        Debug.Print "  [ERROR] '" & configNombre & "' no encontrado en '" & HOJA_CONFIG_FIL & "'. Saltando."
        Exit Sub
    End If

    colTexto   = DetectarColumnaTextos(wsConfig, 3)
    colExtra   = colIndexFilas + 5
    ultimaFila = UltimaFilaReal(wsConfig, colTexto)

    Debug.Print "  [FILAS] Última fila: " & ultimaFila & " | ColTexto: " & colTexto & _
                " | ColConf: " & colIndexFilas & " | ColExtra: " & colExtra

    For i = 3 To ultimaFila
        textoFila  = Trim(wsConfig.Cells(i, colTexto).Value)
        valorConf  = UCase(Trim(wsConfig.Cells(i, colIndexFilas).Value))  ' TRIM+UCASE
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
                                Debug.Print "    [+] Texto extra col " & (colTextoDestino + 1) & " fila " & filaDestino
                            Else
                                Debug.Print "    [!] No se encontró col texto largo en fila " & filaDestino
                            End If
                        Else
                            Debug.Print "    [=] MANTENER sin cambios, fila " & filaDestino
                        End If

                    Case "", UCase(CLIENTE_BOB), UCase(CLIENTE_CELERGO)
                        ' Vacío o nombre de cliente cruzado → ignorar

                    Case Else
                        ' Literal desconocido → QUITAR (ya avisado)
                        listaBorrar.Add filaDestino
                        Debug.Print "    [-] QUITAR (literal '" & valorConf & "') fila " & filaDestino
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

' Busca la columna del cliente en la hoja de config (fila 2 o 3).
' TRIM+UCASE aplicado tanto al valor buscado como al leído.
Private Function BuscarIndiceConfiguracion(ByVal ws As Worksheet, ByVal nombre As String) As Long
    Dim filaEnc   As Long
    Dim c         As Long
    Dim ultimaCol As Long
    Dim hayDatos  As Boolean
    Dim buscado   As String

    buscado  = UCase(Trim(nombre))
    hayDatos = False

    For c = 1 To 15
        If UCase(Trim(ws.Cells(2, c).Value)) = buscado Then
            hayDatos = True
            Exit For
        End If
    Next c
    filaEnc = IIf(hayDatos, 2, 3)

    ultimaCol = ws.Cells(filaEnc, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ultimaCol
        If UCase(Trim(ws.Cells(filaEnc, c).Value)) = buscado Then
            BuscarIndiceConfiguracion = c
            Exit Function
        End If
    Next c

    BuscarIndiceConfiguracion = 0
End Function

' Busca en las primeras 5 filas de la hoja en qué columna está un encabezado.
' TRIM+UCASE en comparación.
Private Function EncontrarColumnaPorTexto(ByVal ws As Worksheet, ByVal txt As String) As Long
    Dim r      As Long
    Dim c      As Long
    Dim buscado As String
    buscado = UCase(Trim(txt))

    For r = 1 To 5
        For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            If UCase(Trim(ws.Cells(r, c).Value)) = buscado Then
                EncontrarColumnaPorTexto = c
                Exit Function
            End If
        Next c
    Next r
    EncontrarColumnaPorTexto = 0
End Function

' Detecta la columna con el texto más largo en una fila (referencia para filas de config).
Private Function DetectarColumnaTextos(ByVal ws As Worksheet, ByVal fila As Long) As Long
    Dim col       As Long
    Dim maxLen    As Long
    Dim colMax    As Long
    Dim lenActual As Long
    Dim ultimaCol As Long

    maxLen    = 0
    colMax    = 6
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

Private Function UltimaFilaReal(ByVal ws As Worksheet, ByVal col As Long) As Long
    UltimaFilaReal = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' Busca en qué fila de la hoja destino aparece un fragmento de texto.
Private Function BuscarFilaPorTexto(ByVal ws As Worksheet, ByVal txt As String) As Long
    Dim r          As Long
    Dim c          As Long
    Dim ultimaFila As Long
    Dim ultimaCol  As Long
    Dim fragmento  As String
    Dim textoCelda As String

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

' Devuelve la columna con el texto más largo en una fila (para saber dónde añadir texto extra).
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

' Borra filas o columnas de mayor a menor índice para no desplazar posiciones.
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

    ' Ordenar descendente (burbuja — válido para colecciones pequeñas)
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
