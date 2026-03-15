Attribute VB_Name = "Módulo2"

Option Explicit

' ======================================================================================
' CONSTANTES GLOBALES DE MÓDULO
' ======================================================================================

Private Const PASSWORD_HOJAS   As String = "ADP"
Private Const CONFIG_MANTENER  As String = "MANTENER"
Private Const CONFIG_QUITAR    As String = "QUITAR"
Private Const UNIDAD_RED       As String = "O:\"
' *** Rutas de destino por cliente — ajustar de forma independiente si cambian ***
Private Const RUTA_BOB     As String = "O:\ADP_SP\Clientes_Bob_LOPD\J18_ANCERT\J18_ANCERT\PRUEBAS\"
Private Const RUTA_CELERGO As String = "O:\ADP_SP\Clientes_Bob_LOPD\J18_ANCERT\J18_ANCERT\PRUEBAS\"
Private Const HOJA_DATOS       As String = "Analisis conceptos BOB"
Private Const HOJA_PREGUNTAS   As String = "Preguntas generales"

' Nombres de hojas de configuración — siempre comparar con UCASE+TRIM
Private Const HOJA_CONFIG_COL  As String = "columnas"
Private Const HOJA_CONFIG_FIL  As String = "filas"

' Nombres de clientes válidos — siempre comparar con UCASE+TRIM
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


' ======================================================================================
' FUNCIÓN CENTRALIZADA
' Flujo:
'   1. Validaciones previas (hojas config, literales, cliente)
'   2. Determinar carpeta destino:
'        - Si O:\ accesible → usar RUTA_POR_DEFECTO
'        - Si no           → MsgBox aviso + BrowseForFolder local
'   3. Con la carpeta ya conocida → calcular versión mirando los ficheros existentes
'   4. InputBox con nombre sugerido (CLIENTE_BASE_Vnn) para confirmar o editar
'   5. Guardar
' ======================================================================================

Private Sub GenerarArchivoIndividual(ByVal clienteNombre As String)
    Dim wsConfigCol  As Worksheet
    Dim nombreBase   As String
    Dim rutaCarpeta  As String
    Dim nombreFinal  As String
    Dim indiceCol    As Long

    Debug.Print String(80, "=")
    Debug.Print "INICIO PROCESO PARA: " & clienteNombre
    Debug.Print String(80, "=")

    ' ── 1a. Hoja de configuración de columnas ────────────────────────────────────────
    Set wsConfigCol = HojaConfiguracion(HOJA_CONFIG_COL)
    If wsConfigCol Is Nothing Then
        MsgBox "Error crítico: no se encontró la hoja '" & HOJA_CONFIG_COL & "'.", vbCritical
        Exit Sub
    End If

    ' ── 1b. Validar literales no reconocidos ─────────────────────────────────────────
    '   BOB, CELERGO, MANTENER, QUITAR y vacío son los únicos valores permitidos.
    '   Cualquier otro se tratará como QUITAR tras avisar al usuario.
    On Error GoTo ErrorValidacion
    ValidarLiteralesConfiguracion clienteNombre
    On Error GoTo 0

    ' ── 1c. Confirmar que el cliente existe en config ─────────────────────────────────
    indiceCol = BuscarIndiceConfiguracion(wsConfigCol, clienteNombre)
    If indiceCol = 0 Then
        MsgBox "Error: no se encontró '" & clienteNombre & "' en la hoja '" & _
               HOJA_CONFIG_COL & "'.", vbExclamation
        Exit Sub
    End If
    Debug.Print "[OK] '" & clienteNombre & "' en columna: " & indiceCol

    ' ── 2. Nombre base del fichero maestro ───────────────────────────────────────────
    nombreBase = LimpiarExtension(ThisWorkbook.Name)
    Debug.Print "[OK] Nombre base: " & nombreBase

    ' ── 3. Determinar carpeta destino ────────────────────────────────────────────────
    rutaCarpeta = ResolverCarpetaDestino(clienteNombre)
    If rutaCarpeta = "" Then
        Debug.Print "[CANCELADO] Sin carpeta destino."
        Exit Sub
    End If
    Debug.Print "[OK] Carpeta destino: " & rutaCarpeta

    ' ── 4. Calcular versión mirando los ficheros existentes en ESA carpeta ────────────
    Dim versionTag  As String
    Dim patronBase  As String
    patronBase = UCase(Trim(clienteNombre)) & "_" & nombreBase & "_"
    versionTag = CalcularVersionSugerida(rutaCarpeta, patronBase)
    Debug.Print "[OK] Versión calculada: " & versionTag

    ' ── 5. InputBox: confirmar o editar el nombre (sin extensión) ────────────────────
    Dim nombreSugerido As String
    nombreSugerido = patronBase & versionTag

    Dim nombreEditado As String
    nombreEditado = InputBox( _
        "Carpeta destino:" & vbCrLf & "  " & rutaCarpeta & vbCrLf & vbCrLf & _
        "Versión detectada: " & versionTag & vbCrLf & vbCrLf & _
        "Confirme o edite el nombre del fichero (sin extensión):", _
        "Nombre del archivo de salida", _
        nombreSugerido)

    ' Cancelar si InputBox devuelve vacío
    If Trim(nombreEditado) = "" Then
        Debug.Print "[CANCELADO] El usuario canceló el InputBox."
        Exit Sub
    End If

    ' Limpiar extensión por si el usuario la escribió
    nombreEditado = LimpiarExtension(Trim(nombreEditado))
    nombreFinal   = rutaCarpeta & nombreEditado & ".xlsx"
    Debug.Print "[OK] Fichero final: " & nombreFinal

    ' ── 6. Procesar y guardar ────────────────────────────────────────────────────────
    GestionarEntorno True
    EjecutarGeneracion clienteNombre, indiceCol, nombreFinal
    GestionarEntorno False

    Debug.Print String(80, "=")
    Debug.Print "COMPLETADO: " & clienteNombre
    Debug.Print String(80, "=")

    MsgBox "Proceso finalizado correctamente para " & clienteNombre & "." & vbCrLf & _
           "Archivo generado en:" & vbCrLf & nombreFinal, vbInformation
    Exit Sub

ErrorValidacion:
    Debug.Print "[CANCELADO] Usuario canceló tras aviso de literales."
    On Error GoTo 0
    GestionarEntorno False   ' por si se activó antes del error
End Sub

' ======================================================================================
' RESOLUCIÓN DE CARPETA DESTINO
' ─ Si O:\ accesible    → devuelve RUTA_POR_DEFECTO (sin diálogo)
' ─ Si O:\ NO accesible → MsgBox informando + BrowseForFolder para elegir ruta local
'                          Si el usuario cancela el BrowseForFolder → devuelve ""
' ======================================================================================

Private Function ResolverCarpetaDestino(ByVal clienteNombre As String) As String
    Dim ruta        As String
    Dim rutaCliente As String

    ' Seleccionar ruta según cliente — cada uno tiene su constante independiente
    Select Case UCase(Trim(clienteNombre))
        Case UCase(Trim(CLIENTE_BOB))
            rutaCliente = RUTA_BOB
        Case UCase(Trim(CLIENTE_CELERGO))
            rutaCliente = RUTA_CELERGO
        Case Else
            rutaCliente = RUTA_BOB   ' Fallback seguro
    End Select

    If UnidadRedConectada(UNIDAD_RED) Then
        ' Red disponible: usar ruta del cliente, sin preguntar
        ruta = rutaCliente
        Debug.Print "  [RED] Usando ruta: " & ruta
    Else
        ' Red no disponible: avisar y pedir carpeta local
        MsgBox "La unidad de red " & UNIDAD_RED & " no está accesible." & vbCrLf & vbCrLf & _
               "A continuación seleccione una carpeta local donde guardar el archivo.", _
               vbExclamation, "Unidad de red no disponible"

        ruta = SeleccionarCarpetaLocal("C:\")
        If ruta = "" Then
            ResolverCarpetaDestino = ""
            Exit Function
        End If
        Debug.Print "  [LOCAL] Carpeta seleccionada: " & ruta
    End If

    If Right(ruta, 1) <> "\" Then ruta = ruta & "\"
    ResolverCarpetaDestino = ruta
End Function

' ======================================================================================
' SELECCIÓN DE CARPETA LOCAL (BrowseForFolder)
' ======================================================================================

Private Function SeleccionarCarpetaLocal(Optional ByVal rutaInicial As String = "C:\") As String
    Dim shellApp    As Object
    Dim shellFolder As Object

    On Error Resume Next
    Set shellApp    = CreateObject("Shell.Application")
    Set shellFolder = shellApp.BrowseForFolder( _
                          0, "Seleccione la carpeta donde guardar el archivo:", 0, rutaInicial)
    On Error GoTo 0

    If shellFolder Is Nothing Then
        SeleccionarCarpetaLocal = ""
    Else
        SeleccionarCarpetaLocal = shellFolder.Self.Path
    End If

    Set shellFolder = Nothing
    Set shellApp    = Nothing
End Function

' ======================================================================================
' CALCULAR VERSIÓN SUGERIDA
' Recibe la carpeta ya resuelta y el patrón base (p.ej. "BOB_Fichero_")
' Busca BOB_Fichero_V01.xlsx, V02... y devuelve la siguiente disponible.
' ======================================================================================

Private Function CalcularVersionSugerida(ByVal rutaCarpeta As String, _
                                          ByVal patronBase  As String) As String
    Dim versionNum As Integer
    Dim maxVersion As Integer

    maxVersion = 0
    versionNum = 1

    Do While versionNum <= 99
        Dim candidato As String
        candidato = rutaCarpeta & patronBase & "V" & Format(versionNum, "00") & ".xlsx"
        If Dir(candidato) <> "" Then
            maxVersion = versionNum
            versionNum = versionNum + 1
        Else
            Exit Do
        End If
    Loop

    CalcularVersionSugerida = "V" & Format(maxVersion + 1, "00")
End Function

' ======================================================================================
' OBTENER HOJA DE CONFIGURACIÓN — búsqueda con UCASE+TRIM en el nombre
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
' Valores permitidos (todos comparados con UCASE+TRIM):
'   MANTENER, QUITAR, BOB, CELERGO, "" (vacío)
' Cualquier otro se avisa y, si el usuario acepta, se trata como QUITAR.
' Si el usuario cancela se eleva error para abortar el proceso.
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
        Set ws = HojaConfiguracion(nombresHojas(h))   ' UCASE+TRIM en nombre de hoja
        If ws Is Nothing Then GoTo SiguienteHoja

        colIdx = BuscarIndiceConfiguracion(ws, clienteNombre)   ' UCASE+TRIM en cliente
        If colIdx = 0 Then GoTo SiguienteHoja

        ' Fila de inicio: 4 para "columnas", 3 para "filas"
        filaInicio = IIf(UCase(Trim(nombresHojas(h))) = UCase(Trim(HOJA_CONFIG_COL)), 4, 3)
        ultimaFila = UltimaFilaReal(ws, colIdx)

        For i = filaInicio To ultimaFila
            valor = UCase(Trim(ws.Cells(i, colIdx).Value))   ' UCASE+TRIM en valor

            Select Case valor
                Case "", _
                     UCase(Trim(CONFIG_MANTENER)), _
                     UCase(Trim(CONFIG_QUITAR)), _
                     UCase(Trim(CLIENTE_BOB)), _
                     UCase(Trim(CLIENTE_CELERGO))
                    ' Valor permitido — no hacer nada
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
' VERIFICAR UNIDAD DE RED — UCASE+TRIM en el nombre de unidad
' ======================================================================================

Private Function UnidadRedConectada(ByVal unidad As String) As Boolean
    Dim fso   As Object
    Dim drive As Object

    UnidadRedConectada = False

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function

    Dim letra As String
    letra = UCase(Trim(Left(Trim(unidad), 1)))

    If fso.DriveExists(letra) Then
        Set drive = fso.GetDrive(letra)
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
' EXTRAER CARPETA DE UNA RUTA COMPLETA
' ======================================================================================

Private Function ExtraerCarpeta(ByVal rutaFichero As String) As String
    Dim i As Long
    For i = Len(rutaFichero) To 1 Step -1
        If Mid(rutaFichero, i, 1) = "\" Then
            ExtraerCarpeta = Left(rutaFichero, i)
            Exit Function
        End If
    Next i
    ExtraerCarpeta = ""
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
' ─ Temporal en %TEMP%, borrado previo sin error
' ─ nombreFinal es la ruta completa ya resuelta (carpeta + nombre + .xlsx)
' ─ SaveAs con xlOpenXMLWorkbook + DisplayAlerts=False → sin diálogo de macros
' ======================================================================================

Private Sub EjecutarGeneracion(ByVal idConfig As String, ByVal indiceCol As Long, _
                                ByVal nombreFinal As String)
    Dim wbCopia   As Workbook
    Dim fTemporal As String
    Dim seguridad As Long
    Dim rutaTemp  As String

    rutaTemp  = Environ("TEMP")
    If Right(rutaTemp, 1) <> "\" Then rutaTemp = rutaTemp & "\"
    fTemporal = rutaTemp & "~tmp_" & UCase(Trim(idConfig)) & "_" & _
                LimpiarExtension(ThisWorkbook.Name) & ".xlsm"

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

    ' Abrir sin disparar macros automáticas
    seguridad = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityLow
    Set wbCopia = Workbooks.Open(fTemporal, UpdateLinks:=0)
    Application.AutomationSecurity = seguridad
    Debug.Print "[4] Temporal abierto"

    Debug.Print "[5] Desprotegiendo hojas..."
    DesprotegerTodasHojas wbCopia, PASSWORD_HOJAS

    Debug.Print "[6] Procesando columnas..."
    ProcesarColumnas wbCopia, idConfig, indiceCol

    Debug.Print "[7] Procesando filas..."
    ProcesarFilas wbCopia, idConfig

    ' Eliminar hojas de config — DisplayAlerts ya False por GestionarEntorno
    Debug.Print "[8] Eliminando hojas de configuración..."
    On Error Resume Next
    wbCopia.Worksheets(HOJA_CONFIG_COL).Delete
    wbCopia.Worksheets(HOJA_CONFIG_FIL).Delete
    On Error GoTo 0

    Debug.Print "[9] Verificación final de protección..."
    AsegurarHojasDesprotegidas wbCopia

    ' Asegurar carpeta destino
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

    ' Guardar como .xlsx sin diálogo de macros
    Debug.Print "[10] Guardando como .xlsx..."
    wbCopia.SaveAs Filename:=nombreFinal, FileFormat:=xlOpenXMLWorkbook
    wbCopia.Close SaveChanges:=False
    Debug.Print "  [OK] Guardado: " & nombreFinal

    On Error Resume Next
    Kill fTemporal
    On Error GoTo 0
    Debug.Print "[11] Temporal eliminado"
    Debug.Print "[COMPLETADO] " & idConfig
End Sub

' ======================================================================================
' PROCESADO DE COLUMNAS — UCASE+TRIM en todos los valores leídos
' Literales desconocidos → QUITAR (usuario ya fue avisado)
' ======================================================================================

Private Sub ProcesarColumnas(ByRef wb As Workbook, ByVal configNombre As String, _
                              ByVal colIndex As Long)
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
        Debug.Print "  [AVISO] Hoja '" & HOJA_DATOS & "' no encontrada en copia. Saltando."
        Exit Sub
    End If

    DesprotegerHoja wsDest, PASSWORD_HOJAS

    ultimaFila = UltimaFilaReal(wsConfig, 2)
    Debug.Print "  [COLUMNAS] Última fila: " & ultimaFila & " | Col config: " & colIndex

    For i = 4 To ultimaFila
        nombreCol = Trim(wsConfig.Cells(i, 2).Value)
        valorConf = UCase(Trim(wsConfig.Cells(i, colIndex).Value))   ' UCASE+TRIM

        If nombreCol <> "" Then
            Select Case valorConf
                Case UCase(Trim(CONFIG_QUITAR))
                    colDest = EncontrarColumnaPorTexto(wsDest, nombreCol)
                    If colDest > 0 Then
                        listaBorrar.Add colDest
                        Debug.Print "    [-] QUITAR col '" & nombreCol & "' (pos " & colDest & ")"
                    Else
                        Debug.Print "    [!] No encontrada: '" & nombreCol & "'"
                    End If

                Case UCase(Trim(CONFIG_MANTENER))
                    ' Conservar tal cual

                Case "", UCase(Trim(CLIENTE_BOB)), UCase(Trim(CLIENTE_CELERGO))
                    ' Vacío o cliente cruzado → ignorar

                Case Else
                    ' Literal desconocido → QUITAR (ya avisado)
                    colDest = EncontrarColumnaPorTexto(wsDest, nombreCol)
                    If colDest > 0 Then
                        listaBorrar.Add colDest
                        Debug.Print "    [-] QUITAR ('" & valorConf & "') col '" & nombreCol & "'"
                    Else
                        Debug.Print "    [!] No encontrada ('" & valorConf & "'): '" & nombreCol & "'"
                    End If
            End Select
        End If
    Next i

    Debug.Print "    Total columnas a quitar: " & listaBorrar.Count
    BorrarElementos wsDest, listaBorrar, "COLUMNA"
End Sub

' ======================================================================================
' PROCESADO DE FILAS — UCASE+TRIM en todos los valores leídos
' Literales desconocidos → QUITAR (usuario ya fue avisado)
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
        Debug.Print "  [AVISO] Hoja '" & HOJA_PREGUNTAS & "' no encontrada en copia. Saltando."
        Exit Sub
    End If

    DesprotegerHoja wsDest, PASSWORD_HOJAS

    colIndexFilas = BuscarIndiceConfiguracion(wsConfig, configNombre)   ' UCASE+TRIM interno
    If colIndexFilas = 0 Then
        Debug.Print "  [ERROR] '" & configNombre & "' no encontrado en '" & HOJA_CONFIG_FIL & "'."
        Exit Sub
    End If

    colTexto   = DetectarColumnaTextos(wsConfig, 3)
    colExtra   = colIndexFilas + 5
    ultimaFila = UltimaFilaReal(wsConfig, colTexto)

    Debug.Print "  [FILAS] Última fila: " & ultimaFila & " | ColTexto: " & colTexto & _
                " | ColConf: " & colIndexFilas & " | ColExtra: " & colExtra

    For i = 3 To ultimaFila
        textoFila  = Trim(wsConfig.Cells(i, colTexto).Value)
        valorConf  = UCase(Trim(wsConfig.Cells(i, colIndexFilas).Value))   ' UCASE+TRIM
        textoExtra = Trim(wsConfig.Cells(i, colExtra).Value)

        If Len(textoFila) > 5 Then
            filaDestino = BuscarFilaPorTexto(wsDest, textoFila)

            If filaDestino > 0 Then
                Select Case valorConf
                    Case UCase(Trim(CONFIG_QUITAR))
                        listaBorrar.Add filaDestino
                        Debug.Print "    [-] QUITAR fila " & filaDestino & _
                                    " ('" & Left(textoFila, 35) & "...')"

                    Case UCase(Trim(CONFIG_MANTENER))
                        If textoExtra <> "" Then
                            colTextoDestino = EncontrarColumnaConTextoLargo(wsDest, filaDestino)
                            If colTextoDestino > 0 Then
                                wsDest.Cells(filaDestino, colTextoDestino + 1).Value = textoExtra
                                Debug.Print "    [+] Texto extra col " & _
                                            (colTextoDestino + 1) & " fila " & filaDestino
                            Else
                                Debug.Print "    [!] No se encontró col texto largo, fila " & filaDestino
                            End If
                        Else
                            Debug.Print "    [=] MANTENER sin cambios, fila " & filaDestino
                        End If

                    Case "", UCase(Trim(CLIENTE_BOB)), UCase(Trim(CLIENTE_CELERGO))
                        ' Vacío o cliente cruzado → ignorar

                    Case Else
                        ' Literal desconocido → QUITAR (ya avisado)
                        listaBorrar.Add filaDestino
                        Debug.Print "    [-] QUITAR ('" & valorConf & "') fila " & filaDestino
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

' Busca la columna del cliente en la hoja de config (fila 2 o 3) — UCASE+TRIM siempre
Private Function BuscarIndiceConfiguracion(ByVal ws As Worksheet, _
                                            ByVal nombre As String) As Long
    Dim buscado   As String
    Dim filaEnc   As Long
    Dim c         As Long
    Dim ultimaCol As Long
    Dim hayDatos  As Boolean

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

' Busca encabezado en primeras 5 filas — UCASE+TRIM en comparación
Private Function EncontrarColumnaPorTexto(ByVal ws As Worksheet, _
                                           ByVal txt As String) As Long
    Dim buscado As String
    Dim r       As Long
    Dim c       As Long
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
            If Len(textoCelda) > 10 And _
               InStr(1, textoCelda, fragmento, vbTextCompare) > 0 Then
                BuscarFilaPorTexto = r
                Exit Function
            End If
        Next c
    Next r

    BuscarFilaPorTexto = 0
End Function

Private Function EncontrarColumnaConTextoLargo(ByVal ws As Worksheet, _
                                                ByVal fila As Long) As Long
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

' Borra filas/columnas de mayor a menor índice para no desplazar posiciones
Private Sub BorrarElementos(ByVal ws As Worksheet, ByVal indices As Collection, _
                             ByVal tipo As String)
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
