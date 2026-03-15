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

    ' 2. Validar literales no reconocidos en columnas y filas ANTES de continuar
    '    Todo literal distinto de MANTENER/QUITAR se trata como QUITAR tras avisar al usuario
    ValidarLiteralesConfiguracion clienteNombre

    ' 3. Validar que el cliente existe en la configuración
    indiceCol = BuscarIndiceConfiguracion(wsConfigCol, clienteNombre)
    If indiceCol = 0 Then
        MsgBox "Error: no se encontró la configuración para '" & clienteNombre & "' en la hoja 'columnas'.", vbExclamation
        Exit Sub
    End If
    Debug.Print "[OK] '" & clienteNombre & "' encontrado en columna: " & indiceCol

    ' 4. Obtener ruta de destino
    rutaDestino = ObtenerRutaDestino()
    If rutaDestino = "" Then
        Debug.Print "[CANCELADO] El usuario canceló la operación."
        Exit Sub
    End If
    Debug.Print "[OK] Ruta seleccionada: " & rutaDestino

    ' 5. Nombre base del archivo sin extensión
    nombreBase = LimpiarExtension(ThisWorkbook.Name)
    Debug.Print "[OK] Nombre base: " & nombreBase

    ' 6. Resolver versión del fichero de salida (con control de sobreescritura)
    Dim versionFinal As String
    versionFinal = ResolverVersionArchivo(rutaDestino, clienteNombre, nombreBase)
    If versionFinal = "" Then
        Debug.Print "[CANCELADO] El usuario canceló la selección de versión."
        Exit Sub
    End If
    Debug.Print "[OK] Versión resuelta: " & versionFinal

    ' 7. Optimizar rendimiento
    GestionarEntorno True

    ' 8. Generar archivo
    EjecutarGeneracion clienteNombre, indiceCol, rutaDestino, nombreBase, versionFinal

    ' 9. Restaurar entorno
    GestionarEntorno False

    Debug.Print String(80, "=")
    Debug.Print "PROCESO COMPLETADO PARA: " & clienteNombre
    Debug.Print String(80, "=")

    MsgBox "Proceso finalizado correctamente para " & clienteNombre & "." & vbCrLf & _
           "Archivo generado en:" & vbCrLf & _
           rutaDestino & clienteNombre & "_" & nombreBase & "_" & versionFinal & ".xlsx", _
           vbInformation
End Sub

' ======================================================================================
' VALIDACIÓN DE LITERALES EN HOJAS DE CONFIGURACIÓN
' Detecta valores distintos de MANTENER/QUITAR para el cliente dado y avisa al usuario.
' Esos valores se tratarán como QUITAR durante la ejecución.
' ======================================================================================

Private Sub ValidarLiteralesConfiguracion(ByVal clienteNombre As String)
    Dim hojas(1)    As String
    hojas(0) = "columnas"
    hojas(1) = "filas"

    Dim h           As Integer
    Dim ws          As Worksheet
    Dim colIdx      As Long
    Dim ultimaFila  As Long
    Dim i           As Long
    Dim valor       As String
    Dim extraños()  As String
    Dim numExtranios As Long
    Dim msg         As String

    ReDim extraños(0)
    numExtranios = 0

    For h = 0 To 1
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(hojas(h))
        On Error GoTo 0
        If ws Is Nothing Then GoTo SiguienteHoja

        colIdx = BuscarIndiceConfiguracion(ws, clienteNombre)
        If colIdx = 0 Then GoTo SiguienteHoja

        ' Determinar fila de inicio de datos (4 para columnas, 3 para filas)
        Dim filaInicio As Long
        filaInicio = IIf(hojas(h) = "columnas", 4, 3)

        ultimaFila = UltimaFilaReal(ws, colIdx)

        For i = filaInicio To ultimaFila
            valor = UCase(Trim(ws.Cells(i, colIdx).Value))
            If valor <> "" And valor <> CONFIG_MANTENER And valor <> CONFIG_QUITAR Then
                ReDim Preserve extraños(numExtranios)
                extraños(numExtranios) = "Hoja '" & hojas(h) & "' fila " & i & ": '" & _
                                          ws.Cells(i, colIdx).Value & "'"
                numExtranios = numExtranios + 1
            End If
        Next i

SiguienteHoja:
    Next h

    If numExtranios > 0 Then
        msg = "Se han encontrado " & numExtranios & " valor(es) no reconocido(s) en la configuración " & _
              "de '" & clienteNombre & "':" & vbCrLf & vbCrLf
        Dim k As Long
        For k = 0 To numExtranios - 1
            msg = msg & "  • " & extraños(k) & vbCrLf
        Next k
        msg = msg & vbCrLf & "AVISO: Cualquier valor distinto de MANTENER o QUITAR " & _
              "será tratado como QUITAR (la columna/fila se eliminará)." & vbCrLf & vbCrLf & _
              "¿Desea continuar de todas formas?"

        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox(msg, vbExclamation + vbYesNo, "Literales no reconocidos")
        If respuesta = vbNo Then
            ' Abortar: usamos Err.Raise para que el llamador termine limpiamente
            Err.Raise vbObjectError + 1, "ValidarLiteralesConfiguracion", "CANCELADO_POR_USUARIO"
        End If
    End If
End Sub

' ======================================================================================
' RESOLUCIÓN DE VERSIÓN DEL ARCHIVO DE SALIDA
' Devuelve "V01", "V02", etc. según los archivos existentes en carpeta destino.
' Si ya existe V01, pregunta si sobreescribir o crear V02 (y sucesivos).
' Devuelve "" si el usuario cancela.
' ======================================================================================

Private Function ResolverVersionArchivo(ByVal rutaCarpeta As String, _
                                         ByVal cliente As String, _
                                         ByVal nomBase As String) As String
    Dim patronBase As String
    Dim versionNum As Integer
    Dim candidato  As String
    Dim rutaTest   As String

    patronBase = cliente & "_" & nomBase & "_"

    ' Buscar la versión más alta existente
    Dim maxVersion As Integer
    maxVersion = 0
    versionNum = 1
    Do
        candidato = patronBase & "V" & Format(versionNum, "00") & ".xlsx"
        rutaTest  = rutaCarpeta & candidato
        If Dir(rutaTest) <> "" Then
            maxVersion = versionNum
            versionNum = versionNum + 1
        Else
            Exit Do
        End If
    Loop While versionNum <= 99

    If maxVersion = 0 Then
        ' No existe ninguna versión previa → usar V01 directamente
        ResolverVersionArchivo = "V01"
        Exit Function
    End If

    ' Existe al menos una versión: preguntar qué hacer
    Dim vExistente As String
    vExistente = "V" & Format(maxVersion, "00")
    Dim vNueva As String
    vNueva = "V" & Format(maxVersion + 1, "00")

    Dim msg As String
    msg = "Ya existe el archivo:" & vbCrLf & _
          "  " & patronBase & vExistente & ".xlsx" & vbCrLf & vbCrLf & _
          "¿Qué desea hacer?" & vbCrLf & vbCrLf & _
          "  [Sí]      Sobreescribir " & vExistente & vbCrLf & _
          "  [No]      Crear nueva versión " & vNueva & vbCrLf & _
          "  [Cancelar] Abortar el proceso"

    Dim resp As VbMsgBoxResult
    resp = MsgBox(msg, vbQuestion + vbYesNoCancel, "Versión de archivo existente")

    Select Case resp
        Case vbYes
            ResolverVersionArchivo = vExistente      ' Sobreescribir
        Case vbNo
            ResolverVersionArchivo = vNueva          ' Nueva versión
        Case vbCancel
            ResolverVersionArchivo = ""              ' Abortar
    End Select
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

    UnidadRedConectada = False

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function

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

Private Sub DesprotegerHoja(ByVal ws As Worksheet, ByVal pwd As String)
    If Not ws.ProtectContents Then Exit Sub

    Debug.Print "  [PASS] Desprotegiendo: " & ws.Name

    On Error Resume Next
    ws.Unprotect
    If ws.ProtectContents Then _
        ws.Unprotect pwd
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
            Debug.Print "  [AVISO] Hoja aún protegida tras proceso: " & ws.Name
            DesprotegerHoja ws, PASSWORD_HOJAS
        End If
    Next ws
End Sub

' ======================================================================================
' PROCESO PRINCIPAL DE GENERACIÓN
' CAMBIOS:
'   - Temporal en %TEMP% en lugar de ThisWorkbook.Path
'   - Temporal se borra ANTES de crearlo (evita error si quedó de ejecución anterior)
'   - Nombre de salida incluye versión (versionTag) recibida como parámetro
'   - SaveAs con FileFormat=51 (xlsx) suprime el diálogo de "perderá macros"
'     porque xlsx no tiene macros y DisplayAlerts=False ya está activo
' ======================================================================================

Private Sub EjecutarGeneracion(ByVal idConfig As String, ByVal indiceCol As Long, _
                                ByVal rutaCarpeta As String, ByVal nomBase As String, _
                                ByVal versionTag As String)
    Dim wbCopia   As Workbook
    Dim fFinal    As String
    Dim fTemporal As String
    Dim seguridad As Long
    Dim rutaTemp  As String

    ' Ruta %TEMP% del sistema operativo
    rutaTemp  = Environ("TEMP")
    If Right(rutaTemp, 1) <> "\" Then rutaTemp = rutaTemp & "\"

    fFinal    = rutaCarpeta & idConfig & "_" & nomBase & "_" & versionTag & ".xlsx"
    fTemporal = rutaTemp & "~tmp_" & idConfig & "_" & nomBase & ".xlsm"

    Debug.Print "[1] Destino final:    " & fFinal
    Debug.Print "[2] Temporal en TEMP: " & fTemporal

    ' ----- BORRAR TEMPORAL ANTES DE CREAR (aunque no exista, sin error) -----
    On Error Resume Next
    Kill fTemporal
    On Error GoTo 0
    Debug.Print "  [OK] Temporal previo purgado (si existía)"

    ' ----- Crear copia temporal -----
    Debug.Print "[3] Creando copia temporal..."
    ThisWorkbook.SaveCopyAs fTemporal
    Debug.Print "  [OK] Temporal creado"

    ' ----- Abrir temporal sin disparar macros automáticas -----
    seguridad = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityLow
    Set wbCopia = Workbooks.Open(fTemporal, UpdateLinks:=0)
    Application.AutomationSecurity = seguridad
    Debug.Print "[4] Temporal abierto"

    ' ----- Desproteger todas las hojas -----
    Debug.Print "[5] Desprotegiendo hojas..."
    DesprotegerTodasHojas wbCopia, PASSWORD_HOJAS

    ' ----- Procesar columnas y filas -----
    Debug.Print "[6] Procesando columnas..."
    ProcesarColumnas wbCopia, idConfig, indiceCol

    Debug.Print "[7] Procesando filas..."
    ProcesarFilas wbCopia, idConfig

    ' ----- Eliminar hojas de configuración -----
    Debug.Print "[8] Eliminando hojas de configuración..."
    ' DisplayAlerts ya está False gracias a GestionarEntorno True (llamado antes)
    On Error Resume Next
    wbCopia.Worksheets("columnas").Delete
    wbCopia.Worksheets("filas").Delete
    On Error GoTo 0

    ' ----- Verificación final de protección -----
    Debug.Print "[9] Verificación final de protección..."
    AsegurarHojasDesprotegidas wbCopia

    ' ----- Crear ruta si no existe -----
    If Not AsegurarRuta(rutaCarpeta) Then
        MsgBox "Error: no se puede acceder a la ruta " & rutaCarpeta, vbCritical
        wbCopia.Close SaveChanges:=False
        ' Limpiar temporal aunque haya fallo
        On Error Resume Next
        Kill fTemporal
        On Error GoTo 0
        Exit Sub
    End If

    ' ----- Guardar como .xlsx (sin diálogo de macros, DisplayAlerts=False activo) -----
    Debug.Print "[10] Guardando como .xlsx..."
    wbCopia.SaveAs Filename:=fFinal, FileFormat:=xlOpenXMLWorkbook
    wbCopia.Close SaveChanges:=False
    Debug.Print "  [OK] Guardado: " & fFinal

    ' ----- Borrar temporal -----
    On Error Resume Next
    Kill fTemporal
    On Error GoTo 0
    Debug.Print "[11] Temporal eliminado"
    Debug.Print "[COMPLETADO] " & idConfig
End Sub

' ======================================================================================
' PROCESADO DE COLUMNAS
' Literales no reconocidos se tratan como QUITAR (ya avisado en ValidarLiteralesConfiguracion)
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
                    ' Se conserva tal cual
                Case ""
                    ' Celda vacía, ignorar
                Case Else
                    ' Literal no reconocido → tratar como QUITAR (usuario ya fue avisado)
                    colDest = EncontrarColumnaPorTexto(wsDest, nombreCol)
                    If colDest > 0 Then
                        listaBorrar.Add colDest
                        Debug.Print "    [-] QUITAR (literal extranio '" & valorConf & "') columna '" & nombreCol & "' (pos " & colDest & ")"
                    Else
                        Debug.Print "    [!] No encontrada en destino (literal extranio): '" & nombreCol & "'"
                    End If
            End Select
        End If
    Next i

    Debug.Print "    Total columnas a quitar: " & listaBorrar.Count
    BorrarElementos wsDest, listaBorrar, "COLUMNA"
End Sub

' ======================================================================================
' PROCESADO DE FILAS
' Literales no reconocidos se tratan como QUITAR (ya avisado en ValidarLiteralesConfiguracion)
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

    colIndexFilas = BuscarIndiceConfiguracion(wsConfig, configNombre)
    If colIndexFilas = 0 Then
        Debug.Print "  [ERROR] '" & configNombre & "' no encontrado en hoja 'filas'. Saltando."
        Exit Sub
    End If

    colTexto   = DetectarColumnaTextos(wsConfig, 3)
    colExtra   = colIndexFilas + 5
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
                                Debug.Print "    [+] Texto extra en col " & (colTextoDestino + 1) & _
                                            " fila " & filaDestino
                            Else
                                Debug.Print "    [!] No se encontró col texto largo en fila " & filaDestino
                            End If
                        Else
                            Debug.Print "    [=] MANTENER sin cambios, fila " & filaDestino
                        End If

                    Case ""
                        ' Vacío, ignorar

                    Case Else
                        ' Literal no reconocido → tratar como QUITAR (ya avisado)
                        listaBorrar.Add filaDestino
                        Debug.Print "    [-] QUITAR (literal extranio '" & valorConf & "') fila " & filaDestino & _
                                    " ('" & Left(textoFila, 35) & "...')"
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

Private Function BuscarIndiceConfiguracion(ByVal ws As Worksheet, ByVal nombre As String) As Long
    Dim filaEnc   As Long
    Dim c         As Long
    Dim ultimaCol As Long
    Dim hayDatos  As Boolean

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
            If Len(textoCelda) > 10 And InStr(1, textoCelda, fragmento, vbTextCompare) > 0 Then
                BuscarFilaPorTexto = r
                Exit Function
            End If
        Next c
    Next r

    BuscarFilaPorTexto = 0
End Function

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
