Attribute VB_Name = "Modulo_Parche_V3_Final"

Option Explicit

' ----------------------------------------------------------------------------------------------------------------------
'  ESTA FUNCION ES PARA QUE EL EXCEL NO SE PONGA TONTO CON LOS PERMISOS
' ----------------------------------------------------------------------------------------------------------------------
Private Function F_VerificarRuta(ByVal p As String) As Boolean
    On Error Resume Next
    MkDir "C:\CLIENTES"
    MkDir "C:\CLIENTES\PRUEBAS"
    MkDir "C:\CLIENTES\PRUEBAS\BP"
    F_VerificarRuta = (Dir(p, vbDirectory) <> "")
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------------------------------------------------
'  BUSCA FILAS (CON EL ESPACIADO RARO QUE QUERIAS)
' ----------------------------------------------------------------------------------------------------------------------
Private Function F_DondeEstaFila(ByVal h As Worksheet, ByVal t As String) As Long
    Dim r As Long
    Dim u As Long
    Dim b As String
    u = h.Cells(h.Rows.Count, 1).End(xlUp).Row
    b = IIf(Len(t) > 15, Left(t, 15), t)
    For r = 1 To u
      If InStr(1, h.Cells(r, 1).Value, b, 1) > 0 Then
          F_DondeEstaFila = r
          Exit Function
      End If
    Next r
End Function

' ----------------------------------------------------------------------------------------------------------------------
'  BORRADOR DE BLOQUES (SOLO SI EL ID ESTA EN LA LISTA)
' ----------------------------------------------------------------------------------------------------------------------
Private Sub S_LimpiarTodo(ByVal w As Worksheet, ByVal c As Collection, ByVal m As Integer)
    Dim a() As String
    Dim i As Long
    Dim j As Long
    Dim tmp As String
    If c.Count = 0 Then
        Exit Sub
    End If
    ReDim a(1 To c.Count)
    For i = 1 To c.Count
        a(i) = CStr(c(i))
    Next i
    For i = 1 To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If Val(Split(a(i), "|")(0)) < Val(Split(a(j), "|")(0)) Then
                tmp = a(i)
                a(i) = a(j)
                a(j) = tmp
            End If
        Next j
    Next i
    For i = 1 To UBound(a)
        If m = 1 Then
            w.Columns(CLng(a(i))).Delete
        End If
        If m = 2 Then
            Dim p() As String
            p = Split(a(i), "|")
            If UBound(p) >= 1 Then
                If p(1) = "ADD" Then
                    w.Cells(CLng(p(0)), 3).Value = p(2)
                Else
                    w.Rows(CLng(p(0))).Delete
                End If
            Else
                w.Rows(CLng(a(i))).Delete
            End If
        End If
    Next i
End Sub

' **********************************************************************************************************************
'  AQUÍ ESTÁ EL TRUCO PARA EL ERROR DE SEGURIDAD Y LA CORRECCION DEL iD
' **********************************************************************************************************************
Private Sub S_Proceso_Interno(ByVal hC As Worksheet, ByVal id As String, ByVal rB As String, ByVal nO As String)
    Dim wb As Workbook
    Dim fFinal As String
    Dim fTmp As String
    Dim iH As Long
    Dim colC As Collection
    Dim colF As Collection
    Dim seguridad_previa As Long
    
    fFinal = rB & nO & "_" & id & ".xlsx"
    fTmp = ThisWorkbook.Path & "\~tmp" & id & ".xlsm" 
    
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs fTmp
    
    seguridad_previa = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityLow
    
    Set wb = Workbooks.Open(fTmp, UpdateLinks:=0)
    Application.AutomationSecurity = seguridad_previa
    
    ' --- PROCESO DE BORRADO DE COLUMNAS ---
    Dim h1 As Worksheet
    On Error Resume Next
    Set h1 = wb.Worksheets("columnas")
    On Error GoTo 0
    
    If Not h1 Is Nothing Then
        Dim ff As Long
        iH = 0
        For ff = 1 To 5
            If InStr(1, h1.Cells(ff, 2).Value, id, 1) > 0 Or InStr(1, h1.Cells(ff, 3).Value, id, 1) > 0 Then
                iH = ff
                Exit For
            End If
        Next ff
        If iH = 0 Then
            iH = 2
        End If
        
        Dim cc As Long
        Dim pIdx As Long
        For cc = 1 To h1.Cells(iH, h1.Columns.Count).End(xlToLeft).Column
            If UCase(Trim(h1.Cells(iH, cc).Value)) = UCase(id) Then
                pIdx = cc
                Exit For
            End If
        Next cc
        
        If pIdx > 0 Then
            Set colC = New Collection
            Dim rr As Long
            Dim rM As Long
            Dim nEt As String
            rM = h1.Cells(h1.Rows.Count, 2).End(xlUp).Row
            For rr = 4 To rM
                nEt = Trim(h1.Cells(rr, 2).Value)
                If nEt <> "" Then
                    If UCase(Trim(h1.Cells(rr, pIdx).Value)) = "NO" Then
                        Dim jI As Long
                        Dim fE As Long
                        Dim iD_Col As Long ' <--- CAMBIADO PARA QUE NO HAYA DUPLICADOS
                        iD_Col = 0
                        For fE = 1 To 10
                            If wb.Worksheets("FuncionFiltar").Cells(fE, 1).Value <> "" Then
                                Exit For
                            End If
                        Next fE
                        For jI = 1 To wb.Worksheets("FuncionFiltar").Cells(fE, wb.Worksheets("FuncionFiltar").Columns.Count).End(xlToLeft).Column
                            If Trim(wb.Worksheets("FuncionFiltar").Cells(fE, jI).Value) = nEt Then
                                iD_Col = jI
                                Exit For
                            End If
                        Next jI
                        If iD_Col > 0 Then
                            colC.Add iD_Col
                        End If
                    End If
                End If
            Next rr
            S_LimpiarTodo wb.Worksheets("FuncionFiltar"), colC, 1
        End If
    End If
    
    ' --- FASE DE FILAS ---
    Dim h2 As Worksheet
    On Error Resume Next
    Set h2 = wb.Worksheets("filas")
    On Error GoTo 0
    
    If Not h2 Is Nothing And pIdx > 0 Then
        Set colF = New Collection
        rM = h2.Cells(h2.Rows.Count, 6).End(xlUp).Row
        For rr = 3 To rM
            nEt = Trim(h2.Cells(rr, 6).Value)
            If nEt <> "" Then
                Dim fEnc As Long
                fEnc = F_DondeEstaFila(wb.Worksheets("TEXOENFILADOS"), nEt)
                If fEnc > 0 Then
                    If UCase(Trim(h2.Cells(rr, pIdx).Value)) = "NO" Then
                        colF.Add fEnc & "|DEL"
                    End If
                    If Trim(h2.Cells(rr, pIdx + 5).Value) <> "" And UCase(Trim(h2.Cells(rr, pIdx).Value)) <> "NO" Then
                        colF.Add fEnc & "|ADD|" & h2.Cells(rr, pIdx + 5).Value
                    End If
                End If
            End If
        Next rr
        S_LimpiarTodo wb.Worksheets("TEXOENFILADOS"), colF, 2
    End If

    ' --- LIMPIEZA Y CIERRE ---
    Application.DisplayAlerts = False
    wb.Worksheets("columnas").Delete
    wb.Worksheets("filas").Delete
    wb.SaveAs Filename:=fFinal, FileFormat:=51
    wb.Close 0
    
    If Dir(fTmp) <> "" Then
        Kill fTmp
    End If
    Application.DisplayAlerts = True
End Sub

' **********************************************************************************************************************
'  INICIO DE LA OPERACION
' **********************************************************************************************************************
Public Sub CrearExcelesSeparados()
    Dim ws As Worksheet
    Dim lista As New Collection
    Dim v As Variant
    Dim r As String
    Dim n As String
    Dim i As Integer
    Dim cM As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("columnas")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Hoja columnas no encontrada", 16
        Exit Sub
    End If
    
    r = "C:\CLIENTES\PRUEBAS\BP\"
    If Not F_VerificarRuta(r) Then
        MsgBox "No se pudo acceder a la ruta", 16
        Exit Sub
    End If
    
    n = Replace(ThisWorkbook.Name, ".xlsm", "")
    cM = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    
    For i = 3 To cM
        If Trim(ws.Cells(3, i).Value) <> "" Then
            lista.Add Trim(ws.Cells(3, i).Value)
        End If
    Next i
    
    If lista.Count = 0 Then
        Exit Sub
    End If
    
    ' CONFIGURACION DE VELOCIDAD
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = -4135
    
    For Each v In lista
        S_Proceso_Interno ws, CStr(v), r, n
    Next v
    
    Application.Calculation = -4105
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "PROCESO COMPLETADO", 64
End Sub
