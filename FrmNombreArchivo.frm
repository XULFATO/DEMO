VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmNombreArchivo
   Caption         =   "Archivo de salida"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
   Begin Forms.CommandButton BtnExplorador
      Caption         =   "..."
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   1560
      Width           =   570
   End
   Begin Forms.CommandButton BtnCancelar
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   2280
      Width           =   1350
   End
   Begin Forms.CommandButton BtnAceptar
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   2280
      Width           =   1350
   End
   Begin Forms.TextBox TxtNombre
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   6120
   End
   Begin Forms.Label LblInfo
      Height          =   1320
      Left            =   240
      Top             =   120
      Width           =   6840
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmNombreArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Resultado accesible desde fuera del form
Public NombreElegido  As String   ' Nombre de fichero sin extensión
Public CarpetaElegida As String   ' Carpeta final (puede cambiar si el usuario explora)
Public Cancelado      As Boolean

' Carpeta de partida para el explorador
Private m_rutaBase As String

' ======================================================================================
' INICIALIZACIÓN — llamar antes de .Show
' ======================================================================================

Public Sub Inicializar(ByVal rutaBase As String, _
                       ByVal nombreSugerido As String, _
                       ByVal textoInfo As String)
    m_rutaBase        = rutaBase
    CarpetaElegida    = rutaBase
    TxtNombre.Value   = nombreSugerido
    LblInfo.Caption   = textoInfo
    Cancelado         = True   ' Por defecto cancelado hasta que pulse Aceptar
End Sub

' ======================================================================================
' BOTÓN ACEPTAR
' ======================================================================================

Private Sub BtnAceptar_Click()
    Dim nombre As String
    nombre = Trim(TxtNombre.Value)

    ' Quitar extensión por si el usuario la escribió
    nombre = QuitarExtension(nombre)

    If nombre = "" Then
        MsgBox "El nombre no puede estar vacío.", vbExclamation
        TxtNombre.SetFocus
        Exit Sub
    End If

    NombreElegido = nombre
    Cancelado     = False
    Me.Hide
End Sub

' ======================================================================================
' BOTÓN CANCELAR
' ======================================================================================

Private Sub BtnCancelar_Click()
    Cancelado = True
    Me.Hide
End Sub

' ======================================================================================
' BOTÓN EXPLORADOR (...)
' Permite cambiar la carpeta destino; si la cambia actualiza CarpetaElegida
' y recalcula el nombre sugerido con la versión correcta para esa nueva carpeta.
' ======================================================================================

Private Sub BtnExplorador_Click()
    Dim shellApp    As Object
    Dim shellFolder As Object

    On Error Resume Next
    Set shellApp    = CreateObject("Shell.Application")
    Set shellFolder = shellApp.BrowseForFolder( _
                          0, "Seleccione la carpeta de destino:", 0, CarpetaElegida)
    On Error GoTo 0

    If Not shellFolder Is Nothing Then
        Dim nuevaCarpeta As String
        nuevaCarpeta = shellFolder.Self.Path
        If Right(nuevaCarpeta, 1) <> "\" Then nuevaCarpeta = nuevaCarpeta & "\"

        If UCase(Trim(nuevaCarpeta)) <> UCase(Trim(CarpetaElegida)) Then
            CarpetaElegida = nuevaCarpeta
            ' Actualizar label con la nueva carpeta
            Dim lineas() As String
            lineas = Split(LblInfo.Caption, vbCrLf)
            If UBound(lineas) >= 0 Then
                lineas(0) = "Carpeta: " & nuevaCarpeta
                LblInfo.Caption = Join(lineas, vbCrLf)
            End If
        End If
    End If

    Set shellFolder = Nothing
    Set shellApp    = Nothing
End Sub

' ======================================================================================
' CERRAR CON LA X = Cancelar
' ======================================================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then   ' Usuario cerró con la X
        Cancelado = True
        Cancel    = False
    End If
End Sub

' ======================================================================================
' UTILIDAD LOCAL
' ======================================================================================

Private Function QuitarExtension(ByVal nombre As String) As String
    Dim exts As Variant
    Dim e    As Variant
    exts = Array(".xlsx", ".xlsm", ".xls")
    For Each e In exts
        If LCase(Right(nombre, Len(e))) = LCase(e) Then
            nombre = Left(nombre, Len(nombre) - Len(e))
            Exit For
        End If
    Next e
    QuitarExtension = nombre
End Function
