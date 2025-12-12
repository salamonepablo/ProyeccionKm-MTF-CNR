VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormConsKmsCPG 
   Caption         =   "Km Ultima Intervención"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExportExcel 
      Caption         =   "Exportar a &Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   12
      Top             =   9600
      Width           =   1935
   End
   Begin VB.TextBox txtEditGrid 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Height          =   1695
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   19935
      Begin VB.TextBox txtKmsPromedioTemporada 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Text            =   "12500"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdProyCNR 
         Caption         =   "Ver &Kms 0Km / A3 / A2 / A1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8400
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtProyeccionMeses 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   3
         Text            =   "16"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtKmsPromedioMes 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         Text            =   "7500"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdProyeccionABC 
         Caption         =   "Ver &Kms ABC / RP/ RG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Proyección Meses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6405
         TabIndex        =   11
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Km. Tmpda."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5055
         TabIndex        =   10
         Top             =   600
         Width           =   1035
      End
      Begin VB.Image Image2 
         Height          =   945
         Left            =   17760
         Picture         =   "ConsKmCCRRCPG.frx":0000
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1845
      End
      Begin VB.Image Image1 
         Height          =   945
         Left            =   14880
         Picture         =   "ConsKmCCRRCPG.frx":265C9
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1845
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km. Prom."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3600
         TabIndex        =   8
         Top             =   600
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdCopiar 
      Caption         =   "&Copiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   9600
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid FG1 
      Height          =   7500
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   19935
      _ExtentX        =   35163
      _ExtentY        =   13229
      _Version        =   393216
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "FormConsKmsCPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private colCochesRemolcados As Collection

Private Function ObtenerMinFechaKm(ByVal coche As Variant, ByRef db As DAO.Database) As Variant
' Esta función devuelve la fecha MÁS ANTIGUA de kilometraje de un coche.
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT MIN(Fecha) AS MinFecha FROM Kilometraje WHERE Coche = '" & coche & "'"
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        ObtenerMinFechaKm = rs!MinFecha
    Else
        ObtenerMinFechaKm = Null
    End If
    
    rs.Close
    Set rs = Nothing
End Function ' En la sección de declaraciones generales de tu formulario

Private Function CalcularKmTotal(ByVal coche As Variant, ByRef db As DAO.Database) As Double
' Esta función calcula la suma TOTAL de kilómetros para un coche desde el inicio.
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT SUM(Kms_diario) AS SumaTotal FROM Kilometraje WHERE Coche = '" & coche & "'"
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        If IsNull(rs!SumaTotal) Then
            CalcularKmTotal = 0
        Else
            CalcularKmTotal = rs!SumaTotal
        End If
    Else
        CalcularKmTotal = 0
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub RecalcularFila(ByVal fila As Long, ByVal colInicio As Long)
' Recalcula una fila del grid desde una columna específica hacia la derecha.
' (Esta es una versión mejorada de la que teníamos)

    Dim i As Long
    Dim kmsPromedioMes As Double
    Dim valorAnterior As Double
    Dim tipoIntervencion As String
    
    kmsPromedioMes = Val(txtKmsPromedioMes.text)
    tipoIntervencion = FG1.TextMatrix(fila, 1)
    
    ' Bucle que empieza en la columna SIGUIENTE a la que se modificó/reseteó
    For i = colInicio + 1 To FG1.Cols - 1
        ' Limpiamos el valor de la celda anterior para evitar errores con los separadores de miles
        valorAnterior = Val(Replace(FG1.TextMatrix(fila, i - 1), ".", ""))
        
        ' Calcula el nuevo valor, lo pone en la celda y le da formato
        FG1.TextMatrix(fila, i) = Format(valorAnterior + kmsPromedioMes, "#,##0")
        FormatearCelda fila, i, valorAnterior + kmsPromedioMes, tipoIntervencion
    Next i
    
    ' Limpiamos el formato de la celda donde se aplicó la intervención (que ahora es 0)
    FormatearCelda fila, colInicio, 0, tipoIntervencion
End Sub
Private Sub AplicarIntervencion(ByVal fila As Long, ByVal col As Long, ByVal tipoIntervencion As String)
' Esta es la nueva subrutina principal para manejar la lógica en cascada.
    
    Dim nuevaFecha As Date
    Dim kmsPromedioMes As Double
    
    kmsPromedioMes = Val(txtKmsPromedioMes.text)
    
    ' Obtenemos la fecha del encabezado de la columna donde se hizo la intervención
    nuevaFecha = CDate("01 " & FG1.TextMatrix(0, col))
    
    ' --- Lógica en Cascada ---
    Select Case tipoIntervencion
        Case "RG"
            ' --- 1. Resetea la fila RG ---
            FG1.TextMatrix(fila, 2) = Format(nuevaFecha, "dd/mm/yyyy") ' Actualiza fecha
            FG1.TextMatrix(fila, col) = tipoIntervencion ' <-- MEJORA: Pone "RG" en la celda
            ' Le damos un formato especial para que se note que es una intervención manual
            FG1.row = fila: FG1.col = col
            FG1.cellBackColor = &HD8D8D8 ' Gris para destacar
            FG1.cellFontBold = True
            FG1.cellForeColor = vbBlack
            
            ' --- 2. Resetea la fila RP (fila + 1) ---
            FG1.TextMatrix(fila + 1, 2) = Format(nuevaFecha, "dd/mm/yyyy")
            FG1.TextMatrix(fila + 1, col) = "RG" ' Muestra que fue reseteada por una RG
            FG1.row = fila + 1: FG1.col = col
            FG1.cellBackColor = &HD8D8D8
            FG1.cellFontBold = True
            FG1.cellForeColor = vbBlack
            
            ' --- 3. Resetea la fila ABC (fila + 2) ---
            FG1.TextMatrix(fila + 2, 2) = Format(nuevaFecha, "dd/mm/yyyy")
            FG1.TextMatrix(fila + 2, col) = "RG" ' Muestra que fue reseteada por una RG
            FG1.row = fila + 2: FG1.col = col
            FG1.cellBackColor = &HD8D8D8
            FG1.cellFontBold = True
            FG1.cellForeColor = vbBlack
            
            ' --- 4. Inicia la proyección y recalcula las 3 filas ---
            If col + 1 < FG1.Cols Then ' Nos aseguramos de no estar en la última columna
                ' Ponemos el primer valor de la proyección en la celda de la derecha
                FG1.TextMatrix(fila, col + 1) = Format(kmsPromedioMes, "#,##0")
                FG1.TextMatrix(fila + 1, col + 1) = Format(kmsPromedioMes, "#,##0")
                FG1.TextMatrix(fila + 2, col + 1) = Format(kmsPromedioMes, "#,##0")
                
                ' Recalculamos el resto de la fila a partir de esa nueva celda
                RecalcularFila fila, col + 1
                RecalcularFila fila + 1, col + 1
                RecalcularFila fila + 2, col + 1
            End If
            
        Case "RP"
            ' (Lógica similar para RP)
            FG1.TextMatrix(fila, 2) = Format(nuevaFecha, "dd/mm/yyyy")
            FG1.TextMatrix(fila, col) = tipoIntervencion
            FG1.row = fila: FG1.col = col: FG1.cellBackColor = &HD8D8D8: FG1.cellFontBold = True: FG1.cellForeColor = vbBlack
            
            FG1.TextMatrix(fila + 1, 2) = Format(nuevaFecha, "dd/mm/yyyy")
            FG1.TextMatrix(fila + 1, col) = "RP"
            FG1.row = fila + 1: FG1.col = col: FG1.cellBackColor = &HD8D8D8: FG1.cellFontBold = True: FG1.cellForeColor = vbBlack
            
            If col + 1 < FG1.Cols Then
                FG1.TextMatrix(fila, col + 1) = Format(kmsPromedioMes, "#,##0")
                FG1.TextMatrix(fila + 1, col + 1) = Format(kmsPromedioMes, "#,##0")
                RecalcularFila fila, col + 1
                RecalcularFila fila + 1, col + 1
            End If
            
        Case "ABC"
            ' (Lógica similar para ABC)
            FG1.TextMatrix(fila, 2) = Format(nuevaFecha, "dd/mm/yyyy")
            FG1.TextMatrix(fila, col) = tipoIntervencion
            FG1.row = fila: FG1.col = col: FG1.cellBackColor = &HD8D8D8: FG1.cellFontBold = True: FG1.cellForeColor = vbBlack
            
            If col + 1 < FG1.Cols Then
                FG1.TextMatrix(fila, col + 1) = Format(kmsPromedioMes, "#,##0")
                RecalcularFila fila, col + 1
            End If
    End Select
End Sub


Private Sub AplicarFormatoCondicional(ByVal pFila As Long, ByVal pColumna As Long, _
                                      ByVal pValorKm As Double, ByVal pTipoIntervencion As String)
    With FG1
        .row = pFila
        .col = pColumna
        
        Select Case pTipoIntervencion
            Case "0 Km"
                If pValorKm > 2400000 Then
                    .cellBackColor = RGB(255, 220, 220) ' Rosa claro
                    .cellForeColor = RGB(255, 0, 0)   ' Rojo
                Else
                    .cellBackColor = .BackColor       ' Restaurar color de fondo
                    .cellForeColor = .ForeColor       ' Restaurar color de texto
                End If
            Case "A3"
                If pValorKm > 800000 Then
                    .cellBackColor = RGB(220, 230, 255) ' Azul claro
                    .cellForeColor = RGB(0, 0, 150)   ' Azul oscuro
                Else
                    .cellBackColor = .BackColor
                    .cellForeColor = .ForeColor
                End If
            Case "A2"
                If pValorKm > 400000 Then
                    .cellBackColor = RGB(255, 255, 200) ' Amarillo claro
                    .cellForeColor = RGB(150, 150, 0)   ' Amarillo oscuro
                Else
                    .cellBackColor = .BackColor
                    .cellForeColor = .ForeColor
                End If
            Case "A1"
                If pValorKm > 200000 Then
                    .cellBackColor = RGB(220, 255, 220) ' Verde claro
                    .cellForeColor = RGB(0, 150, 0)   ' Verde oscuro
                Else
                    .cellBackColor = .BackColor
                    .cellForeColor = .ForeColor
                End If
        End Select
    End With
End Sub
Private Sub cmdExportExcel_Click()
    Dim xlApp As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim r As Long ' Row index for grid
    Dim c As Long ' Column index for grid
    
    Dim cellText As String
    Dim cellBackColor As Long
    Dim cellForeColor As Long
    Dim cellFontBold As Boolean
    Dim gridBackColor As Long ' Para guardar el color de fondo por defecto del grid
    Dim numericValue As Double ' Para guardar el valor numérico
    
    If FG1.Rows <= 1 Then
        MsgBox "La grilla está vacía. No hay nada para exportar.", vbInformation
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler

    Set xlApp = New Excel.Application
    Set xlWB = xlApp.Workbooks.Add
    Set xlSheet = xlWB.Worksheets(1)
    
    xlApp.Visible = False
    
    ' Guardamos el color de fondo por defecto del grid
    gridBackColor = FG1.BackColor
    
    With FG1
        For r = 0 To .Rows - 1
            For c = 0 To .Cols - 1
                
                ' Leer propiedades del grid
                .row = r
                .col = c
                cellText = .text
                cellBackColor = .cellBackColor
                cellForeColor = .cellForeColor
                cellFontBold = .cellFontBold
                
                ' Escribir y formatear celda de Excel
                With xlSheet.Cells(r + 1, c + 1)
                
                    ' --- MEJORA 2: Transferir NÚMERO o TEXTO según la columna ---
                    If (c = 1 Or c >= 3) And r > 0 Then ' Si es columna de KM (1 o >=3) Y no es la fila de encabezado
                        ' Convertimos el texto formateado a número antes de pasarlo a Excel
                        numericValue = Val(Replace(cellText, ".", ""))
                        .Value = numericValue
                        .NumberFormat = "#,##0" ' Aplicamos formato de miles en Excel
                    Else
                        ' Para encabezados y columnas de texto, pasamos el texto tal cual
                        .Value = cellText
                        .NumberFormat = "@" ' Formato de texto
                    End If
                    ' -----------------------------------------------------------

                    ' --- MEJORA 1: Manejar color de fondo por defecto ---
                    If cellBackColor = gridBackColor Then
                        ' Si la celda tiene el color de fondo por defecto del grid,
                        ' le ponemos fondo blanco en Excel.
                        .Interior.Color = vbWhite
                    Else
                        ' Si tiene un color específico (formato condicional), lo aplicamos.
                        .Interior.Color = cellBackColor
                    End If
                    ' ----------------------------------------------------

                    .Font.Color = cellForeColor
                    .Font.Bold = cellFontBold
                    
                    ' Alineación (sin cambios)
                    Select Case c
                        Case 0, 2, 3 ' Módulo, Fecha, Tipo
                             .HorizontalAlignment = xlCenter
                        Case 1 ' Km Acumulados (y proyecciones implícitas)
                             .HorizontalAlignment = xlRight
                    End Select
                     ' El formato de número ya se aplicó arriba
                End With
            Next c
        Next r
    End With

    xlSheet.Columns.AutoFit
    xlApp.Visible = True
    xlApp.UserControl = True
    
    Set xlSheet = Nothing
    Set xlWB = Nothing
    ' No xlApp = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error al exportar a Excel:" & vbCrLf & Err.Description & vbCrLf & _
           "Asegúrese de que Microsoft Excel esté instalado correctamente.", vbCritical, "Error de Exportación"
    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If
    Set xlSheet = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
End Sub
Private Sub cmdProyCNR_Click()
    
    CargarCochesCNR
    
    ' --- Validaciones iniciales ---
    If Not IsNumeric(txtKmsPromedioMes.text) Or Val(txtKmsPromedioMes.text) <= 0 Then
        MsgBox "Por favor, ingrese un valor numérico válido para los Kms Promedio por Mes.", vbExclamation
        txtKmsPromedioMes.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtProyeccionMeses.text) Or Val(txtProyeccionMeses.text) <= 0 Then
        MsgBox "Por favor, ingrese una cantidad de meses válida para la proyección.", vbExclamation
        txtProyeccionMeses.SetFocus
        Exit Sub
    End If
    ' --- NUEVA VALIDACIÓN: Kms Promedio por Temporada ---
    If Not IsNumeric(txtKmsPromedioTemporada.text) Or Val(txtKmsPromedioTemporada.text) <= 0 Then
        MsgBox "Por favor, ingrese un valor numérico válido para los Kms Promedio por Temporada.", vbExclamation
        txtKmsPromedioTemporada.SetFocus
        Exit Sub
    End If

    Call CargarCochesCNR
    
    If colCochesRemolcados Is Nothing Or colCochesRemolcados.Count = 0 Then
        MsgBox "No hay coches CNR cargados en la lista para procesar.", vbExclamation
        Exit Sub
    End If

    ' --- Declaración de variables ---
    Dim db As DAO.Database
    Dim numCoche As Variant
    Dim rutaBase As String, clave As String
    Dim kmsPromedioMes As Double
    Dim kmsPromedioTemporada As Double ' NUEVA VARIABLE
    Dim mesesProyeccion As Integer
    Dim filaActual As Integer
    Dim i As Integer
    
    kmsPromedioMes = Val(txtKmsPromedioMes.text)
    kmsPromedioTemporada = Val(txtKmsPromedioTemporada.text) ' ASIGNAR NUEVA VARIABLE
    mesesProyeccion = CInt(txtProyeccionMeses.text)
    filaActual = 1

    ' --- Preparación del Grid (el encabezado de meses debe seguir siendo dinámico) ---
    With FG1
        .Clear
        .Cols = 4 + mesesProyeccion
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "N° Coche"
        .TextMatrix(0, 1) = "Intervención"
        .TextMatrix(0, 2) = "Fecha"
        .TextMatrix(0, 3) = "Km Acumulado"
        
        ' --- Modificación aquí: El encabezado mostrará el mes real de la proyección ---
        For i = 1 To mesesProyeccion
            .TextMatrix(0, 3 + i) = Format(DateAdd("m", i, Date), "mmm yy")
        Next i
        
        .row = 0: .RowSel = 0: .col = 0: .ColSel = .Cols - 1: .cellFontBold = True
        ' (Ya tienes FG1.Height = 5500 o similar en Form_Load, si no, ponlo ahí)
    End With

    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCRR.mdb"
    clave = "theidol-1995"

    On Error GoTo ErrorHandler
    Set db = OpenDatabase(rutaBase, False, False, "MS Access;PWD=" & clave)
    On Error GoTo 0

    For Each numCoche In colCochesRemolcados
    
        ' --- Variables para la lógica de cálculo (sin cambios en declaraciones aquí) ---
        Dim kmTotal As Double, fechaInicio As Variant
        Dim fechaA3 As Variant, fechaA2 As Variant, fechaA1 As Variant
        Dim kmA3 As Double, kmA2 As Double, kmA1 As Double
        Dim fechaEfectivaA3 As Variant, fechaEfectivaA2 As Variant, fechaEfectivaA1 As Variant
        Dim proyKmTotal As Double, proyKmA3 As Double, proyKmA2 As Double, proyKmA1 As Double ' Agregamos estas declaraciones si no estaban

        ' --- PASO 1: OBTENER DATOS BASE Y FECHAS DE INTERVENCIÓN (sin cambios) ---
        kmTotal = CalcularKmTotal(numCoche, db)
        fechaInicio = ObtenerMinFechaKm(numCoche, db)
        
        fechaA3 = ObtenerMaxFecha(numCoche, "A3", db)
        fechaA2 = ObtenerMaxFecha(numCoche, "A2", db)
        fechaA1 = ObtenerMaxFecha(numCoche, "A1", db)
        
        ' --- PASO 2: APLICAR LÓGICA DE JERARQUÍA MEJORADA Y SEGURA (sin cambios) ---
        ' (Mantener el código de A3, A2, A1 tal como te lo pasé en la última corrección del error 94)
        
        ' --- Cálculo para A3 ---
        If Not IsNull(fechaA3) Then
            fechaEfectivaA3 = fechaA3
            kmA3 = CalcularSumaKm(numCoche, fechaEfectivaA3, db)
        Else
            fechaEfectivaA3 = fechaInicio
            kmA3 = kmTotal
        End If

        ' --- Cálculo para A2 ---
        If Not IsNull(fechaA2) Then
            If Not IsNull(fechaEfectivaA3) Then
                If CDate(fechaA2) > CDate(fechaEfectivaA3) Then
                    fechaEfectivaA2 = fechaA2
                Else
                    fechaEfectivaA2 = fechaEfectivaA3
                End If
            Else
                fechaEfectivaA2 = fechaA2
            End If
        Else
            fechaEfectivaA2 = fechaEfectivaA3
        End If
        If Not IsNull(fechaEfectivaA2) Then
            kmA2 = CalcularSumaKm(numCoche, fechaEfectivaA2, db)
        Else
            kmA2 = kmTotal
        End If

        ' --- Cálculo para A1 ---
        If Not IsNull(fechaA1) Then
            If Not IsNull(fechaEfectivaA2) Then
                If CDate(fechaA1) > CDate(fechaEfectivaA2) Then
                    fechaEfectivaA1 = fechaA1
                Else
                    fechaEfectivaA1 = fechaEfectivaA2
                End If
            Else
                fechaEfectivaA1 = fechaA1
            End If
        Else
            fechaEfectivaA1 = fechaEfectivaA2
        End If
        If Not IsNull(fechaEfectivaA1) Then
            kmA1 = CalcularSumaKm(numCoche, fechaEfectivaA1, db)
        Else
            kmA1 = kmTotal
        End If

        ' --- PASO 3: LLENAR EL GRID CON LAS 4 FILAS (sin cambios en los datos iniciales) ---
        FG1.Rows = FG1.Rows + 4
        
        FG1.TextMatrix(filaActual, 0) = numCoche
        FG1.TextMatrix(filaActual, 1) = "0 Km"
        FG1.TextMatrix(filaActual, 2) = IIf(IsNull(fechaInicio), "N/A", Format(fechaInicio, "dd/mm/yyyy"))
        FG1.TextMatrix(filaActual, 3) = Format(kmTotal, "#,##0")
        
        FG1.TextMatrix(filaActual + 1, 0) = numCoche
        FG1.TextMatrix(filaActual + 1, 1) = "A3"
        FG1.TextMatrix(filaActual + 1, 2) = IIf(IsNull(fechaEfectivaA3), "N/A", Format(fechaEfectivaA3, "dd/mm/yyyy"))
        FG1.TextMatrix(filaActual + 1, 3) = Format(kmA3, "#,##0")
        
        FG1.TextMatrix(filaActual + 2, 0) = numCoche
        FG1.TextMatrix(filaActual + 2, 1) = "A2"
        FG1.TextMatrix(filaActual + 2, 2) = IIf(IsNull(fechaEfectivaA2), "N/A", Format(fechaEfectivaA2, "dd/mm/yyyy"))
        FG1.TextMatrix(filaActual + 2, 3) = Format(kmA2, "#,##0")
        
        FG1.TextMatrix(filaActual + 3, 0) = numCoche
        FG1.TextMatrix(filaActual + 3, 1) = "A1"
        FG1.TextMatrix(filaActual + 3, 2) = IIf(IsNull(fechaEfectivaA1), "N/A", Format(fechaEfectivaA1, "dd/mm/yyyy"))
        FG1.TextMatrix(filaActual + 3, 3) = Format(kmA1, "#,##0")
        
        ' --- PASO 4: EJECUTAR LA PROYECCIÓN CON KILOMETRAJE DE TEMPORADA ---
        proyKmTotal = kmTotal
        proyKmA3 = kmA3
        proyKmA2 = kmA2
        proyKmA1 = kmA1
        
        For i = 1 To mesesProyeccion
            Dim mesProyeccion As Integer
            mesProyeccion = Month(DateAdd("m", i, Date)) ' Obtiene el número de mes (1 a 12)
            
            Dim kmAConsiderar As Double
            If mesProyeccion >= 12 Or (mesProyeccion >= 1 And mesProyeccion <= 3) Then ' Diciembre, Enero, Febrero, Marzo
                kmAConsiderar = kmsPromedioTemporada
            Else
                kmAConsiderar = kmsPromedioMes
            End If
            
            proyKmTotal = proyKmTotal + kmAConsiderar
            proyKmA3 = proyKmA3 + kmAConsiderar
            proyKmA2 = proyKmA2 + kmAConsiderar
            proyKmA1 = proyKmA1 + kmAConsiderar
            
            FG1.TextMatrix(filaActual, 3 + i) = Format(proyKmTotal, "#,##0")
            FG1.TextMatrix(filaActual + 1, 3 + i) = Format(proyKmA3, "#,##0")
            FG1.TextMatrix(filaActual + 2, 3 + i) = Format(proyKmA2, "#,##0")
            FG1.TextMatrix(filaActual + 3, 3 + i) = Format(proyKmA1, "#,##0")
            
            ' --- APLICAR FORMATO CONDICIONAL DESPUÉS DE CALCULAR Y ASIGNAR EL VALOR ---
            Call AplicarFormatoCondicional(filaActual, 3 + i, proyKmTotal, "0 Km")
            Call AplicarFormatoCondicional(filaActual + 1, 3 + i, proyKmA3, "A3")
            Call AplicarFormatoCondicional(filaActual + 2, 3 + i, proyKmA2, "A2")
            Call AplicarFormatoCondicional(filaActual + 3, 3 + i, proyKmA1, "A1")

        Next i
        
        filaActual = filaActual + 4
    Next numCoche

    MsgBox "Proceso y proyección para coches CNR finalizado.", vbInformation
    
GoTo Limpieza
ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
Limpieza:
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
End Sub
Private Sub txtEditGrid_LostFocus()
' Se activa cuando el TextBox pierde el foco. Ahora diferencia entre números y siglas.
    
    Dim filaEditada As Long, colEditada As Long
    Dim textoIngresado As String
    Dim tipoFila As String
    
    txtEditGrid.Visible = False
    filaEditada = FG1.row
    colEditada = FG1.col
    
    ' Convertimos a mayúsculas y quitamos espacios para una comparación segura
    textoIngresado = UCase(Trim(txtEditGrid.text))
    
    ' Verificamos si el texto ingresado es una de las siglas de intervención
    If textoIngresado = "RG" Or textoIngresado = "RP" Or textoIngresado = "ABC" Then
        ' --- LÓGICA NUEVA PARA INTERVENCIONES ---
        
        ' Obtenemos el tipo de la fila ("RG", "RP" o "ABC") desde la columna 1
        tipoFila = FG1.TextMatrix(filaEditada, 1)
        
        ' Solo aplicamos la intervención si coincide con el tipo de la fila
        If textoIngresado = tipoFila Then
            AplicarIntervencion filaEditada, colEditada, textoIngresado
        Else
            MsgBox "No puede ingresar una intervención '" & textoIngresado & "' en una fila de tipo '" & tipoFila & "'.", vbExclamation, "Entrada no válida"
        End If
        
    ElseIf IsNumeric(Replace(textoIngresado, ".", "")) Then
        ' --- LÓGICA ANTIGUA PARA EDITAR NÚMEROS ---
        
        Dim valorLimpio As Double
        valorLimpio = Val(Replace(textoIngresado, ".", ""))
        
        FG1.TextMatrix(filaEditada, colEditada) = Format(valorLimpio, "#,##0")
        RecalcularFila filaEditada, colEditada
        
    Else
        ' Si no es ni una sigla válida ni un número, no hacemos nada.
    End If
End Sub

Private Sub CargarCochesFijos()
    ' Inicializa la colección para empezar de cero
    Set colCochesRemolcados = New Collection

    ' Añade cada número de coche a la colección
    ' Simplemente ponés un .Add por cada número
        colCochesRemolcados.Add 2521
        colCochesRemolcados.Add 2524
        colCochesRemolcados.Add 2526
        colCochesRemolcados.Add 2530
        colCochesRemolcados.Add 2533
        colCochesRemolcados.Add 2555
        colCochesRemolcados.Add 2569
        colCochesRemolcados.Add 2571
        colCochesRemolcados.Add 2588
        colCochesRemolcados.Add 2591
        colCochesRemolcados.Add 2597
        colCochesRemolcados.Add 2598
        colCochesRemolcados.Add 2612
        colCochesRemolcados.Add 2615
        colCochesRemolcados.Add 2621
        colCochesRemolcados.Add 2624
        colCochesRemolcados.Add 2626
        colCochesRemolcados.Add 2632
        colCochesRemolcados.Add 3019
        colCochesRemolcados.Add 3038
        colCochesRemolcados.Add 3203
        colCochesRemolcados.Add 3501
        colCochesRemolcados.Add 3503
        colCochesRemolcados.Add 3504
        colCochesRemolcados.Add 3509
        colCochesRemolcados.Add 3515
        colCochesRemolcados.Add 3524
        colCochesRemolcados.Add 3529
        colCochesRemolcados.Add 3534
        colCochesRemolcados.Add 3545
        colCochesRemolcados.Add 3547
        colCochesRemolcados.Add 3558
        colCochesRemolcados.Add 3566
        colCochesRemolcados.Add 3570
        colCochesRemolcados.Add 3580
        colCochesRemolcados.Add 3601
        colCochesRemolcados.Add 3606
        colCochesRemolcados.Add 3617
        colCochesRemolcados.Add 3618
        colCochesRemolcados.Add 3620
        colCochesRemolcados.Add 3646
        colCochesRemolcados.Add 3655
        colCochesRemolcados.Add 3675
        colCochesRemolcados.Add 3678
        colCochesRemolcados.Add 3685
        colCochesRemolcados.Add 3694
        colCochesRemolcados.Add 3701
        colCochesRemolcados.Add 3715
        colCochesRemolcados.Add 3716
        colCochesRemolcados.Add 3717
        colCochesRemolcados.Add 3727
        colCochesRemolcados.Add 3734
        colCochesRemolcados.Add 3738
        colCochesRemolcados.Add 3749
        colCochesRemolcados.Add 3753
        colCochesRemolcados.Add 3762
        colCochesRemolcados.Add 3765
        colCochesRemolcados.Add 3768
        colCochesRemolcados.Add 3770
        colCochesRemolcados.Add 3811
    
    ' ...y así sucesivamente hasta completar los 60...

    'MsgBox "Se cargaron " & colCochesRemolcados.Count & " coches fijos en memoria."
    
End Sub
Private Sub CargoComboMes()

'   cmbMes.AddItem ("Enero")
'   cmbMes.AddItem ("Febrero")
'   cmbMes.AddItem ("Marzo")
'   cmbMes.AddItem ("Abril")
'   cmbMes.AddItem ("Mayo")
'   cmbMes.AddItem ("Junio")
'   cmbMes.AddItem ("Julio")
'   cmbMes.AddItem ("Agosto")
'   cmbMes.AddItem ("Septiembre")
'   cmbMes.AddItem ("Octubre")
'   cmbMes.AddItem ("Noviembre")
'   cmbMes.AddItem ("Diciembre")
   
End Sub

Private Sub CargarCochesCNR()

' Crea una colección con los coches remolcados CNR a procesar.
    
    ' Inicializa la colección para empezar de cero
    Set colCochesRemolcados = New Collection

    ' Lista completa de coches CNR
    colCochesRemolcados.Add "CDA004"
    colCochesRemolcados.Add "CDA006"
    colCochesRemolcados.Add "CPA002"
    colCochesRemolcados.Add "CPA003"
    colCochesRemolcados.Add "CPA005"
    colCochesRemolcados.Add "CPA006"
    colCochesRemolcados.Add "CPA013"
    colCochesRemolcados.Add "CPA014"
    colCochesRemolcados.Add "CPA018"
    colCochesRemolcados.Add "CPA019"
    colCochesRemolcados.Add "CPA020"
    colCochesRemolcados.Add "CPA025"
    colCochesRemolcados.Add "CPA026"
    colCochesRemolcados.Add "CPA027"
    colCochesRemolcados.Add "CPA031"
    colCochesRemolcados.Add "CPA033"
    colCochesRemolcados.Add "CPA034"
    colCochesRemolcados.Add "CPA035"
    colCochesRemolcados.Add "CPA037"
    colCochesRemolcados.Add "CPA039"
    colCochesRemolcados.Add "CPA041"
    colCochesRemolcados.Add "CPA042"
    colCochesRemolcados.Add "CPA044"
    colCochesRemolcados.Add "CPA045"
    colCochesRemolcados.Add "CPA046"
    colCochesRemolcados.Add "CPA058"
    colCochesRemolcados.Add "CPA060"
    colCochesRemolcados.Add "CPA063"
    colCochesRemolcados.Add "CPA064"
    colCochesRemolcados.Add "CPA068"
    colCochesRemolcados.Add "CPA070"
    colCochesRemolcados.Add "CPA072"
    colCochesRemolcados.Add "CPA073"
    colCochesRemolcados.Add "CPA079"
    colCochesRemolcados.Add "CPA080"
    colCochesRemolcados.Add "CPA081"
    colCochesRemolcados.Add "CPA082"
    colCochesRemolcados.Add "CPA083"
    colCochesRemolcados.Add "CPA084"
    colCochesRemolcados.Add "CPA085"
    colCochesRemolcados.Add "CPA086"
    colCochesRemolcados.Add "CPA087"
    colCochesRemolcados.Add "CPA088"
    colCochesRemolcados.Add "CPA089"
    colCochesRemolcados.Add "CPA090"
    colCochesRemolcados.Add "CRA003"
    colCochesRemolcados.Add "CRA008"
    colCochesRemolcados.Add "CRA009"
    colCochesRemolcados.Add "CRA010"
    colCochesRemolcados.Add "CRA016"
    colCochesRemolcados.Add "CRA017"

    colCochesRemolcados.Add "CRA018"
    colCochesRemolcados.Add "FG001"
    colCochesRemolcados.Add "FG008"
    colCochesRemolcados.Add "FG009"
    colCochesRemolcados.Add "FG010"
    colCochesRemolcados.Add "FG012"
    colCochesRemolcados.Add "FG017"
    colCochesRemolcados.Add "FG018"
    colCochesRemolcados.Add "FG019"
    colCochesRemolcados.Add "FS003"
    colCochesRemolcados.Add "FS005"
    colCochesRemolcados.Add "FS007"
    colCochesRemolcados.Add "FS011"
    colCochesRemolcados.Add "FS013"
    colCochesRemolcados.Add "FS017"
    colCochesRemolcados.Add "FS019"
    colCochesRemolcados.Add "FS020"
    colCochesRemolcados.Add "PUA003"
    colCochesRemolcados.Add "PUA004"
    colCochesRemolcados.Add "PUA015"
    colCochesRemolcados.Add "PUA016"
    colCochesRemolcados.Add "PUA018"
    colCochesRemolcados.Add "PUA019"
    colCochesRemolcados.Add "PUA021"
    colCochesRemolcados.Add "PUA022"
    colCochesRemolcados.Add "PUA023"
    colCochesRemolcados.Add "PUA026"
    colCochesRemolcados.Add "PUA027"
    colCochesRemolcados.Add "PUA028"
    colCochesRemolcados.Add "PUA036"
    colCochesRemolcados.Add "PUA037"
    colCochesRemolcados.Add "PUAD001"
    colCochesRemolcados.Add "PUAD003"
    colCochesRemolcados.Add "PUAD004"
    colCochesRemolcados.Add "PUAD005"
    colCochesRemolcados.Add "PUAD009"
    colCochesRemolcados.Add "PUAD010"
    colCochesRemolcados.Add "PUAD013"
    colCochesRemolcados.Add "PUAD020"

End Sub
Private Sub FormateoGrilla()
    
    ' --- INICIALIZACIÓN DEL MSFLEXGRID ---
    ' Esto va ANTES del bucle "For Each"
    Dim filaActual As Integer
    filaActual = 1 ' Empezamos en la fila 1, la 0 es para los encabezados
    
    With FG1
        .Clear ' Limpia el grid por si tenía datos anteriores
        .Cols = 4 ' Necesitamos 4 columnas
        .Rows = 1 ' Empezamos con una sola fila para los encabezados
        .FixedRows = 1 ' Deja la fila de encabezados fija al hacer scroll
    
        ' Definir los encabezados de las columnas
        .TextMatrix(0, 0) = "N° Coche"
        .TextMatrix(0, 1) = "Intervención"
        .TextMatrix(0, 2) = "Fecha Última ABC"
        .TextMatrix(0, 3) = "Km Acumulados"
    
        ' Ajustar el ancho de las columnas (el valor está en Twips)
        .ColWidth(0) = 1200 ' Ancho para N° Coche
        .ColWidth(1) = 1200 ' Ancho para Intervención
        .ColWidth(2) = 1200 ' Ancho para Fecha
        .ColWidth(3) = 1500 ' Ancho para Km
    End With

End Sub

Private Sub UltimaABC()

    Dim db1 As DAO.Database
    Dim ws As DAO.Workspace
    Dim rutaBase As String
    Dim clave As String
    Dim tQuery, tQuery2, vSQL, vSQL2
    Dim TotalKm As Double
    Dim FechaUABC As Date
    
    FechaUABC = "01/01/1900"
    TotalKm = 0
    
    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCRR.mdb"
    clave = "theidol-1995"

    ' Crear el workspace usando el motor Jet
    Set ws = DBEngine.Workspaces(0)

    On Error GoTo ErrorHandler

    ' Abrir la base de datos con contraseña
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    vSQL = "SELECT * FROM DETENCIONES WHERE Coche Like '" & cmbCCRR.text & "' AND  INTERVENCION='ABC' ORDER BY FECHA_HASTA"
    
    'MsgBox (vSQL)
    
    Set tQuery = db.OpenRecordset(vSQL)
    
    If Not tQuery.EOF Then
        tQuery.MoveLast
        lblInfo.Caption = "Trabajando con " & Format(tQuery.RecordCount, "#,##0") & " regitros de Intervenciones ABC"
     Else
    End If
    
    tQuery.MoveFirst
            
    While Not tQuery.EOF
        If tQuery!Fecha_hasta > FechaUABC Then FechaUABC = tQuery!Fecha_hasta
        tQuery.MoveNext
    Wend

    lblUltimaABC2.ForeColor = vbRed
    lblUltimaABC2.AutoSize = True
    lblUltimaABC2.FontBold = True
    lblUltimaABC2.FontSize = 10
    lblUltimaABC2.Caption = Format$(FechaUABC, "DD/MM/YYYY")
    lblUltimaABC2.Visible = True
    
    vSQL2 = "SELECT * FROM KILOMETRAJE WHERE Coche Like '" & cmbCCRR.text & "' AND  FECHA >=#" & Format$(FechaUABC, "MM/DD/YYYY") & "#"
    
    'MsgBox (vSQL2)
    
    Set tQuery2 = db.OpenRecordset(vSQL2)
    
    If Not tQuery2.EOF Then
        tQuery2.MoveLast
        lblInfo.Caption = "Trabajando con " & Format(tQuery2.RecordCount, "#,##0") & " regitros de Kms de ABC"
     Else
    End If
    
    If Not tQuery2.EOF Then tQuery2.MoveFirst
            
    While Not tQuery2.EOF
        TotalKm = TotalKm + tQuery2!Kms_Diario
        tQuery2.MoveNext
    Wend
    
    lblKmABC2.ForeColor = vbRed
    lblKmABC2.AutoSize = True
    lblKmABC2.FontBold = True
    lblKmABC2.FontSize = 10
    lblKmABC2.Caption = Format$(TotalKm, "Standard")
    lblKmABC2.Visible = True
    
    tQuery.Close
    tQuery2.Close
    db1.Close
    
ErrorHandler:
   
    Select Case Err
        Case 3021
            MsgBox "No se encuentra ABC para el vehículo consultado", vbCritical
            Resume Next
        'Case Else
        '   MsgBox "Err" & " " & Err.Description, vbCritical
    End Select
   
'    Set db = Nothing
'    Set ws = Nothing

End Sub
Private Sub UltimaRev()
    
    Dim db1 As DAO.Database
    Dim ws As DAO.Workspace
    Dim rutaBase As String
    Dim clave As String
    Dim tQuery, tQuery2, vSQL, vSQL2
    Dim TotalKm As Double
    Dim FechaURev As Date
    Dim TipoRev As String
    
    FechaURev = "01/01/1900"
    TotalKm = 0
    
    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCRR.mdb"
    clave = "theidol-1995"

    ' Crear el workspace usando el motor Jet
    Set ws = DBEngine.Workspaces(0)

    On Error GoTo ErrorHandler

    ' Abrir la base de datos con contraseña
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    'vSQL = "SELECT * FROM DETENCIONES WHERE Coche Like '" & cmbCCRR.text & "' AND  INTERVENCION='A' OR INTERVENCION='AB' OR INTERVENCION='E' ORDER BY FECHA_HASTA"
    vSQL = "SELECT * FROM DETENCIONES WHERE Coche Like '" & cmbCCRR.text & "'  ORDER BY FECHA_HASTA"
    
    'MsgBox (vSQL)
    
    Set tQuery = db.OpenRecordset(vSQL)
    
    If Not tQuery.EOF Then
        tQuery.MoveLast
        lblInfo.Caption = "Trabajando con " & Format(tQuery.RecordCount, "#,##0") & " regitros de Intervenciones Periódicas"
     Else
    End If
    
    tQuery.MoveFirst
            
    While Not tQuery.EOF
        Select Case tQuery!Intervencion
            Case "A"
                If tQuery!Fecha_hasta > FechaURev Then
                    FechaURev = tQuery!Fecha_hasta
                    TipoRev = tQuery!Intervencion
                End If
                tQuery.MoveNext
            
            Case "AB"
                If tQuery!Fecha_hasta > FechaURev Then
                    FechaURev = tQuery!Fecha_hasta
                    TipoRev = tQuery!Intervencion
                End If
                tQuery.MoveNext
            
            Case "E"
                If tQuery!Fecha_hasta > FechaURev Then
                    FechaURev = tQuery!Fecha_hasta
                    TipoRev = tQuery!Intervencion
                End If
                tQuery.MoveNext
            Case Else
                tQuery.MoveNext
        End Select
    Wend

    lblUltimaRev2.ForeColor = vbBlack
    lblUltimaRev2.AutoSize = True
    lblUltimaRev2.FontBold = True
    lblUltimaRev2.FontSize = 10
    lblUltimaRev2.Caption = Format$(FechaURev, "DD/MM/YYYY")
    lblUltimaRev2.Visible = True
    
    lblTipo2.ForeColor = vbBlack
    lblTipo2.AutoSize = True
    lblTipo2.FontBold = True
    lblTipo2.FontSize = 10
    lblTipo2.Caption = TipoRev
    lblTipo2.Visible = True
    
    vSQL2 = "SELECT * FROM KILOMETRAJE WHERE Coche Like '" & cmbCCRR.text & "' AND  FECHA >=#" & Format$(FechaURev, "MM/DD/YYYY") & "#"
    
    'MsgBox (vSQL2)
    
    Set tQuery2 = db.OpenRecordset(vSQL2)
    
    If Not tQuery2.EOF Then
        tQuery2.MoveLast
        lblInfo.Caption = "Trabajando con " & Format(tQuery2.RecordCount, "#,##0") & " regitros de Kms de Periódicas"
     Else
    End If
    
    If Not tQuery2.EOF Then
        tQuery2.MoveFirst
      Else
        lblKmRev2.ForeColor = vbBlack
        lblKmRev2.AutoSize = True
        lblKmRev2.FontBold = True
        lblKmRev2.FontSize = 10
        lblKmRev2.Caption = Format$(TotalKm, "Standard")
        lblKmRev2.Visible = True
       Exit Sub
    End If
            
    While Not tQuery2.EOF
        TotalKm = TotalKm + tQuery2!Kms_Diario
        tQuery2.MoveNext
    Wend
    
    lblKmRev2.ForeColor = vbBlack
    lblKmRev2.AutoSize = True
    lblKmRev2.FontBold = True
    lblKmRev2.FontSize = 10
    lblKmRev2.Caption = Format$(TotalKm, "Standard")
    lblKmRev2.Visible = True
    
    tQuery.Close
    tQuery2.Close
    db1.Close
    
ErrorHandler:
   
    Select Case Err
        Case 3021
            MsgBox "No se encuentra Rev para el vehículo consultado", vbCritical
            Resume Next
        'Case Else
        '   MsgBox "Err" & " " & Err.Description, vbCritical
    End Select
   
'    Set db = Nothing
'    Set ws = Nothing

End Sub

Private Sub UltimaRG()
    Dim db1 As DAO.Database
    Dim ws As DAO.Workspace
    Dim rutaBase As String
    Dim clave As String
    Dim tQuery, tQuery2, vSQL, vSQL2
    Dim TotalKm As Double
    Dim FechaURG As Date
    
    FechaURG = "01/01/1900"
    TotalKm = 0
    
    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCRR.mdb"
    clave = "theidol-1995"

    ' Crear el workspace usando el motor Jet
    Set ws = DBEngine.Workspaces(0)

    On Error GoTo ErrorHandler

    ' Abrir la base de datos con contraseña
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    vSQL = "SELECT * FROM DETENCIONES WHERE Coche Like '" & cmbCCRR.text & "' AND  INTERVENCION='RG' ORDER BY FECHA_HASTA"
    
    'MsgBox (vSQL)
    
    Set tQuery = db.OpenRecordset(vSQL)
    
    tQuery.MoveFirst
            
    While Not tQuery.EOF
        If tQuery!Fecha_hasta > FechaURG Then FechaURG = tQuery!Fecha_hasta
        tQuery.MoveNext
    Wend

    If IsNumeric(Left(cmbCCRR.text, 1)) Then
        lblUltimaRG2.ForeColor = vbBlue
        lblUltimaRG2.AutoSize = True
        lblUltimaRG2.FontBold = True
        lblUltimaRG2.FontSize = 10
        lblUltimaRG2.Caption = Format$(FechaURG, "DD/MM/YYYY")
        lblUltimaRG2.Visible = True
        'MsgBox ("Es Materfer")
     Else
        'MsgBox ("Es CNR")
        lblUltimaRGCNR.ForeColor = vbBlue
        lblUltimaRGCNR.AutoSize = True
        lblUltimaRGCNR.FontBold = True
        lblUltimaRGCNR.FontSize = 10
        lblUltimaRGCNR.Caption = Format$(FechaURG, "DD/MM/YYYY")
        lblUltimaRGCNR.Visible = True
    End If
    
    
    vSQL2 = "SELECT * FROM KILOMETRAJE WHERE Coche Like '" & cmbCCRR.text & "' AND  FECHA >=#" & Format$(FechaURG, "MM/DD/YYYY") & "#"
    
    'MsgBox (vSQL2)
    
    Set tQuery2 = db.OpenRecordset(vSQL2)
    
    tQuery2.MoveFirst
            
    While Not tQuery2.EOF
        TotalKm = TotalKm + tQuery2!Kms_Diario
        tQuery2.MoveNext
    Wend
    
    If IsNumeric(Left(cmbCCRR.text, 1)) Then
        lblKmRG2.ForeColor = vbBlue
        lblKmRG2.AutoSize = True
        lblKmRG2.FontBold = True
        lblKmRG2.FontSize = 10
        lblKmRG2.Caption = Format$(TotalKm, "Standard")
        lblKmRG2.Visible = True
     Else
        lblKmRGCNR.ForeColor = vbBlue
        lblKmRGCNR.AutoSize = True
        lblKmRGCNR.FontBold = True
        lblKmRGCNR.FontSize = 10
        lblKmRGCNR.Caption = Format$(TotalKm, "Standard")
        lblKmRGCNR.Visible = True
    End If
    
    tQuery.Close
    tQuery2.Close
    db1.Close
    
ErrorHandler:
   
    Select Case Err
        Case 3021
            MsgBox "No se encuentra RG para el vehículo consultado", vbCritical
            Resume Next
        'Case Else
        '   MsgBox "Err" & " " & Err.Description, vbCritical
    End Select
   
'    Set db = Nothing
'    Set ws = Nothing
End Sub

Private Sub UltimaRP()

    Dim db1 As DAO.Database
    Dim ws As DAO.Workspace
    Dim rutaBase As String
    Dim clave As String
    Dim tQuery, tQuery2, vSQL, vSQL2
    Dim TotalKm As Double
    Dim FechaURP As Date
    
    FechaURP = "01/01/1900"
    TotalKm = 0
    
    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCRR.mdb"
    clave = "theidol-1995"

    ' Crear el workspace usando el motor Jet
    Set ws = DBEngine.Workspaces(0)

    On Error GoTo ErrorHandler

    ' Abrir la base de datos con contraseña
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    vSQL = "SELECT * FROM DETENCIONES WHERE Coche Like '" & cmbCCRR.text & "' AND  INTERVENCION='RP' ORDER BY FECHA_HASTA"
    
    'MsgBox (vSQL)
    
    Set tQuery = db.OpenRecordset(vSQL)
    
    tQuery.MoveFirst
            
    While Not tQuery.EOF
        If tQuery!Fecha_hasta > FechaURP Then FechaURP = tQuery!Fecha_hasta
        tQuery.MoveNext
    Wend

   ' lblUltimaRP2.ForeColor = vbGreen
    lblUltimaRP2.AutoSize = True
    lblUltimaRP2.FontBold = True
    lblUltimaRP2.FontSize = 10
    lblUltimaRP2.Caption = Format$(FechaURP, "DD/MM/YYYY")
    lblUltimaRP2.Visible = True
    
    vSQL2 = "SELECT * FROM KILOMETRAJE WHERE Coche Like '" & cmbCCRR.text & "' AND  FECHA >=#" & Format$(FechaURP, "MM/DD/YYYY") & "#"
    
    'MsgBox (vSQL2)
    
    Set tQuery2 = db.OpenRecordset(vSQL2)
    
    tQuery2.MoveFirst
            
    While Not tQuery2.EOF
        TotalKm = TotalKm + tQuery2!Kms_Diario
        tQuery2.MoveNext
    Wend
    
   ' lblKmRP2.ForeColor = vbCyan
    lblKmRP2.AutoSize = True
    lblKmRP2.FontBold = True
    lblKmRP2.FontSize = 10
    lblKmRP2.Caption = Format$(TotalKm, "Standard")
    lblKmRP2.Visible = True
    
    tQuery.Close
    tQuery2.Close
    db1.Close
    
ErrorHandler:
   
    Select Case Err
        Case 3021
            MsgBox "No se encuentra RP para el vehículo consultado", vbCritical
            Resume Next
        'Case Else
        '   MsgBox "Err" & " " & Err.Description, vbCritical
    End Select
   
'    Set db = Nothing
'    Set ws = Nothing

End Sub


Private Sub cmdCopiar_Click()

    Dim row As Long, col As Long
    Dim text As String

    For row = 0 To FG1.Rows - 1
        For col = 0 To FG1.Cols - 1
            text = text & FG1.TextMatrix(row, col)
            If col < FG1.Cols - 1 Then text = text & vbTab
        Next col
        text = text & vbCrLf
    Next row

    Clipboard.Clear
    Clipboard.SetText text
    MsgBox "Copiado al portapapeles. Pegalo en Excel con Ctrl+V.", vbInformation


End Sub


Private Sub CopyFlexGridToClipboard(fg As MSFlexGrid)
    Dim row As Long, col As Long
    Dim text As String

    For row = 0 To fg.Rows - 1
        For col = 0 To fg.Cols - 1
            text = text & fg.TextMatrix(row, col)
            If col < fg.Cols - 1 Then text = text & vbTab
        Next col
        text = text & vbCrLf
    Next row

    Clipboard.Clear
    Clipboard.SetText text
    MsgBox "Copiado al portapapeles. Pegalo en Excel con Ctrl+V.", vbInformation
End Sub


Private Sub cmdProyeccionABC_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Function ObtenerMaxFecha(ByVal coche As Variant, ByVal tipoIntervencion As String, ByRef db As DAO.Database) As Variant
' Esta función devuelve la fecha máxima para un coche y tipo de intervención específicos.
' Devuelve Null si no encuentra ninguna.
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT MAX(Fecha_hasta) AS MaxFecha FROM Detenciones WHERE Coche = '" & coche & "' AND Intervencion = '" & tipoIntervencion & "'"
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        ObtenerMaxFecha = rs!MaxFecha
    Else
        ObtenerMaxFecha = Null
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub FormatearCelda(ByVal fila As Long, ByVal col As Long, ByVal valorKm As Double, ByVal tipoIntervencion As String)
' Esta subrutina aplica O QUITA el formato condicional a una celda.

    Const ROJO_SUAVE As Long = &HE0E0FF
    Const AMARILLO_SUAVE As Long = &HE0FFFF
    Const CELESTE_CLARO As Long = &HFFF0E0
    
    ' Seleccionamos la celda que vamos a formatear
    FG1.row = fila
    FG1.col = col

    ' --- MEJORA: Primero, restauramos los colores por defecto ---
    ' Esto asegura que si una celda deja de cumplir la condición, se limpie.
    FG1.cellBackColor = vbWhite ' O el color de fondo que prefieras
    FG1.cellForeColor = vbBlack

    ' Ahora, aplicamos la lógica de formato condicional solo si se cumple
    Select Case tipoIntervencion
        Case "RG"
            If valorKm > 480000 Then
                FG1.cellBackColor = ROJO_SUAVE
                FG1.cellForeColor = vbRed
            End If
        Case "RP"
            If valorKm > 240000 Then
                FG1.cellBackColor = AMARILLO_SUAVE
                FG1.cellForeColor = &H808080   ' Amarillo oscuro
            End If
        Case "ABC"
            If valorKm > 120000 Then
                FG1.cellBackColor = CELESTE_CLARO
                FG1.cellForeColor = vbBlue
            End If
    End Select
End Sub
Private Sub cmdProyeccionABC_Click()
    
    CargarCochesFijos
    
    ' --- (Validaciones iniciales - sin cambios) ---
    If Not IsNumeric(txtKmsPromedioMes.text) Or Val(txtKmsPromedioMes.text) <= 0 Then
        MsgBox "Por favor, ingrese un valor numérico válido para los Kms Promedio por Mes.", vbExclamation
        txtKmsPromedioMes.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtProyeccionMeses.text) Or Val(txtProyeccionMeses.text) <= 0 Then
        MsgBox "Por favor, ingrese una cantidad de meses válida para la proyección.", vbExclamation
        txtProyeccionMeses.SetFocus
        Exit Sub
    End If
    If colCochesRemolcados Is Nothing Or colCochesRemolcados.Count = 0 Then
        MsgBox "No hay coches cargados para procesar.", vbExclamation
        Exit Sub
    End If

    ' --- (Declaración de variables - sin cambios) ---
    Dim db As DAO.Database
    Dim numCoche As Variant
    Dim rutaBase As String, clave As String
    Dim fechaRG As Variant, fechaRP As Variant, fechaABC As Variant
    Dim kmRG As Double, kmRP As Double, kmABC As Double
    Dim fechaParaSuma As Date
    Dim kmsPromedioMes As Double
    Dim mesesProyeccion As Integer
    Dim i As Integer
    Dim proyKmRG As Double, proyKmRP As Double, proyKmABC As Double
    Dim filaActual As Integer
    
    kmsPromedioMes = Val(txtKmsPromedioMes.text)
    mesesProyeccion = CInt(txtProyeccionMeses.text)
    filaActual = 1

    ' --- (Inicialización del Grid - sin cambios) ---
    With FG1
        .Clear
        .Cols = 4 + mesesProyeccion
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "N° Coche"
        .TextMatrix(0, 1) = "Intervención"
        .TextMatrix(0, 2) = "Fecha"
        .TextMatrix(0, 3) = "Km Acumulado"
        .ColWidth(0) = 900
        .ColWidth(1) = 1000
        .ColWidth(2) = 1100
        .ColWidth(3) = 1200
        For i = 1 To mesesProyeccion
            .TextMatrix(0, 3 + i) = Format(DateAdd("m", i, Date), "mmm yy")
            .ColWidth(3 + i) = 1100
        Next i
        .row = 0
        .RowSel = 0
        .col = 0
        .ColSel = .Cols - 1
        .cellFontBold = True
    End With

    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCRR.mdb"
    clave = "theidol-1995"

    On Error GoTo ErrorHandler
    Set db = OpenDatabase(rutaBase, False, False, "MS Access;PWD=" & clave)
    On Error GoTo 0

    For Each numCoche In colCochesRemolcados
    
        fechaRG = ObtenerMaxFecha(numCoche, "RG", db)
        fechaRP = ObtenerMaxFecha(numCoche, "RP", db)
        fechaABC = ObtenerMaxFecha(numCoche, "ABC", db)

        ' --- CORRECCIÓN: Lógica de cálculo de KM con sintaxis de bloque IF ---
        
        ' Lógica para RG
        If Not IsNull(fechaRG) Then
            kmRG = CalcularSumaKm(numCoche, fechaRG, db)
        Else
            kmRG = 0
        End If

        ' Lógica para RP
        If Not IsNull(fechaRP) Then
            If Not IsNull(fechaRG) And fechaRP < fechaRG Then
                kmRP = kmRG
            Else
                kmRP = CalcularSumaKm(numCoche, fechaRP, db)
            End If
        Else
            kmRP = 0
        End If

        ' Lógica para ABC
        If Not IsNull(fechaABC) Then
            If Not IsNull(fechaRG) And Not IsNull(fechaRP) Then
                fechaParaSuma = IIf(fechaRG > fechaRP, fechaRG, fechaRP)
            ElseIf Not IsNull(fechaRG) Then
                fechaParaSuma = fechaRG
            ElseIf Not IsNull(fechaRP) Then
                fechaParaSuma = fechaRP
            Else
                fechaParaSuma = #1/1/1900#
            End If
            
            If fechaABC < fechaParaSuma Then
                If fechaParaSuma = fechaRG Then
                    kmABC = kmRG
                Else
                    kmABC = kmRP
                End If
            Else
                kmABC = CalcularSumaKm(numCoche, fechaABC, db)
            End If
        Else
            kmABC = 0
        End If
        
        ' --- (El resto del código para llenar el grid y la proyección no cambia) ---
        FG1.Rows = FG1.Rows + 3
        
        proyKmRG = kmRG
        proyKmRP = kmRP
        proyKmABC = kmABC
        
        ' --- Fila RG ---
        FG1.TextMatrix(filaActual, 0) = numCoche
        FG1.TextMatrix(filaActual, 1) = "RG"
        FG1.TextMatrix(filaActual, 2) = IIf(IsNull(fechaRG), "N/A", Format(fechaRG, "dd/mm/yyyy"))
        FG1.TextMatrix(filaActual, 3) = Format(kmRG, "#,##0")
        FG1.row = filaActual: FG1.col = 0: FG1.cellFontBold = True
        FG1.col = 1: FG1.cellFontBold = False: FG1.cellForeColor = vbRed
        FG1.col = 2: FG1.cellForeColor = vbBlack
        FormatearCelda filaActual, 3, kmRG, "RG"

        ' --- Fila RP ---
        FG1.TextMatrix(filaActual + 1, 0) = numCoche
        FG1.TextMatrix(filaActual + 1, 1) = "RP"
        FG1.TextMatrix(filaActual + 1, 2) = IIf(IsNull(fechaRP), "N/A", Format(fechaRP, "dd/mm/yyyy"))
        FG1.TextMatrix(filaActual + 1, 3) = Format(kmRP, "#,##0")
        FG1.row = filaActual + 1: FG1.col = 0: FG1.cellFontBold = True
        FG1.col = 1: FG1.cellFontBold = False: FG1.cellForeColor = &H808080
        FG1.col = 2: FG1.cellForeColor = vbBlack
        FormatearCelda filaActual + 1, 3, kmRP, "RP"

        ' --- Fila ABC ---
        FG1.TextMatrix(filaActual + 2, 0) = numCoche
        FG1.TextMatrix(filaActual + 2, 1) = "ABC"
        FG1.TextMatrix(filaActual + 2, 2) = IIf(IsNull(fechaABC), "N/A", Format(fechaABC, "dd/mm/yyyy"))
        FG1.TextMatrix(filaActual + 2, 3) = Format(kmABC, "#,##0")
        FG1.row = filaActual + 2: FG1.col = 0: FG1.cellFontBold = True
        FG1.col = 1: FG1.cellFontBold = False: FG1.cellForeColor = vbBlue
        FG1.col = 2: FG1.cellForeColor = vbBlack
        FormatearCelda filaActual + 2, 3, kmABC, "ABC"
        
        For i = 1 To mesesProyeccion
            proyKmRG = proyKmRG + kmsPromedioMes
            proyKmRP = proyKmRP + kmsPromedioMes
            proyKmABC = proyKmABC + kmsPromedioMes
            
            FG1.TextMatrix(filaActual, 3 + i) = Format(proyKmRG, "#,##0")
            FormatearCelda filaActual, 3 + i, proyKmRG, "RG"
            
            FG1.TextMatrix(filaActual + 1, 3 + i) = Format(proyKmRP, "#,##0")
            FormatearCelda filaActual + 1, 3 + i, proyKmRP, "RP"
            
            FG1.TextMatrix(filaActual + 2, 3 + i) = Format(proyKmABC, "#,##0")
            FormatearCelda filaActual + 2, 3 + i, proyKmABC, "ABC"
        Next i
        
        filaActual = filaActual + 3
    Next numCoche

    MsgBox "Proceso y proyección finalizados."
    
GoTo Limpieza
ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
Limpieza:
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
End Sub

Private Sub FG1_DblClick()
' Se activa al hacer doble clic en el grid. Prepara la celda para edición.
    
    ' Solo permite editar las columnas de kilometraje (desde la 3 en adelante)
    If FG1.MouseCol <= 2 Then Exit Sub
    
    ' Mueve el TextBox de edición sobre la celda seleccionada
    With txtEditGrid
        .Move FG1.CellLeft + FG1.Left, FG1.CellTop + FG1.Top, FG1.CellWidth, FG1.CellHeight
        .text = FG1.text
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub txtEditGrid_KeyPress(KeyAscii As Integer)
' Se activa al presionar una tecla en el TextBox de edición.
    
    ' Si se presiona Enter (código 13)
    If KeyAscii = 13 Then
        KeyAscii = 0 ' Evita el "bip" del sistema
        txtEditGrid.Visible = False ' Oculta el TextBox
        
        ' Llama al evento LostFocus para procesar el cambio
        FG1.SetFocus
    End If
    
    ' Si se presiona Escape (código 27)
    If KeyAscii = 27 Then
        KeyAscii = 0
        txtEditGrid.Visible = False
        FG1.SetFocus
    End If
End Sub

Private Sub Form_Load()
   
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim vSQL As String
    Dim rutaBase As String
    Dim clave As String
    Dim tKms, tCCRR
    Dim i As Integer

    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCRR.mdb"
    clave = "theidol-1995"

    ' Crear el workspace usando el motor Jet
    Set ws = DBEngine.Workspaces(0)

    On Error GoTo ErrorHandler

    ' Abrir la base de datos con contraseña
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)

    'MsgBox "Base de datos abierta correctamente.", vbInformation

    ' Aquí podés trabajar con la base: db.TableDefs, db.Execute, etc.
        Set tKms = db.OpenRecordset("Kilometraje", dbOpenTable)
        Set tCCRR = db.OpenRecordset("Coches", dbOpenTable)
        
    'CargarCochesFijos
    'CargarCochesCNR
    'CargoComboMes
        
    db.Close
    Set db = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error en la ejecución del programa: " & Err.Description, vbCritical
    Set db = Nothing
    Set ws = Nothing

End Sub
Private Function CalcularSumaKm(ByVal coche As Variant, ByVal fechaInicio As Date, ByRef db As DAO.Database) As Double
' Esta función calcula la suma de kilómetros para un coche desde una fecha de inicio.
' Devuelve 0 si no hay kilómetros para sumar.
    Dim rs As DAO.Recordset
    Dim sql As String
    
    ' Asegúrate que los nombres de campo (Kilometros, Coche, Fecha) son correctos para tu tabla Kilometraje
    ' --- CORRECCIÓN: Cambiado Kms_diario a Kilometros para que coincida con el código anterior ---
    sql = "SELECT SUM(Kms_diario) AS SumaTotal FROM Kilometraje WHERE Coche = '" & coche & "' AND Fecha >= #" & Format(fechaInicio, "mm/dd/yyyy") & "#"
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        ' --- FIX: Reemplazo de Nz() por un bloque If/Else ---
        ' Si el resultado de la suma es Nulo (porque no había registros), devolvemos 0.
        ' Si no, devolvemos el valor de la suma.
        If IsNull(rs!SumaTotal) Then
            CalcularSumaKm = 0
        Else
            CalcularSumaKm = rs!SumaTotal
        End If
    Else
        CalcularSumaKm = 0
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub txtKmsPromedioMes_GotFocus()

    txtKmsPromedioMes.SelLength = Len(txtKmsPromedioMes.text)

End Sub


Private Sub txtKmsPromedioMes_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub txtKmsPromedioTemporada_GotFocus()

    txtKmsPromedioTemporada.SelLength = Len(txtKmsPromedioTemporada.text)

End Sub


Private Sub txtKmsPromedioTemporada_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub txtProyeccionMeses_GotFocus()

    txtProyeccionMeses.SelLength = Len(txtProyeccionMeses.text)

End Sub


Private Sub txtProyeccionMeses_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


