VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormConsultaKmCCEE 
   Caption         =   "Km Recorridos CCEE"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKmsUANBARG 
      Caption         =   "Ver Km AN / BA / RG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   10
      Top             =   720
      Width           =   2175
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
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   5280
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid FG1 
      Height          =   3495
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
   Begin VB.CommandButton cmdKaims 
      Caption         =   "Ver &Km"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.ComboBox cmbCCEE 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtFechaDesde 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtFechaHasta 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblTotalKm 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Kms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   5400
      Width           =   6855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Coche:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha Hasta:"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha Desde:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FormConsultaKmCCEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private colFechasInicio As Collection ' Colección para guardar las fechas de inicio
' Función para obtener la fecha MÁS ANTIGUA de kilometraje de un coche.
Private Function ObtenerMinFechaKm(ByVal coche As String, ByRef db As DAO.Database) As Variant
    Dim rs As DAO.Recordset
    Dim sql As String
    
    ' Usamos MIN(Fecha) para obtener la fecha más baja (la más antigua)
    sql = "SELECT MIN(Fecha) AS MinFecha FROM Kilometraje WHERE Coche = '" & coche & "'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        ObtenerMinFechaKm = rs!MinFecha ' Devuelve la fecha o Null si no hay registros
    Else
        ObtenerMinFechaKm = Null
    End If
    
    rs.Close
    Set rs = Nothing
End Function
' Nueva función para cargar las fechas de inicio directamente en el código.
Private Sub CargarFechasInicioHardcoded()
    On Error GoTo ErrorHandler
    
    ' Inicializamos la colección
    Set colFechasInicio = New Collection
    
    ' Agregamos cada fecha a la colección, usando el módulo como clave.
    colFechasInicio.Add CDate("30/09/2017"), "1"
    colFechasInicio.Add CDate("31/08/2015"), "2"
    colFechasInicio.Add CDate("31/10/2015"), "3"
    colFechasInicio.Add CDate("30/09/2015"), "4"
    colFechasInicio.Add CDate("31/08/2016"), "5"
    colFechasInicio.Add CDate("31/08/2015"), "6"
    colFechasInicio.Add CDate("31/08/2015"), "7"
    colFechasInicio.Add CDate("31/08/2016"), "8"
    colFechasInicio.Add CDate("31/07/2016"), "9"
    colFechasInicio.Add CDate("31/07/2016"), "10"
    colFechasInicio.Add CDate("30/09/2015"), "11"
    colFechasInicio.Add CDate("31/10/2015"), "12"
    colFechasInicio.Add CDate("31/03/2016"), "13"
    colFechasInicio.Add CDate("30/09/2015"), "14"
    colFechasInicio.Add CDate("31/01/2016"), "15"
    colFechasInicio.Add CDate("30/09/2015"), "16"
    colFechasInicio.Add CDate("30/09/2015"), "17"
    colFechasInicio.Add CDate("31/03/2016"), "18"
    colFechasInicio.Add CDate("30/09/2015"), "19"
    colFechasInicio.Add CDate("31/01/2016"), "20"
    colFechasInicio.Add CDate("31/10/2016"), "21"
    colFechasInicio.Add CDate("30/06/2016"), "22"
    colFechasInicio.Add CDate("31/10/2015"), "23"
    colFechasInicio.Add CDate("30/09/2015"), "24"
    colFechasInicio.Add CDate("30/09/2015"), "25"
    colFechasInicio.Add CDate("30/09/2015"), "26"
    colFechasInicio.Add CDate("30/09/2015"), "27"
    colFechasInicio.Add CDate("30/09/2015"), "28"
    colFechasInicio.Add CDate("31/01/2016"), "29"
    colFechasInicio.Add CDate("30/09/2015"), "30"
    colFechasInicio.Add CDate("30/06/2017"), "31"
    colFechasInicio.Add CDate("31/01/2016"), "32"
    colFechasInicio.Add CDate("30/09/2015"), "33"
    colFechasInicio.Add CDate("30/09/2015"), "34"
    colFechasInicio.Add CDate("28/12/2019"), "35"
    colFechasInicio.Add CDate("30/09/2015"), "36"
    colFechasInicio.Add CDate("30/09/2015"), "37"
    colFechasInicio.Add CDate("31/08/2016"), "38"
    colFechasInicio.Add CDate("30/09/2015"), "39"
    colFechasInicio.Add CDate("30/09/2015"), "40"
    colFechasInicio.Add CDate("29/02/2016"), "41"
    colFechasInicio.Add CDate("29/02/2016"), "42"
    colFechasInicio.Add CDate("31/01/2016"), "43"
    colFechasInicio.Add CDate("31/01/2016"), "44"
    colFechasInicio.Add CDate("31/08/2017"), "45"
    colFechasInicio.Add CDate("31/01/2016"), "46"
    colFechasInicio.Add CDate("31/01/2016"), "48"
    colFechasInicio.Add CDate("30/06/2017"), "49"
    colFechasInicio.Add CDate("31/07/2017"), "50"
    colFechasInicio.Add CDate("30/11/2016"), "51"
    colFechasInicio.Add CDate("31/08/2016"), "52"
    colFechasInicio.Add CDate("30/11/2016"), "53"
    colFechasInicio.Add CDate("31/03/2016"), "54"
    colFechasInicio.Add CDate("31/08/2017"), "55"
    colFechasInicio.Add CDate("23/12/2019"), "56"
    colFechasInicio.Add CDate("30/09/2017"), "57"
    colFechasInicio.Add CDate("31/08/2016"), "58"
    colFechasInicio.Add CDate("31/03/2016"), "59"
    colFechasInicio.Add CDate("31/12/2016"), "60"
    colFechasInicio.Add CDate("31/01/2016"), "61"
    colFechasInicio.Add CDate("31/01/2016"), "62"
    colFechasInicio.Add CDate("31/01/2016"), "63"
    colFechasInicio.Add CDate("31/01/2016"), "64"
    colFechasInicio.Add CDate("31/01/2016"), "65"
    colFechasInicio.Add CDate("31/01/2017"), "66"
    colFechasInicio.Add CDate("28/02/2017"), "68"
    colFechasInicio.Add CDate("31/08/2016"), "69"
    colFechasInicio.Add CDate("30/09/2017"), "70"
    colFechasInicio.Add CDate("30/06/2017"), "71"
    colFechasInicio.Add CDate("30/11/2016"), "72"
    colFechasInicio.Add CDate("30/09/2017"), "73"
    colFechasInicio.Add CDate("31/01/2016"), "74"
    colFechasInicio.Add CDate("31/01/2016"), "75"
    colFechasInicio.Add CDate("29/06/2018"), "76"
    colFechasInicio.Add CDate("30/06/2016"), "77"
    colFechasInicio.Add CDate("31/01/2017"), "78"
    colFechasInicio.Add CDate("30/04/2016"), "79"
    colFechasInicio.Add CDate("30/04/2016"), "80"
    colFechasInicio.Add CDate("31/01/2016"), "81"
    colFechasInicio.Add CDate("31/12/2016"), "82"
    colFechasInicio.Add CDate("28/02/2017"), "83"
    colFechasInicio.Add CDate("31/08/2017"), "84"
    colFechasInicio.Add CDate("30/11/2016"), "85"
    colFechasInicio.Add CDate("30/06/2017"), "86"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ocurrió un error al cargar las fechas de inicio en memoria.", vbCritical, "Error en Carga"
End Sub

Private Sub cmbCCEE_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

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

Private Sub cmdKaims_Click()
    
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rutaBase As String
    Dim clave As String
    Dim tQuery, vSQL
    Dim TotalKm As Double
    
    
    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCEE.mdb"
    clave = "theidol-1995"

    ' Crear el workspace usando el motor Jet
    Set ws = DBEngine.Workspaces(0)

    'On Error GoTo ErrorHandler

    ' Abrir la base de datos con contraseña
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    vSQL = "SELECT * FROM KILOMETRAJE WHERE Coche Like '" & cmbCCEE.text & "' AND  FECHA >=#" & Format$(txtFechaDesde.text, "MM/DD/YYYY") & "# AND FECHA <=#" & Format$(txtFechaHasta.text, "MM/DD/YYYY") & "# ORDER BY Coche, Fecha"
    
    'MsgBox (vSQL)
    
    Set tQuery = db.OpenRecordset(vSQL)
    
    tQuery.MoveFirst
    
    FG1.Clear
    FG1.Rows = 2
    
    FG1.row = 0
    
    FG1.col = 0
    FG1.CellFontBold = True
    FG1.text = "COCHE"
    
    FG1.col = 1
    FG1.CellFontBold = True
    FG1.text = "FECHA"
    
    FG1.col = 2
    FG1.CellFontBold = True
    FG1.text = "KM."
    
    TotalKm = 0
        
    While Not tQuery.EOF
        FG1.row = FG1.row + 1
        
        FG1.col = 0
        FG1.text = tQuery!coche
        
        FG1.col = 1
        FG1.text = tQuery!Fecha
                
        FG1.col = 2
        FG1.text = Format$(tQuery!Kms_Diario, "#,###,###,#0.00")
        
        TotalKm = TotalKm + tQuery!Kms_Diario
        
        FG1.Rows = FG1.Rows + 1
        tQuery.MoveNext
    Wend

    Select Case cmbCCEE.text
        Case "*"
            lblTotalKm.Caption = "Total Coche Km. Recorridos: " & Format(TotalKm, "#,###,###,#0.00")
        
        Case "46*"
            lblTotalKm.Caption = "Total Módulo Km. Toshiba Recorridos: " & Format(TotalKm, "#,###,###,#0.00")
        
        Case "56*"
            lblTotalKm.Caption = "Total Módulo Km. CSR Recorridos: " & Format(TotalKm, "#,###,###,#0.00")
    End Select

ErrorHandler:
   
    Select Case Err
        Case 3021
            MsgBox "No se encuentran datos en el rango de fechas especificado", vbCritical
            Resume Next
        'Case Else
        '   MsgBox "Err" & " " & Err.Description, vbCritical
    End Select
   
'    Set db = Nothing
'    Set ws = Nothing

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



Private Sub cmdKmsUANBARG_Click()
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rsModulos As DAO.Recordset
    Dim sql As String
    Dim rutaBase As String
    Dim clave As String
    
    Dim moduloActual As String
    Dim nroModulo As String
    Dim fechaAN As Variant, fechaBA As Variant, fechaRPE As Variant, fechaInicioKm As Variant
    Dim kmAN As Double, kmBA As Double, kmRPE As Double, kmTotal As Double
    Dim fechaHasta As Date
    Dim filaActual As Long
    
    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCEE.mdb"
    clave = "theidol-1995"
    
    On Error GoTo ErrorHandler
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    With FG1
        .Clear
        .Cols = 4
        .Rows = 2
        .FixedRows = 1
        .TextMatrix(0, 0) = "Módulo"
        .TextMatrix(0, 1) = "Km. Acumulados"
        .TextMatrix(0, 2) = "Fecha Interv."
        .TextMatrix(0, 3) = "Tipo"
        .ColWidth(0) = 800
        .ColWidth(1) = 1800
        .ColWidth(2) = 1500
        .ColWidth(3) = 800
        .ColAlignment(0) = 4
        .ColAlignment(1) = 7
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
    End With
    
    sql = "SELECT Coche FROM Coches WHERE Coche LIKE '56*' ORDER BY Coche"
    Set rsModulos = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If rsModulos.EOF Then
        MsgBox "No se encontraron coches eléctricos (módulos 56).", vbInformation
        GoTo Cierre
    End If
    
    fechaHasta = CDate(txtFechaHasta.text)
    filaActual = 1
    
    While Not rsModulos.EOF
        moduloActual = rsModulos!coche
        nroModulo = Mid(moduloActual, 3, 2)
        
        ' Búsqueda de fecha de inicio desde la colección
        fechaInicioKm = Null
        On Error Resume Next
        fechaInicioKm = colFechasInicio(nroModulo)
        On Error GoTo 0
        
        fechaAN = ObtenerMaxFecha(moduloActual, "AN", db)
        fechaBA = ObtenerMaxFecha(moduloActual, "RPG", db)
        fechaRPE = ObtenerMaxFecha(moduloActual, "RPE", db)
        
        ' --- Lógica de cálculo con sintaxis de IF corregida ---
        Dim fechaBaseParaRPE As Variant
        fechaBaseParaRPE = fechaRPE
        If Not IsNull(fechaBaseParaRPE) Then
            kmRPE = CalcularSumaKmDesde(moduloActual, fechaBaseParaRPE, fechaHasta, db)
        Else
            kmRPE = 0
        End If
        
        Dim fechaBaseParaRPG As Variant
        If Not IsNull(fechaRPE) And Not IsNull(fechaBA) Then
            fechaBaseParaRPG = IIf(CDate(fechaRPE) > CDate(fechaBA), fechaRPE, fechaBA)
        ElseIf Not IsNull(fechaRPE) Then
            fechaBaseParaRPG = fechaRPE
        Else
            fechaBaseParaRPG = fechaBA
        End If
        
        If Not IsNull(fechaBaseParaRPG) Then
            kmBA = CalcularSumaKmDesde(moduloActual, fechaBaseParaRPG, fechaHasta, db)
        Else
            kmBA = 0
        End If

        Dim fechaBaseParaAN As Variant
        If Not IsNull(fechaBaseParaRPG) And Not IsNull(fechaAN) Then
            fechaBaseParaAN = IIf(CDate(fechaBaseParaRPG) > CDate(fechaAN), fechaBaseParaRPG, fechaAN)
        ElseIf Not IsNull(fechaBaseParaRPG) Then
            fechaBaseParaAN = fechaBaseParaRPG
        Else
            fechaBaseParaAN = fechaAN
        End If
        
        If Not IsNull(fechaBaseParaAN) Then
            kmAN = CalcularSumaKmDesde(moduloActual, fechaBaseParaAN, fechaHasta, db)
        Else
            kmAN = 0
        End If
        
        kmTotal = CalcularKmTotal(moduloActual, db)
        
        ' Agregamos las filas necesarias
        FG1.Rows = FG1.Rows + 4
        
        ' Llenamos las celdas una por una para mayor claridad
        FG1.TextMatrix(filaActual, 0) = nroModulo
        FG1.TextMatrix(filaActual, 1) = Format(kmAN, "#,##0")
        FG1.TextMatrix(filaActual, 2) = IIf(IsNull(fechaBaseParaAN), "N/A", Format(fechaBaseParaAN, "dd/mm/yyyy"))
        FG1.TextMatrix(filaActual, 3) = "AN"
        
        FG1.TextMatrix(filaActual + 1, 0) = nroModulo
        FG1.TextMatrix(filaActual + 1, 1) = Format(kmBA, "#,##0")
        FG1.TextMatrix(filaActual + 1, 2) = IIf(IsNull(fechaBaseParaRPG), "N/A", Format(fechaBaseParaRPG, "dd/mm/yyyy"))
        FG1.TextMatrix(filaActual + 1, 3) = "BA"
        
        FG1.TextMatrix(filaActual + 2, 0) = nroModulo
        FG1.TextMatrix(filaActual + 2, 1) = Format(kmRPE, "#,##0")
        FG1.TextMatrix(filaActual + 2, 2) = IIf(IsNull(fechaBaseParaRPE), "N/A", Format(fechaBaseParaRPE, "dd/mm/yyyy"))
        FG1.TextMatrix(filaActual + 2, 3) = "PA"
        
        FG1.TextMatrix(filaActual + 3, 0) = nroModulo
        FG1.TextMatrix(filaActual + 3, 1) = Format(kmTotal, "#,##0")
        FG1.TextMatrix(filaActual + 3, 2) = IIf(IsNull(fechaInicioKm), "N/A", Format(fechaInicioKm, "dd/mm/yyyy"))
        FG1.TextMatrix(filaActual + 3, 3) = "0 Km"
        
        filaActual = filaActual + 4
        rsModulos.MoveNext
    Wend
    
    MsgBox "Proceso finalizado. Se han cargado " & rsModulos.RecordCount & " módulos.", vbInformation

Cierre:
    If Not rsModulos Is Nothing Then rsModulos.Close
    If Not db Is Nothing Then db.Close
    Set rsModulos = Nothing
    Set db = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, "Error en Proceso"
    Resume Cierre
End Sub
Private Sub Form_Load()
   
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rutaBase As String
    Dim clave As String
    Dim tKms, tCCRR
    
    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCEE.mdb"
    'clave = "theidol-1995"
    clave = ""

    ' Crear el workspace usando el motor Jet
    Set ws = DBEngine.Workspaces(0)

    On Error GoTo ErrorHandler

    ' Abrir la base de datos con contraseña
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    'MsgBox "Base de datos abierta correctamente.", vbInformation

    ' Aquí podés trabajar con la base: db.TableDefs, db.Execute, etc.
        Set tKms = db.OpenRecordset("Kilometraje", dbOpenTable)
        Set tCCEE = db.OpenRecordset("Coches", dbOpenTable)
        
        tCCEE.MoveFirst
        
        While Not tCCEE.EOF
            cmbCCEE.AddItem (tCCEE!coche)
            tCCEE.MoveNext
        Wend
        cmbCCEE.AddItem ("46*")
        cmbCCEE.AddItem ("56*")
        
        txtFechaDesde.text = Format((Date - 30), "DD/MM/YYYY")
        txtFechaHasta.text = Format((Date), "DD/MM/YYYY")

        Call CargarFechasInicioHardcoded
    
    db.Close
    Set db = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error al abrir la base de datos: " & Err.Description, vbCritical
    Set db = Nothing
    Set ws = Nothing

End Sub

Private Sub txtFechaDesde_GotFocus()

    txtFechaDesde.SelLength = Len(txtFechaDesde.text)

End Sub

Private Sub txtFechaDesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub txtFechaHasta_GotFocus()

    txtFechaHasta.SelLength = Len(txtFechaHasta.text)

End Sub

Private Sub txtFechaHasta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


' Función para obtener la fecha máxima de una intervención específica.
Private Function ObtenerMaxFecha(ByVal coche As String, ByVal tipoIntervencion As String, ByRef db As DAO.Database) As Variant
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT MAX(Fecha_hasta) AS MaxFecha FROM Detenciones WHERE Coche LIKE '" & coche & "' AND Intervencion = '" & tipoIntervencion & "'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot) ' Usamos dbOpenSnapshot para consultas de solo lectura
    
    If Not rs.EOF Then
        ObtenerMaxFecha = rs!MaxFecha ' Devuelve la fecha o Null si no hay registros
    Else
        ObtenerMaxFecha = Null
    End If
    
    rs.Close
    Set rs = Nothing
End Function

' Función para sumar Kms desde una fecha de inicio hasta una fecha de fin.
Private Function CalcularSumaKmDesde(ByVal coche As String, ByVal fechaInicio As Variant, ByVal fechaFin As Date, ByRef db As DAO.Database) As Double
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim suma As Double
    
    ' Si la fecha de inicio es nula, no hay Kms que sumar desde esa intervención.
    If IsNull(fechaInicio) Then
        CalcularSumaKmDesde = 0
        Exit Function
    End If
    
    sql = "SELECT SUM(Kms_Diario) AS SumaTotal FROM Kilometraje WHERE Coche = '" & coche & "' " & _
          "AND Fecha >= #" & Format(fechaInicio, "mm/dd/yyyy") & "# " & _
          "AND Fecha <= #" & Format(fechaFin, "mm/dd/yyyy") & "#"
          
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        ' Nz() es una función de Access SQL que convierte Null en 0.
        ' Aquí lo replicamos en VB6.
        If IsNull(rs!SumaTotal) Then
            suma = 0
        Else
            suma = rs!SumaTotal
        End If
    Else
        suma = 0
    End If
    
    rs.Close
    Set rs = Nothing
    CalcularSumaKmDesde = suma
End Function

' Función para calcular el KM total de un coche desde el inicio.
Private Function CalcularKmTotal(ByVal coche As String, ByRef db As DAO.Database) As Double
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim suma As Double
    
    sql = "SELECT SUM(Kms_Diario) AS SumaTotal FROM Kilometraje WHERE Coche = '" & coche & "'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        If IsNull(rs!SumaTotal) Then
            suma = 0
        Else
            suma = rs!SumaTotal
        End If
    Else
        suma = 0
    End If
    
    rs.Close
    Set rs = Nothing
    CalcularKmTotal = suma
End Function

