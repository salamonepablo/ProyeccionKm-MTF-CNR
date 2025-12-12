VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormConsultaKmLocs 
   Caption         =   "Km Recorridos Locs"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
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
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid FG1 
      Height          =   3495
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   6615
      _ExtentX        =   11668
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
   Begin VB.ComboBox cmbLocs 
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
   Begin VB.Label lblKmsRG 
      Alignment       =   1  'Right Justify
      Caption         =   "Kms URG: "
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
      Left            =   5880
      TabIndex        =   10
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label lblTotalKm 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Kaims"
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
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Locomotora:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Hasta:"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Desde:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FormConsultaKmLocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbLocs_KeyPress(KeyAscii As Integer)

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
    Dim tQuery, vSQL, tKmUR
    Dim TotalKm As Double
    
    
    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseLocs.mdb"
    clave = "theidol-1995"

    ' Crear el workspace usando el motor Jet
    Set ws = DBEngine.Workspaces(0)

    'On Error GoTo ErrorHandler

    ' Abrir la base de datos con contraseña
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    vSQL = "SELECT * FROM KILOMETRAJE WHERE Locs Like '" & cmbLocs.text & "' AND  FECHA >=#" & Format$(txtFechaDesde.text, "MM/DD/YYYY") & "# AND FECHA <=#" & Format$(txtFechaHasta.text, "MM/DD/YYYY") & "# ORDER BY Locs, Fecha"
    
    vSQL2 = "SELECT Kilometraje.Locs, Sum(Kilometraje.Kms_diario) AS SumaDeKms_diario FROM Kilometraje INNER JOIN Ultima_reparacion ON Kilometraje.Locs = Ultima_reparacion.Locs Where (((Kilometraje.Fecha) > [ultima_reparacion].[fecha_hasta])) GROUP BY Kilometraje.Locs;"
    
    'SELECT Detenciones.Locs, Detenciones.Fecha_hasta, Detenciones.Intervencion, Detenciones.Lugar, Intervenciones.Intervencion_descripcion, Lugares.Lugar_descripcion
    'FROM Lugares INNER JOIN (Intervenciones INNER JOIN (Fecha_ultima_entrada_reparacion INNER JOIN Detenciones ON (Detenciones.Locs = Fecha_ultima_entrada_reparacion.Locs) AND (Fecha_ultima_entrada_reparacion.MáxDeFecha_hasta = Detenciones.Fecha_hasta)) ON Intervenciones.Intervencion_tipo = Detenciones.Intervencion) ON Lugares.Lugar_codigo = Detenciones.Lugar
    'WHERE (((Detenciones.Intervencion)<>"AL" And (Detenciones.Intervencion)<>"DI"));

    'SELECT Detenciones.Locs, Max(Detenciones.Fecha_hasta) AS MáxDeFecha_hasta
    'From Detenciones
    'Where (((Detenciones.Intervencion) = "RG" Or (Detenciones.Intervencion) = "N1" Or (Detenciones.Intervencion) = "N2" Or (Detenciones.Intervencion) = "N3" Or (Detenciones.Intervencion) = "N4" Or (Detenciones.Intervencion) = "N5" Or (Detenciones.Intervencion) = "N6"))
    'GROUP BY Detenciones.Locs
    'Having (((Max(Detenciones.Fecha_hasta)) Is Not Null))
    'ORDER BY Detenciones.Locs;




    'MsgBox (vSQL)
    
    Set tQuery = db.OpenRecordset(vSQL)
    
    'MsgBox (vSQL2)
    Set tKmUR = db.OpenRecordset(vSQL2)
    
    tKmUR.MoveFirst
    lblKmsRG.Caption = "Km.URG = " & Format(tKmUR!SumaDeKms_diario, "#,###,###,#0.00")
    
    tQuery.MoveFirst
    
    FG1.Clear
    FG1.Rows = 2
    
    FG1.row = 0
    
    FG1.col = 0
    FG1.ColWidth(0) = 1100
    FG1.CellFontBold = True
    FG1.text = "LOC"
    
    FG1.col = 1
    FG1.CellFontBold = True
    FG1.text = "FECHA"
    
    FG1.col = 2
    FG1.ColWidth(2) = 2000
    FG1.CellFontBold = True
    FG1.text = "KM."
    
    TotalKm = 0
        
    While Not tQuery.EOF
        FG1.row = FG1.row + 1
        
        FG1.col = 0
        FG1.text = tQuery!Locs
        
        FG1.col = 1
        FG1.text = tQuery!Fecha
                
        FG1.col = 2
        FG1.text = Format$(tQuery!Kms_Diario, "#,###,###,#0.00")
        
        TotalKm = TotalKm + tQuery!Kms_Diario
        
        FG1.Rows = FG1.Rows + 1
        tQuery.MoveNext
    Wend

    lblTotalKm.Caption = "Total Km. Recorridos: " & Format(TotalKm, "#,###,###,#0.00")

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


Private Sub Form_Load()
   
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rutaBase As String
    Dim clave As String
    Dim tKms, tLocs

    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseLocs.mdb"
    clave = "theidol-1995"

    ' Crear el workspace usando el motor Jet
    Set ws = DBEngine.Workspaces(0)

    On Error GoTo ErrorHandler

    ' Abrir la base de datos con contraseña
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    'MsgBox "Base de datos abierta correctamente.", vbInformation

    ' Aquí podés trabajar con la base: db.TableDefs, db.Execute, etc.
        Set tKms = db.OpenRecordset("Kilometraje", dbOpenTable)
        Set tLocs = db.OpenRecordset("Locomotoras", dbOpenTable)
        
        tLocs.MoveFirst
        
        While Not tLocs.EOF
            cmbLocs.AddItem (tLocs!Locs)
            tLocs.MoveNext
        Wend
        
        cmbLocs.AddItem ("*")
        
        txtFechaDesde.text = Format((Date - 30), "DD/MM/YYYY")
        txtFechaHasta.text = Format((Date), "DD/MM/YYYY")

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


