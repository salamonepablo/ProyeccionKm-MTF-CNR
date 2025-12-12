VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormConsultaIntCCRR 
   Caption         =   "Intervenciones CCRR"
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
      Cols            =   5
      FixedCols       =   0
   End
   Begin VB.CommandButton cmdInt 
      Caption         =   "Ver &Intervenciones"
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
   Begin VB.ComboBox cmbCoche 
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
      Caption         =   "Coche:"
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
Attribute VB_Name = "FormConsultaIntCCRR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCoche_KeyPress(KeyAscii As Integer)

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



Private Sub cmdInt_Click()

    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rutaBase As String
    Dim clave As String
    Dim tQuery, vSQL
    Dim TotalInts As Double
    
    
    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCRR.mdb"
    clave = "theidol-1995"

    ' Crear el workspace usando el motor Jet
    Set ws = DBEngine.Workspaces(0)

 '   On Error GoTo ErrorHandler

    ' Abrir la base de datos con contraseña
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    vSQL = "SELECT * FROM Detenciones WHERE Coche Like '" & cmbCoche.text & "' AND  Fecha_desde >=#" & Format$(txtFechaDesde.text, "MM/DD/YYYY") & "# AND Fecha_hasta <=#" & Format$(txtFechaHasta.text, "MM/DD/YYYY") & "# ORDER BY Coche, Fecha_hasta"
    
    'MsgBox (vSQL)
    
    Set tQuery = db.OpenRecordset(vSQL)
    
    If Not tQuery.EOF Then
        tQuery.MoveFirst
     Else
        MsgBox "No se encuentran datos en el rango de fechas especificado", vbCritical
    End If
    
    FG1.Clear
    FG1.Rows = 2
    
    FG1.row = 0
    
    FG1.col = 0
    FG1.ColWidth(0) = 1100
    FG1.CellFontBold = True
    FG1.text = "Coche"
    
    FG1.col = 1
    FG1.CellFontBold = True
    FG1.text = "Fecha Ini"
    
    FG1.col = 2
    FG1.CellFontBold = True
    FG1.text = "Fecha Fin"
    
    FG1.col = 3
    FG1.ColWidth(3) = 1200
    FG1.CellFontBold = True
    FG1.text = "Intervención"
    
    FG1.col = 4
    FG1.CellFontBold = True
    FG1.text = "Lugar"
    
    TotalInts = 0
        
    While Not tQuery.EOF
        FG1.row = FG1.row + 1
        
        FG1.col = 0
        FG1.text = tQuery!coche
        
        FG1.col = 1
        FG1.text = tQuery!Fecha_desde
        
        FG1.col = 2
        FG1.text = tQuery!Fecha_hasta
        
        FG1.col = 3
        FG1.text = tQuery!Intervencion
        
        FG1.col = 4
        FG1.text = tQuery!Lugar
                
        TotalInts = TotalInts + 1
        
        FG1.Rows = FG1.Rows + 1
        tQuery.MoveNext
    Wend

    lblTotalKm.Caption = "Total Intervenciones en el Período: " & Format(TotalInts, "#,###,###,#0")

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
    Dim tKms, tCCRR, tInts

    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCRR.mdb"
    clave = "theidol-1995"

    ' Crear el workspace usando el motor Jet
    Set ws = DBEngine.Workspaces(0)

    On Error GoTo ErrorHandler

    ' Abrir la base de datos con contraseña
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    'MsgBox "Base de datos abierta correctamente.", vbInformation

    ' Aquí podés trabajar con la base: db.TableDefs, db.Execute, etc.
        'Set tKms = db.OpenRecordset("Kilometraje", dbOpenTable)
        Set tCCRR = db.OpenRecordset("Coches", dbOpenTable)
        
        tCCRR.MoveFirst
        
        While Not tCCRR.EOF
            cmbCoche.AddItem (tCCRR!coche)
            tCCRR.MoveNext
        Wend
        
        cmbCoche.AddItem ("*")
        
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


