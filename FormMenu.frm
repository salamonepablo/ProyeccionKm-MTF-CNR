VERSION 5.00
Begin VB.Form FormMenu 
   Caption         =   "Ver Info BBDD"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11625
   Icon            =   "FormMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Menú"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin VB.Frame Frame5 
         Caption         =   "Observaciones"
         Height          =   2895
         Left            =   5880
         TabIndex        =   11
         Top             =   2400
         Width           =   4815
         Begin VB.TextBox Text1 
            Height          =   2295
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   12
            Text            =   "FormMenu.frx":10CA
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Coches Eléctricos"
         Height          =   1695
         Left            =   480
         TabIndex        =   8
         Top             =   3600
         Width           =   4935
         Begin VB.CommandButton cmdKmCCEE 
            Caption         =   "Kilometrajes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton cmdIntCCEE 
            Caption         =   "Intervenciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   2640
            TabIndex        =   9
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Locomotoras"
         Height          =   1695
         Left            =   5880
         TabIndex        =   5
         Top             =   480
         Width           =   4935
         Begin VB.CommandButton cmdKmLocs 
            Caption         =   "Kilometrajes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton cmdIntLocs 
            Caption         =   "Intervenciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   2640
            TabIndex        =   6
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Coches Remolcados"
         Height          =   2775
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   4935
         Begin VB.CommandButton cmdCCRR 
            Caption         =   "Km + Intervenciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton cmdIntCCRR 
            Caption         =   "Intervenciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   4
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CommandButton cmdProyCR 
            Caption         =   "Proyecciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   2640
            TabIndex        =   3
            Top             =   360
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCCRR_Click()

    FormConsultaKmCCRR.Show
    
End Sub


Private Sub cmdCCRR_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub cmdIntCCEE_Click()

    FormConsultaIntCCEE.Show

End Sub

Private Sub cmdIntCCEE_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub cmdIntCCRR_Click()

    FormConsultaIntCCRR.Show

End Sub

Private Sub cmdIntCCRR_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub cmdIntLocs_Click()

    FormConsultaIntLocs.Show

End Sub

Private Sub cmdIntLocs_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub cmdKmCCEE_Click()
    
    FormConsultaKmCCEE.Show

End Sub

Private Sub cmdKmCCEE_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub cmdKmLocs_Click()

    FormConsultaKmLocs.Show

End Sub

Private Sub cmdKmLocs_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub cmdProyCR_Click()

    FormConsKmsCPG.Show

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    'Text1.MultiLine = True
    Text1.text = "Hacer proyeccion CNR " & Chr(13) & "Proyeccion CSR "
    
End Sub


