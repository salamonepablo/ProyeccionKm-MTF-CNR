VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormConsultaKmCCRR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Km Recorridos CCRR"
   ClientHeight    =   10380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10380
   ScaleWidth      =   20310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame12 
      Height          =   1455
      Left            =   360
      TabIndex        =   68
      Top             =   8760
      Width           =   19695
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   720
         TabIndex        =   69
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Info"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   19080
         TabIndex        =   70
         Top             =   1080
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdCopiar2 
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
      Left            =   7440
      TabIndex        =   67
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Frame Frame11 
      Caption         =   "Ultrasonido"
      Height          =   975
      Left            =   14400
      TabIndex        =   60
      Top             =   6480
      Width           =   2655
      Begin VB.Label lblKmsUS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km. US:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1425
         TabIndex        =   64
         Top             =   600
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblFechaUS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ultimo US:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1140
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   62
         Top             =   645
         Width           =   390
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   61
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Última A3"
      Height          =   975
      Left            =   17400
      TabIndex        =   55
      Top             =   2760
      Width           =   2655
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   59
         Top             =   285
         Width           =   600
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   58
         Top             =   645
         Width           =   390
      End
      Begin VB.Label lblFechaCNR_A3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ultima A3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1185
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblKmCNR_A3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km. A3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1470
         TabIndex        =   56
         Top             =   600
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Última Revisión"
      Height          =   1455
      Left            =   17400
      TabIndex        =   48
      Top             =   6000
      Width           =   2655
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   54
         Top             =   285
         Width           =   600
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   53
         Top             =   1005
         Width           =   390
      End
      Begin VB.Label lblFechaCNR_Rev 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ultima Rev:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1035
         TabIndex        =   52
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblKmCNR_Rev 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km. Rev:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   51
         Top             =   960
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblTipoCNR_Rev 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Rev:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   1230
         TabIndex        =   50
         Top             =   600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   420
         TabIndex        =   49
         Top             =   645
         Width           =   450
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Última A2"
      Height          =   975
      Left            =   17400
      TabIndex        =   43
      Top             =   3840
      Width           =   2655
      Begin VB.Label lblKmCNR_A2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km. A2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1470
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblFechaCNR_A2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ultima A2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1185
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   45
         Top             =   645
         Width           =   390
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   44
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Última A1"
      Height          =   975
      Left            =   17400
      TabIndex        =   38
      Top             =   4920
      Width           =   2655
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   42
         Top             =   285
         Width           =   600
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   41
         Top             =   645
         Width           =   390
      End
      Begin VB.Label lblFechaCNR_A1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ultima A1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   1185
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblKmCNR_A1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km.A1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   1545
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Última RG"
      Height          =   975
      Left            =   17400
      TabIndex        =   33
      Top             =   1680
      Width           =   2655
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   37
         Top             =   645
         Width           =   390
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   36
         Top             =   285
         Width           =   600
      End
      Begin VB.Label lblKmCNR_RG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km. RG:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1455
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblFechaCNR_RG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ultima RG:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1170
         TabIndex        =   34
         Top             =   255
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1455
      Left            =   240
      TabIndex        =   29
      Top             =   0
      Width           =   19815
      Begin VB.TextBox txtFechaHasta 
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
         Left            =   5760
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtFechaDesde 
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
         Left            =   3000
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbCCRR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8040
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdKaims 
         Caption         =   "Ver &Info de Coche"
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
         Left            =   10320
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   1215
         Left            =   120
         Picture         =   "AbrirBaseCCRR.frx":0000
         Stretch         =   -1  'True
         Top             =   160
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   945
         Left            =   17520
         Picture         =   "AbrirBaseCCRR.frx":3D05
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1845
      End
      Begin VB.Image Image1 
         Height          =   945
         Left            =   14640
         Picture         =   "AbrirBaseCCRR.frx":2A2CE
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1845
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         Height          =   195
         Left            =   1890
         TabIndex        =   32
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Hasta:"
         Height          =   255
         Left            =   4440
         TabIndex        =   31
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Coche:"
         Height          =   195
         Left            =   7425
         TabIndex        =   30
         Top             =   600
         Width           =   510
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Última Revisión"
      Height          =   1455
      Left            =   14400
      TabIndex        =   22
      Top             =   4920
      Width           =   2655
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   28
         Top             =   645
         Width           =   450
      End
      Begin VB.Label lblTipo2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Rev:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1230
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblKmRev2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km. Rev:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblUltimaRev2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ultima Rev:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1035
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   24
         Top             =   1005
         Width           =   390
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   23
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Última ABC"
      Height          =   975
      Left            =   14400
      TabIndex        =   17
      Top             =   3840
      Width           =   2655
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   21
         Top             =   285
         Width           =   600
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   20
         Top             =   645
         Width           =   390
      End
      Begin VB.Label lblUltimaABC2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ultima ABC:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1005
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblKmABC2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km. ABC:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1290
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Última RP"
      Height          =   975
      Left            =   14400
      TabIndex        =   10
      Top             =   2760
      Width           =   2655
      Begin VB.Label lblKmRP2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km. RP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   1425
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblUltimaRP2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ultima RP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   1140
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   12
         Top             =   645
         Width           =   390
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   11
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Última RG"
      Height          =   975
      Left            =   14400
      TabIndex        =   7
      Top             =   1680
      Width           =   2655
      Begin VB.Label lblUltimaRG2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ultima RG:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1170
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblKmRG2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km. RG:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1455
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblUltimaRG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   9
         Top             =   285
         Width           =   600
      End
      Begin VB.Label lblKmRG 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Km.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   8
         Top             =   645
         Width           =   390
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
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   7680
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid FG1 
      Height          =   5775
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid FG2 
      Height          =   5775
      Left            =   7200
      TabIndex        =   65
      Top             =   1680
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
   End
   Begin VB.Label lblTotalIntervenciones 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Ints"
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
      Left            =   8880
      TabIndex        =   66
      Top             =   7680
      Visible         =   0   'False
      Width           =   4575
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
      TabIndex        =   5
      Top             =   7680
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "FormConsultaKmCCRR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TInfoIntervencion
    ' La fecha debe ser Variant para poder manejar valores Null
    Fecha As Variant
    Tipo As String
    Encontrado As Boolean
End Type
Private Function ObtenerUltimaIntervencion(ByVal coche As String, ByVal tiposIntervencion As String, ByRef db As DAO.Database) As TInfoIntervencion
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim resultado As TInfoIntervencion
    
    sql = "SELECT TOP 1 Fecha_hasta, Intervencion FROM Detenciones " & _
          "WHERE Coche = '" & coche & "' AND Intervencion IN (" & tiposIntervencion & ") " & _
          "ORDER BY Fecha_hasta DESC"
          
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        ' Usamos IsNull para asignar correctamente el valor de la fecha
        If IsNull(rs!Fecha_hasta) Then
            resultado.Fecha = Null
        Else
            resultado.Fecha = rs!Fecha_hasta
        End If
        resultado.Tipo = rs!Intervencion
        resultado.Encontrado = True
    Else
        resultado.Fecha = Null
        resultado.Tipo = ""
        resultado.Encontrado = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    ObtenerUltimaIntervencion = resultado
End Function



Private Function CalcularSumaKm(ByVal coche As Variant, ByVal fechaInicio As Date, ByRef db As DAO.Database) As Double
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT SUM(Kms_Diario) AS SumaTotal FROM Kilometraje WHERE Coche = '" & coche & "' AND Fecha >= #" & Format(fechaInicio, "mm/dd/yyyy") & "#"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
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
Private Function ObtenerMaxFecha(ByVal coche As Variant, ByVal tipoIntervencion As String, ByRef db As DAO.Database) As Variant
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT MAX(Fecha_hasta) AS MaxFecha FROM Detenciones WHERE Coche = '" & coche & "' AND Intervencion = '" & tipoIntervencion & "'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        ' Devolvemos directamente el resultado de la consulta (puede ser Null)
        ObtenerMaxFecha = rs!MaxFecha
    Else
        ObtenerMaxFecha = Null
    End If
    
    rs.Close
    Set rs = Nothing
End Function


Private Sub UltimaABC(ByRef db As DAO.Database, ByVal fechaBaseRG As Variant, ByVal kmDesdeRG As Double, ByVal fechaEfectivaRP As Variant, ByVal kmEfectivosRP As Double)
    Dim FechaUABC As Variant
    Dim TotalKm As Double
    Dim fechaBaseEfectiva As Variant
    Dim kmDesdeBaseEfectiva As Double
    Dim tipoBase As String
    
    If Not IsNull(fechaEfectivaRP) Then
        fechaBaseEfectiva = fechaEfectivaRP
        kmDesdeBaseEfectiva = kmEfectivosRP
        tipoBase = "RP"
    Else
        fechaBaseEfectiva = fechaBaseRG
        kmDesdeBaseEfectiva = kmDesdeRG
        tipoBase = "RG"
    End If
    
    lblUltimaABC2.ForeColor = vbBlue
    lblKmABC2.ForeColor = vbBlue
    lblUltimaABC2.Visible = True
    lblKmABC2.Visible = True
    
    FechaUABC = ObtenerMaxFecha(cmbCCRR.text, "ABC", db)
    
    ' --- LÓGICA DE COMPARACIÓN SEGURA ---
    If Not IsNull(FechaUABC) And Not IsNull(fechaBaseEfectiva) Then
        If CDate(FechaUABC) < CDate(fechaBaseEfectiva) Then
            lblUltimaABC2.Caption = "Sin ABC"
            lblKmABC2.Caption = Format(kmDesdeBaseEfectiva, "Standard")
        Else
            TotalKm = CalcularSumaKm(cmbCCRR.text, CDate(FechaUABC), db)
            lblUltimaABC2.Caption = Format(FechaUABC, "DD/MM/YYYY")
            lblKmABC2.Caption = Format(TotalKm, "Standard")
        End If
    ElseIf IsNull(FechaUABC) Then
        lblUltimaABC2.Caption = "Sin ABC"
        lblKmABC2.Caption = Format(kmDesdeBaseEfectiva, "Standard")
    Else ' Solo existe FechaUABC
        TotalKm = CalcularSumaKm(cmbCCRR.text, CDate(FechaUABC), db)
        lblUltimaABC2.Caption = Format(FechaUABC, "DD/MM/YYYY")
        lblKmABC2.Caption = Format(TotalKm, "Standard")
    End If
End Sub
Private Sub UltimaUltrasonido(ByRef db As DAO.Database)
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim fechaUS As Variant
    Dim kmUS As Double
    
    sql = "SELECT MAX(Fecha) AS MaxFecha FROM Ultrasonido WHERE Coche = '" & cmbCCRR.text & "'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        fechaUS = rs!MaxFecha
    End If
    rs.Close
    Set rs = Nothing
    
    lblFechaUS.Visible = True
    lblKmsUS.Visible = True
    lblFechaUS.ForeColor = RGB(255, 140, 0)
    lblKmsUS.ForeColor = RGB(255, 140, 0)
    
    If IsNull(fechaUS) Then
        lblFechaUS.Caption = "Sin Datos"
        lblKmsUS.Caption = "0.00"
    Else
        kmUS = CalcularSumaKm(cmbCCRR.text, CDate(fechaUS), db)
        lblFechaUS.Caption = Format(fechaUS, "DD/MM/YYYY")
        lblKmsUS.Caption = Format(kmUS, "Standard")
    End If
End Sub
Private Sub UltimaRev(ByRef db As DAO.Database, ByVal fechaBaseRG As Variant, ByVal kmDesdeRG As Double)
    Dim ultimaRevInfo As TInfoIntervencion
    Dim TotalKm As Double
    
    lblUltimaRev2.Visible = True
    lblTipo2.Visible = True
    lblKmRev2.Visible = True
    
    lblUltimaRev2.ForeColor = vbBlack
    lblTipo2.ForeColor = vbBlack
    lblKmRev2.ForeColor = vbBlack
    
    ultimaRevInfo = ObtenerUltimaIntervencion(cmbCCRR.text, "'A', 'AB'", db)
    
    If ultimaRevInfo.Encontrado And Not IsNull(fechaBaseRG) Then
        If ultimaRevInfo.Fecha < CDate(fechaBaseRG) Then
            lblUltimaRev2.Caption = "Sin Rev"
            lblTipo2.Caption = "N/A"
            lblKmRev2.Caption = Format(kmDesdeRG, "Standard")
        Else
            TotalKm = CalcularSumaKm(cmbCCRR.text, ultimaRevInfo.Fecha, db)
            lblUltimaRev2.Caption = Format(ultimaRevInfo.Fecha, "DD/MM/YYYY")
            lblTipo2.Caption = ultimaRevInfo.Tipo
            lblKmRev2.Caption = Format(TotalKm, "Standard")
        End If
    ElseIf Not ultimaRevInfo.Encontrado Then
        lblUltimaRev2.Caption = "Sin Rev"
        lblTipo2.Caption = "N/A"
        lblKmRev2.Caption = Format(kmDesdeRG, "Standard")
    Else ' Solo existe la revision A o AB
        TotalKm = CalcularSumaKm(cmbCCRR.text, ultimaRevInfo.Fecha, db)
        lblUltimaRev2.Caption = Format(ultimaRevInfo.Fecha, "DD/MM/YYYY")
        lblTipo2.Caption = ultimaRevInfo.Tipo
        lblKmRev2.Caption = Format(TotalKm, "Standard")
    End If
End Sub

Private Sub UltimaRG(ByRef db As DAO.Database, ByRef outFechaRG As Variant, ByRef outKmRG As Double)
    Dim fechaRG As Variant
    Dim kmRG As Double
    
    If IsNumeric(Left(cmbCCRR.text, 1)) Then
        lblUltimaRG2.ForeColor = vbRed
        lblKmRG2.ForeColor = vbRed
        lblUltimaRG2.Visible = True
        lblKmRG2.Visible = True
    End If
    
    fechaRG = ObtenerMaxFecha(cmbCCRR.text, "RG", db)
    
    If IsNull(fechaRG) Then
        kmRG = 0
        If IsNumeric(Left(cmbCCRR.text, 1)) Then
            lblUltimaRG2.Caption = "Sin Datos"
            lblKmRG2.Caption = "0.00"
        End If
    Else
        kmRG = CalcularSumaKm(cmbCCRR.text, CDate(fechaRG), db)
        If IsNumeric(Left(cmbCCRR.text, 1)) Then
            lblUltimaRG2.Caption = Format(fechaRG, "DD/MM/YYYY")
            lblKmRG2.Caption = Format(kmRG, "Standard")
        End If
    End If
    
    outFechaRG = fechaRG
    outKmRG = kmRG
End Sub
Private Sub UltimaRP(ByRef db As DAO.Database, ByVal fechaBaseRG As Variant, ByVal kmDesdeRG As Double, ByRef outFechaRP As Variant, ByRef outKmRP As Double)
    Dim FechaURP As Variant
    Dim TotalKm As Double
    
    lblUltimaRP2.Visible = True
    lblKmRP2.Visible = True
    outFechaRP = Null
    outKmRP = 0
    
    FechaURP = ObtenerMaxFecha(cmbCCRR.text, "RP", db)
    
    If Not IsNull(FechaURP) And Not IsNull(fechaBaseRG) Then
        If CDate(FechaURP) < CDate(fechaBaseRG) Then
            lblUltimaRP2.Caption = "Sin RP"
            lblKmRP2.Caption = Format(kmDesdeRG, "Standard")
        Else
            TotalKm = CalcularSumaKm(cmbCCRR.text, CDate(FechaURP), db)
            lblUltimaRP2.Caption = Format(FechaURP, "DD/MM/YYYY")
            lblKmRP2.Caption = Format(TotalKm, "Standard")
            outFechaRP = FechaURP
            outKmRP = TotalKm
        End If
    ElseIf IsNull(FechaURP) Then
        lblUltimaRP2.Caption = "Sin RP"
        lblKmRP2.Caption = Format(kmDesdeRG, "Standard")
    Else ' Solo existe FechaURP
        TotalKm = CalcularSumaKm(cmbCCRR.text, CDate(FechaURP), db)
        lblUltimaRP2.Caption = Format(FechaURP, "DD/MM/YYYY")
        lblKmRP2.Caption = Format(TotalKm, "Standard")
        outFechaRP = FechaURP
        outKmRP = TotalKm
    End If
End Sub


Private Sub cmbCCRR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmdCopiar_Click()
    
    ' Llama a la rutina genérica y le dice que copie la grilla FG1
    CopiarGrilla FG1

End Sub

Private Sub cmdCopiar2_Click()
    
    ' Llama a la rutina genérica y le dice que copie la grilla FG2
    CopiarGrilla FG2

End Sub

' Esta subrutina genérica puede copiar cualquier grilla FlexGrid
Public Sub CopiarGrilla(fg As MSFlexGrid)
    Dim row As Long, col As Long
    Dim texto As String

    If fg.Rows <= 1 Then
        MsgBox "No hay datos para copiar.", vbInformation
        Exit Sub
    End If

    For row = 0 To fg.Rows - 1
        For col = 0 To fg.Cols - 1
            texto = texto & fg.TextMatrix(row, col)
            If col < fg.Cols - 1 Then texto = texto & vbTab
        Next col
        texto = texto & vbCrLf
    Next row

    Clipboard.Clear
    Clipboard.SetText texto
    MsgBox "Grilla copiada al portapapeles. Podés pegarla en Excel.", vbInformation
End Sub

Private Sub cmdKaims_Click()
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rutaBase As String, clave As String
    
    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCRR.mdb"
    clave = "theidol-1995"
    
    On Error GoTo ErrorHandler
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    LimpiarGrillasYLabels
    
    Dim esMaterfer As Boolean
    esMaterfer = IsNumeric(Left(cmbCCRR.text, 1))
    
    If esMaterfer Then
        ProcesarCocheMaterfer db
    Else
        ProcesarCocheCNR db
    End If
    
    CargarGrillaKilometraje db
    CargarGrillaIntervenciones db
    UltimaUltrasonido db
    
    lblInfo.Caption = "© 2025 SPC Software Services ®"

Cierre:
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Cierre
End Sub
Private Sub LimpiarGrillasYLabels()
' Esta subrutina limpia y oculta TODOS los resultados para una nueva consulta.
    FG1.Rows = 1
    FG2.Rows = 1
    lblTotalKm.Caption = ""
    lblTotalIntervenciones.Caption = ""

    ' --- Oculta todos los labels de resultados de Materfer ---
    lblUltimaRG2.Visible = False
    lblKmRG2.Visible = False
    lblUltimaRP2.Visible = False
    lblKmRP2.Visible = False
    lblUltimaABC2.Visible = False
    lblKmABC2.Visible = False
    lblUltimaRev2.Visible = False
    lblTipo2.Visible = False
    lblKmRev2.Visible = False
    lblFechaUS.Visible = False
    lblKmsUS.Visible = False
    
    ' --- Oculta todos los labels de resultados de CNR ---
    ' (Asegúrate de que los nombres coincidan con los de tu formulario)
    lblFechaCNR_RG.Visible = False
    lblKmCNR_RG.Visible = False
    lblFechaCNR_A3.Visible = False
    lblKmCNR_A3.Visible = False
    lblFechaCNR_A2.Visible = False
    lblKmCNR_A2.Visible = False
    lblFechaCNR_A1.Visible = False
    lblKmCNR_A1.Visible = False
    lblFechaCNR_Rev.Visible = False
    lblKmCNR_Rev.Visible = False
    lblTipoCNR_Rev.Visible = False
End Sub

Private Sub ProcesarCocheMaterfer(ByRef db As DAO.Database)
    Dim fechaBaseRG As Variant, kmDesdeRG As Double
    Dim fechaEfectivaRP As Variant, kmEfectivosRP As Double
    
    lblUltimaRG2.Visible = True
    lblKmRG2.Visible = True
    lblUltimaRP2.Visible = True
    lblKmRP2.Visible = True
    lblUltimaABC2.Visible = True
    lblKmABC2.Visible = True
    lblUltimaRev2.Visible = True
    lblTipo2.Visible = True
    lblKmRev2.Visible = True
    
    Call UltimaRG(db, fechaBaseRG, kmDesdeRG)
    Call UltimaRP(db, fechaBaseRG, kmDesdeRG, fechaEfectivaRP, kmEfectivosRP)
    Call UltimaABC(db, fechaBaseRG, kmDesdeRG, fechaEfectivaRP, kmEfectivosRP)
    Call UltimaRev(db, fechaBaseRG, kmDesdeRG)
End Sub

Private Sub ProcesarCocheCNR(ByRef db As DAO.Database)
    ' Jerarquía: RG > A3 > A2 > A1 > SEM > MEN
    
    ' --- Muestra los labels de CNR ---
    lblFechaCNR_RG.Visible = True
    lblKmCNR_RG.Visible = True
    lblFechaCNR_A3.Visible = True
    lblKmCNR_A3.Visible = True
    lblFechaCNR_A2.Visible = True
    lblKmCNR_A2.Visible = True
    lblFechaCNR_A1.Visible = True
    lblKmCNR_A1.Visible = True
    lblFechaCNR_Rev.Visible = True
    lblKmCNR_Rev.Visible = True
    lblTipoCNR_Rev.Visible = True
    
    Dim fechaEfectiva As Variant
    Dim kmEfectivos As Double
    Dim tipoEfectivo As String
    
    ' --- PASO 1: Procesar la RG (Color Rojo) ---
    Dim fechaRG As Variant
    fechaRG = ObtenerMaxFecha(cmbCCRR.text, "RG", db)
    lblFechaCNR_RG.ForeColor = vbRed
    lblKmCNR_RG.ForeColor = vbRed
    
    If Not IsNull(fechaRG) Then
        fechaEfectiva = fechaRG
        kmEfectivos = CalcularSumaKm(cmbCCRR.text, CDate(fechaEfectiva), db)
        tipoEfectivo = "RG"
    Else
        fechaEfectiva = Null
        kmEfectivos = 0
        tipoEfectivo = "N/A"
    End If
    lblFechaCNR_RG.Caption = IIf(IsNull(fechaRG), "Sin Datos", Format(fechaRG, "DD/MM/YYYY"))
    lblKmCNR_RG.Caption = Format(kmEfectivos, "Standard")
    
    ' --- PASO 2: Procesar A3 (Color Verde, como RP) ---
    Dim fechaA3 As Variant
    fechaA3 = ObtenerMaxFecha(cmbCCRR.text, "A3", db)
    lblFechaCNR_A3.ForeColor = &HC000&
    lblKmCNR_A3.ForeColor = &HC000&
    
    If Not IsNull(fechaA3) Then
        If IsNull(fechaEfectiva) Or CDate(fechaA3) > CDate(fechaEfectiva) Then
            fechaEfectiva = fechaA3
            kmEfectivos = CalcularSumaKm(cmbCCRR.text, CDate(fechaEfectiva), db)
            tipoEfectivo = "A3"
        End If
    End If
    lblFechaCNR_A3.Caption = IIf(IsNull(fechaA3), "Sin Datos", Format(fechaA3, "DD/MM/YYYY"))
    lblKmCNR_A3.Caption = Format(kmEfectivos, "Standard")
    
    ' --- PASO 3: Procesar A2 (Color Azul, como ABC) ---
    Dim fechaA2 As Variant
    fechaA2 = ObtenerMaxFecha(cmbCCRR.text, "A2", db)
    lblFechaCNR_A2.ForeColor = vbBlue
    lblKmCNR_A2.ForeColor = vbBlue
    
    If Not IsNull(fechaA2) Then
        If IsNull(fechaEfectiva) Or CDate(fechaA2) > CDate(fechaEfectiva) Then
            fechaEfectiva = fechaA2
            kmEfectivos = CalcularSumaKm(cmbCCRR.text, CDate(fechaEfectiva), db)
            tipoEfectivo = "A2"
        End If
    End If
    lblFechaCNR_A2.Caption = IIf(IsNull(fechaA2), "Sin Datos", Format(fechaA2, "DD/MM/YYYY"))
    lblKmCNR_A2.Caption = Format(kmEfectivos, "Standard")
    
    ' --- PASO 4: Procesar A1 (Usaremos Cian para distinguirlo) ---
    Dim fechaA1 As Variant
    fechaA1 = ObtenerMaxFecha(cmbCCRR.text, "A1", db)
    lblFechaCNR_A1.ForeColor = &H800080
    lblKmCNR_A1.ForeColor = &H800080
    
    If Not IsNull(fechaA1) Then
        If IsNull(fechaEfectiva) Or CDate(fechaA1) > CDate(fechaEfectiva) Then
            fechaEfectiva = fechaA1
            kmEfectivos = CalcularSumaKm(cmbCCRR.text, CDate(fechaEfectiva), db)
            tipoEfectivo = "A1"
        End If
    End If
    lblFechaCNR_A1.Caption = IIf(IsNull(fechaA1), "Sin Datos", Format(fechaA1, "DD/MM/YYYY"))
    lblKmCNR_A1.Caption = Format(kmEfectivos, "Standard")
    
    ' --- PASO 5: Procesar SEM y MEN (para la última revisión) ---
    Dim revInfo As TInfoIntervencion
    revInfo = ObtenerUltimaIntervencion(cmbCCRR.text, "'SEM', 'MEN'", db)
    
    If revInfo.Encontrado Then
        If IsNull(fechaEfectiva) Or revInfo.Fecha > CDate(fechaEfectiva) Then
            fechaEfectiva = revInfo.Fecha
            kmEfectivos = CalcularSumaKm(cmbCCRR.text, CDate(fechaEfectiva), db)
            tipoEfectivo = revInfo.Tipo
        End If
    End If
    
    ' --- Coloreamos el label de la última revisión efectiva (Color Negro) ---
    Dim colorFinal As Long
    Select Case tipoEfectivo
        Case "RG": colorFinal = vbRed
        Case "A3": colorFinal = &HC000&
        Case "A2": colorFinal = vbBlue
        Case "A1": colorFinal = vbCyan
        Case Else: colorFinal = vbBlack ' SEM, MEN y N/A quedan en negro
    End Select
    
    lblTipoCNR_Rev.ForeColor = colorFinal
    lblFechaCNR_Rev.ForeColor = colorFinal
    lblKmCNR_Rev.ForeColor = colorFinal
    
    lblTipoCNR_Rev.Caption = tipoEfectivo
    If Not IsNull(fechaEfectiva) Then
        lblFechaCNR_Rev.Caption = Format(fechaEfectiva, "DD/MM/YYYY")
        lblKmCNR_Rev.Caption = Format(kmEfectivos, "Standard")
    Else
        lblFechaCNR_Rev.Caption = "Sin Datos"
        lblKmCNR_Rev.Caption = "0.00"
    End If
End Sub
Private Sub CargarGrillaKilometraje(ByRef db As DAO.Database)
' Carga la grilla FG1 con el detalle de Kms para el período seleccionado.
    Dim tQuery As DAO.Recordset
    Dim vSQL As String
    Dim TotalKm As Double
    
    vSQL = "SELECT * FROM KILOMETRAJE WHERE Coche Like '" & cmbCCRR.text & "' AND FECHA >=#" & Format$(txtFechaDesde.text, "MM/DD/YYYY") & "# AND FECHA <=#" & Format$(txtFechaHasta.text, "MM/DD/YYYY") & "# ORDER BY FECHA"
    Set tQuery = db.OpenRecordset(vSQL, dbOpenSnapshot)
    
    FG1.TextMatrix(0, 0) = "Coche"
    FG1.TextMatrix(0, 1) = "Fecha"
    FG1.TextMatrix(0, 2) = "KM."
    
    TotalKm = 0
    If Not tQuery.EOF Then
        While Not tQuery.EOF
            FG1.AddItem tQuery!coche & vbTab & Format(tQuery!Fecha, "dd/mm/yyyy") & vbTab & Format$(tQuery!Kms_Diario, "#,##0.00")
            TotalKm = TotalKm + tQuery!Kms_Diario
            tQuery.MoveNext
        Wend
    End If
    tQuery.Close
    
    lblTotalKm.Visible = True
    lblTotalKm.Caption = "Total Km. Recorridos: " & Format(TotalKm, "#,##0.00")
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


Private Sub cmdSalir_Click()

    Unload Me

End Sub


Private Sub Form_Load()
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rutaBase As String
    Dim clave As String
    Dim tCCRR As DAO.Recordset

    FormConsultaKmCCRR.Width = 20430
    FormConsultaKmCCRR.Height = 10845
    FormConsultaKmCCRR.Top = 0
    FormConsultaKmCCRR.Left = 30
        
    rutaBase = "g:\Material Rodante\IFM\DOCUMENT\baseCCRR.mdb"
    clave = "theidol-1995"
    
    On Error GoTo ErrorHandler
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    
    Set tCCRR = db.OpenRecordset("Coches", dbOpenSnapshot)
    
    While Not tCCRR.EOF
        cmbCCRR.AddItem (tCCRR!coche)
        tCCRR.MoveNext
    Wend
    
    cmbCCRR.AddItem ("*")
    
    txtFechaDesde.text = Format((Date - 30), "DD/MM/YYYY")
    txtFechaHasta.text = Format((Date), "DD/MM/YYYY")

CierreForm:
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    Set ws = Nothing
    Set tCCRR = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error al abrir la base de datos: " & Err.Description, vbCritical
    Resume CierreForm
End Sub

Private Sub CargarGrillaIntervenciones(ByRef db As DAO.Database)
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim contador As Long

    FG2.Rows = 1
    
    sql = "SELECT Coche, Fecha_desde, Fecha_hasta, Intervencion, Lugar " & _
          "FROM Detenciones " & _
          "WHERE Coche = '" & cmbCCRR.text & "' " & _
          "AND Fecha_desde <= #" & Format(txtFechaHasta.text, "mm/dd/yyyy") & "# " & _
          "AND Fecha_hasta >= #" & Format(txtFechaDesde.text, "mm/dd/yyyy") & "# " & _
          "ORDER BY Fecha_hasta DESC"

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    FG2.TextMatrix(0, 0) = "Coche"
    FG2.TextMatrix(0, 1) = "Fecha INI"
    FG2.TextMatrix(0, 2) = "Fecha FIN"
    FG2.TextMatrix(0, 3) = "Intervención"
    FG2.TextMatrix(0, 4) = "Lugar"
    
    contador = 0
    If Not rs.EOF Then
        While Not rs.EOF
            FG2.AddItem rs!coche & vbTab & _
                        Format(rs!Fecha_desde, "dd/mm/yyyy") & vbTab & _
                        Format(rs!Fecha_hasta, "dd/mm/yyyy") & vbTab & _
                        rs!Intervencion & vbTab & _
                        rs!Lugar
            
            contador = contador + 1
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    ' --- LÍNEA CORREGIDA ---
    ' Ahora esta línea actualizará tu nueva etiqueta.
    lblTotalIntervenciones.Visible = True
    lblTotalIntervenciones.Caption = "Total Intervenciones: " & contador
    
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



