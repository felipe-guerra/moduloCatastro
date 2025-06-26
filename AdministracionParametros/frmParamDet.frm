VERSION 5.00
Begin VB.Form frmDet_Parametros 
   Caption         =   "Parámetros Generales"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   62
      Text            =   "Text2"
      Top             =   7920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   8040
      Width           =   975
   End
   Begin VB.Frame fraParametrosGenerales 
      Caption         =   "Administración de Parámetros para el proceso de Determinación."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   14775
      Begin VB.Frame fraPrediosPublicos 
         Caption         =   "Predios Públicos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   10440
         TabIndex        =   50
         Top             =   960
         Width           =   4095
         Begin VB.Frame fraExoneracion 
            Caption         =   "Exoneración"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   3855
            Begin VB.CheckBox chkValor2Publico 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor 2:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   360
               TabIndex        =   56
               Top             =   1320
               Width           =   2775
            End
            Begin VB.CheckBox chkValor1Publico 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor 1:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   360
               TabIndex        =   55
               Top             =   960
               Width           =   2775
            End
            Begin VB.CheckBox chkTasaAdminPublico 
               Caption         =   "Tasa Administrativa:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   360
               TabIndex        =   54
               Top             =   600
               Width           =   2775
            End
            Begin VB.Label lblValor2Exonera 
               BackColor       =   &H8000000A&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   3240
               TabIndex        =   59
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label lblValor1Exonera 
               BackColor       =   &H8000000A&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   3240
               TabIndex        =   58
               Top             =   960
               Width           =   495
            End
            Begin VB.Label lblTAdminExonera 
               BackColor       =   &H8000000A&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   3240
               TabIndex        =   57
               Top             =   600
               Width           =   495
            End
         End
         Begin VB.Label lblAplicaEmision 
            Caption         =   "Aplica para la Emisión:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   480
            Width           =   2415
         End
      End
      Begin VB.Frame fraDepreciacion 
         Caption         =   "Depreciación de edificación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   40
         Top             =   6240
         Width           =   10095
         Begin VB.OptionButton optRealRural 
            Caption         =   "Real"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8760
            TabIndex        =   48
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton optSimulacionRural 
            Caption         =   "Simulación"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7320
            TabIndex        =   47
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton optRealUrbana 
            Caption         =   "Real"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8760
            TabIndex        =   46
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optSimulacionUrbana 
            Caption         =   "Simulación"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7320
            TabIndex        =   45
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label lblAnioEmisionRural 
            BackColor       =   &H8000000A&
            Caption         =   "2014"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   255
            Left            =   6000
            TabIndex        =   44
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblAnioEmisionUrbana 
            BackColor       =   &H8000000A&
            Caption         =   "2014"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   255
            Left            =   6000
            TabIndex        =   43
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblDepreciacionRural 
            Caption         =   "Año de depreciación Rural:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   255
            Left            =   3000
            TabIndex        =   42
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label lblDepreciacionUrbana 
            Caption         =   "Año de depreciación Urbana:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   255
            Left            =   3000
            TabIndex        =   41
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame fraPrediosUrbanos 
         Caption         =   "Predios Urbanos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4575
         Left            =   5280
         TabIndex        =   20
         Top             =   960
         Width           =   4935
         Begin VB.Frame fraNuevosValoresUrbano 
            Caption         =   "Nuevos Valores a cobrar Urbano"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1215
            Left            =   120
            TabIndex        =   34
            Top             =   3240
            Width           =   4695
            Begin VB.TextBox txtNombreVal2Urb 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   38
               Top             =   740
               Width           =   2535
            End
            Begin VB.TextBox txtNombreVal1Urb 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   37
               Top             =   360
               Width           =   2535
            End
            Begin VB.CheckBox chkActivar2Urb 
               Alignment       =   1  'Right Justify
               Caption         =   "Activar"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   960
               TabIndex        =   36
               Top             =   760
               Width           =   975
            End
            Begin VB.CheckBox chkActivar1Urb 
               Alignment       =   1  'Right Justify
               Caption         =   "Activar"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   960
               TabIndex        =   35
               Top             =   380
               Width           =   975
            End
            Begin VB.Label lblValor2Urbano 
               Caption         =   "Valor 2:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   120
               TabIndex        =   60
               Top             =   760
               Width           =   735
            End
            Begin VB.Label lblValor1Urbano 
               Caption         =   "Valor 1:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   380
               Width           =   735
            End
         End
         Begin VB.Frame fraBomberosUrbano 
            Caption         =   "Bomberos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1215
            Left            =   120
            TabIndex        =   29
            Top             =   1920
            Width           =   4695
            Begin VB.CheckBox chkExencionesBomberosUrbano 
               Alignment       =   1  'Right Justify
               Caption         =   "Si se desea que las exenciones se apliquen para el impuesto de bomberos"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   495
               Left            =   360
               TabIndex        =   32
               Top             =   680
               Width           =   3855
            End
            Begin VB.CheckBox chkBomberosUrbano 
               Alignment       =   1  'Right Justify
               Caption         =   "Si se tiene convenio para el cobro del impuesto de bomberos"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   375
               Left            =   360
               TabIndex        =   31
               Top             =   300
               Width           =   3855
            End
         End
         Begin VB.CheckBox chkRegulacionUrbana 
            Alignment       =   1  'Right Justify
            Caption         =   "Existe regulación Urbana:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   375
            Left            =   900
            TabIndex        =   28
            Top             =   1440
            Width           =   2820
         End
         Begin VB.TextBox txtTasaAdministrativaU 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   27
            Text            =   "2.0000"
            Top             =   1000
            Width           =   975
         End
         Begin VB.TextBox txtTasaMunicipalUDecimal 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   26
            Text            =   "0.00060"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtTasaMunicipalUPorc 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   25
            Text            =   "0.6000"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblPorMilUrbano 
            Caption         =   "por mil ="
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   255
            Left            =   3000
            TabIndex        =   33
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblTasaAdministrativaU 
            Caption         =   "Tasa Administrativa: $"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   255
            Left            =   480
            TabIndex        =   24
            Top             =   1040
            Width           =   2175
         End
         Begin VB.Label lblTarifaMunicipalU 
            Caption         =   "Tarifa Municipal:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame fraPrediosRurales 
         Caption         =   "Predios Rurales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   4575
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   4935
         Begin VB.Frame fraNuevosValoresRural 
            Caption         =   "Nuevos Valores a cobrar Rural"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   1215
            Left            =   120
            TabIndex        =   12
            Top             =   3240
            Width           =   4695
            Begin VB.TextBox txtNombreVal2Rur 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   16
               Top             =   740
               Width           =   2535
            End
            Begin VB.TextBox txtNombreVal1Rur 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   15
               Top             =   360
               Width           =   2535
            End
            Begin VB.CheckBox chkActivar2Rur 
               Alignment       =   1  'Right Justify
               Caption         =   "Activa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   960
               TabIndex        =   14
               Top             =   760
               Width           =   975
            End
            Begin VB.CheckBox chkActivar1Rur 
               Alignment       =   1  'Right Justify
               Caption         =   "Activar"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   960
               TabIndex        =   13
               Top             =   380
               Width           =   975
            End
            Begin VB.Label lblValor2Rural 
               Caption         =   "Valor 2:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   760
               Width           =   760
            End
            Begin VB.Label lblValor1Rural 
               Caption         =   "Valor 1:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   380
               Width           =   735
            End
         End
         Begin VB.Frame fraBomberosRural 
            Caption         =   "Bomberos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   1215
            Left            =   120
            TabIndex        =   7
            Top             =   1920
            Width           =   4695
            Begin VB.CheckBox chkExencionesBomberosRural 
               Alignment       =   1  'Right Justify
               Caption         =   "Si se desea que las exenciones se apliquen para el impuesto de bomberos"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   495
               Left            =   240
               TabIndex        =   9
               Top             =   680
               Width           =   3855
            End
            Begin VB.CheckBox chkBomberosRural 
               Alignment       =   1  'Right Justify
               Caption         =   "Si se tiene convenio para el cobro del impuesto de bomberos"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00761616&
               Height          =   375
               Left            =   240
               TabIndex        =   8
               Top             =   300
               Width           =   3855
            End
         End
         Begin VB.CheckBox chkExcencionesRurales 
            Alignment       =   1  'Right Justify
            Caption         =   "Calcular excenciones para Predios Rurales:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   480
            Left            =   360
            TabIndex        =   6
            Top             =   1420
            Width           =   3375
         End
         Begin VB.TextBox txtTasaAdmiRural 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   5
            Text            =   "2.0000"
            Top             =   1000
            Width           =   975
         End
         Begin VB.TextBox txtTasaMunicipalRDecimal 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "0.00070"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtTasaMunicipalRPorc 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            MaxLength       =   5
            TabIndex        =   3
            Text            =   "0.7000"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblPorMilRural 
            Caption         =   "por mil ="
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   255
            Left            =   3000
            TabIndex        =   11
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblTasaAdmiRural 
            Caption         =   "Tasa Administrativa: $"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   1040
            Width           =   2175
         End
         Begin VB.Label lblTarifaMunicipalR 
            Caption         =   "Tarifa Municipal:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00761616&
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.TextBox txtBasePrestPorc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   22
         Text            =   "20.0000"
         Top             =   5740
         Width           =   1095
      End
      Begin VB.TextBox txtRBU 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "450.0000"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblBasePrest 
         Caption         =   "Base porcentual de préstamos :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00761616&
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   5760
         Width           =   3855
      End
      Begin VB.Label lblRBU 
         Caption         =   "Remuneración Básica Unificada (RBU) :  $"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00761616&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   480
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   8760
      TabIndex        =   30
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   6000
      TabIndex        =   49
      Top             =   8040
      Width           =   1215
   End
End
Attribute VB_Name = "frmDet_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=== CÓDIGO VB6 - EVENTOS Y LÓGICA DEL FORMULARIO ===

'Variables globales del formulario
Private arrNuevos As Collection
Private arrCambios As Collection
Private anioU As String
Private anioR As String
Private exoPublicoValor1 As String
Private exoPublicoValor2 As String
Private exoPublicotAdminis As String
Private tipoPredialArchivo As String
Private AñoTrabajo As Integer
Private ExoneraDominioPublico As String

'Objetos de conexión y transacción (adaptar según tu sistema)
Private con As Object 'Conexión a base de datos
Private Transac As Object 'Transacción
Private Objsp As Object 'Stored procedures
Private ObjSpTes As Object 'SP Tesorería

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Set arrNuevos = New Collection
    
    'Inicializar variables
    tipoPredialArchivo = "A" 'A = Ambos, U = Urbano, R = Rural
    AñoTrabajo = year(Now)
    ExoneraDominioPublico = "000"
    
    'Configurar conexión (adaptar según tu sistema)
    'con.TipoBDD = Nombrebdd
    'con.DataSource = servidor
    
    'Asignar eventos
    Call AsignarEventos
    
    'Inicializar formulario
    Call fnc_Inicio
    
    'Configurar visibilidad según tipo de predial
    If tipoPredialArchivo = "U" Then
        fraPrediosUrbanos.Visible = True
        fraPrediosRurales.Visible = False
    ElseIf tipoPredialArchivo = "R" Then
        fraPrediosRurales.Visible = True
        fraPrediosUrbanos.Visible = False
    Else
        fraPrediosUrbanos.Visible = True
        fraPrediosRurales.Visible = True
    End If
    
    'Cargar datos adicionales
    Call datatablemodis(arrNuevos)
    Call fnc_SelectDepreciacion
    Call fnc_SelectExoneraPublico
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error abriendo esta pantalla, por favor comuníquese con el administrador", vbCritical
End Sub

Private Sub AsignarEventos()
    'Los eventos en VB6 se asignan automáticamente por nombre
    'Este método se mantiene por compatibilidad pero no es necesario en VB6
End Sub

Private Sub fnc_Inicio()
    On Error GoTo ErrorHandler
    
    Dim rs As Object 'Recordset
    
    'Ejecutar consulta de parámetros (adaptar según tu sistema de datos)
    'Set rs = con.EjecutaRecordset("sp_ConsultaParametros")
    
    'Simulación de datos para prueba - reemplazar con datos reales
    txtTasaMunicipalUPorc.text = "0.6000"
    txtTasaMunicipalRPorc.text = "0.7000"
    txtTasaMunicipalUDecimal.text = format(CDbl(DatoACero(txtTasaMunicipalUPorc.text)) / 1000, "0.00000")
    txtTasaMunicipalRDecimal.text = format(CDbl(DatoACero(txtTasaMunicipalRPorc.text)) / 1000, "0.00000")
    
    txtTasaAdministrativaU.text = "2.0000"
    txtTasaAdmiRural.text = "2.0000"
    txtRBU.text = "450.0000"
    txtBasePrestPorc.text = "20.0000"
    
    'Configurar checkboxes por defecto
    chkBomberosUrbano.Value = 0
    chkRegulacionUrbana.Value = 0
    chkExcencionesRurales.Value = 0
    chkBomberosRural.Value = 0
    chkExencionesBomberosRural.Value = 0
    chkExencionesBomberosUrbano.Value = 0
    
    'Si hay datos reales, usar este código:
    '
    'If Not rs.EOF Then
    '    txtTasaMunicipalUPorc.Text = rs.Fields(1).Value
    '    txtTasaMunicipalRPorc.Text = rs.Fields(2).Value
    '    txtTasaMunicipalUDecimal.Text = Format(CDbl(DatoACero(txtTasaMunicipalUPorc.Text)) / 1000, "0.00000")
    '    txtTasaMunicipalRDecimal.Text = Format(CDbl(DatoACero(txtTasaMunicipalRPorc.Text)) / 1000, "0.00000")
    '
    '    txtTasaAdministrativaU.Text = rs.Fields(3).Value
    '    chkBomberosUrbano.Value = IIf(rs.Fields(4).Value, 1, 0)
    '    txtRBU.Text = rs.Fields(5).Value
    '    chkRegulacionUrbana.Value = IIf(rs.Fields(6).Value, 1, 0)
    '    txtBasePrestPorc.Text = rs.Fields(7).Value
    '
    '    'Continuar con el resto de campos...
    'End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo ErrorHandler
    
    Dim resultado As Integer
    
    'Validar datos antes de grabar
    If Not ValidarTodos() Then
        Exit Sub
    End If
    
    'Iniciar transacción (adaptar según tu sistema)
    'If con.FncConexion() Then
    '    Set Transac = con.IniciaTrans()
    
        'RBU
        'con.LimpiarParametro()
        'con.agregaParametro("@RBU", CDbl(txtRBU.Text))
        'con.agregaParametro("@Codigo", "0205")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'TARIFA MUNICIPAL URBANA
        'con.LimpiarParametro()
        'con.agregaParametro("@TarifaU", CDbl(txtTasaMunicipalUPorc.Text))
        'con.agregaParametro("@Codigo", "0201")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'TARIFA MUNICIPAL RURAL
        'con.LimpiarParametro()
        'con.agregaParametro("@TarifaR", CDbl(txtTasaMunicipalRPorc.Text))
        'con.agregaParametro("@Codigo", "0202")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'TARIFA ADMINISTRATIVA URBANA
        'con.LimpiarParametro()
        'con.agregaParametro("@TarifaAdmi", CDbl(txtTasaAdministrativaU.Text))
        'con.agregaParametro("@Codigo", "0203")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'TARIFA ADMINISTRATIVA RURAL
        'con.LimpiarParametro()
        'con.agregaParametro("@TarifaAdmiRural", CDbl(txtTasaAdmiRural.Text))
        'con.agregaParametro("@Codigo", "0209")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'BASE PORCENTUAL DE PRESTAMOS
        'con.LimpiarParametro()
        'con.agregaParametro("@TarifaU", CDbl(txtBasePrestPorc.Text))
        'con.agregaParametro("@Codigo", "0207")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'BOMBEROS URBANO
        Dim tmpBomberos As String
        If chkBomberosUrbano.Value = 1 Then
            tmpBomberos = "1"
        Else
            tmpBomberos = "0"
        End If
        
        'con.LimpiarParametro()
        'con.agregaParametro("@Bomberos", tmpBomberos)
        'con.agregaParametro("@Codigo", "0204")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'EXENCIONES BOMBEROS URBANO
        If chkExencionesBomberosUrbano.Value = 1 Then
            tmpBomberos = "1"
        Else
            tmpBomberos = "0"
        End If
        
        'con.LimpiarParametro()
        'con.agregaParametro("@excBomberos", tmpBomberos)
        'con.agregaParametro("@Codigo", "0212")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'BOMBEROS RURAL
        Dim tmpBomberosRural As String
        If chkBomberosRural.Value = 1 Then
            tmpBomberosRural = "1"
        Else
            tmpBomberosRural = "0"
        End If
        
        'con.LimpiarParametro()
        'con.agregaParametro("@Bomberos", tmpBomberosRural)
        'con.agregaParametro("@Codigo", "0210")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'EXENCIONES BOMBEROS RURAL
        If chkExencionesBomberosRural.Value = 1 Then
            tmpBomberos = "1"
        Else
            tmpBomberos = "0"
        End If
        
        'con.LimpiarParametro()
        'con.agregaParametro("@excBomberos", tmpBomberos)
        'con.agregaParametro("@Codigo", "0211")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'EXCENCIONES RURALES
        Dim tmpexcenciones As String
        If chkExcencionesRurales.Value = 1 Then
            tmpexcenciones = "1"
        Else
            tmpexcenciones = "0"
        End If
        
        'con.LimpiarParametro()
        'con.agregaParametro("@ExcencionesRur", tmpexcenciones)
        'con.agregaParametro("@Codigo", "0208")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'REGULACION URBANA
        Dim tmpREGuRBANA As String
        If chkRegulacionUrbana.Value = 1 Then
            tmpREGuRBANA = "1"
        Else
            tmpREGuRBANA = "0"
        End If
        
        'con.LimpiarParametro()
        'con.agregaParametro("@Bomberos", tmpREGuRBANA)
        'con.agregaParametro("@Codigo", "0206")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametros")
        
        'NUEVOS VALORES
        'VALOR 1 URBANO
        If chkActivar1Urb.Value = 1 Then
            If Trim(txtNombreVal1Urb.text) = "" Then
                MsgBox "Ingrese el nombre del nuevo valor 1 para Urbano", vbInformation
                Exit Sub
            End If
        End If
        
        'con.LimpiarParametro()
        'con.agregaParametro("@CoeDeE_Descripcion", txtNombreVal1Urb.Text)
        'con.agregaParametro("@CoeDeE_Codigo", "0213")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametrosDescripcion")
        
        'VALOR 2 URBANO
        If chkActivar2Urb.Value = 1 Then
            If Trim(txtNombreVal2Urb.text) = "" Then
                MsgBox "Ingrese el nombre del nuevo valor 2 para Urbano", vbInformation
                Exit Sub
            End If
        End If
        
        'con.LimpiarParametro()
        'con.agregaParametro("@nombreval2", txtNombreVal2Urb.Text)
        'con.agregaParametro("@Codigo", "0214")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametrosDescripcion")
        
        'VALOR 1 RURAL
        If chkActivar1Rur.Value = 1 Then
            If Trim(txtNombreVal1Rur.text) = "" Then
                MsgBox "Ingrese el nombre del nuevo valor 1 para Rural", vbInformation
                Exit Sub
            End If
        End If
        
        'con.LimpiarParametro()
        'con.agregaParametro("@nombreval1", txtNombreVal1Rur.Text)
        'con.agregaParametro("@Codigo", "0215")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametrosDescripcion")
        
        'VALOR 2 RURAL
        If chkActivar2Rur.Value = 1 Then
            If Trim(txtNombreVal2Rur.text) = "" Then
                MsgBox "Ingrese el nombre del nuevo valor 2 para Rural", vbInformation
                Exit Sub
            End If
        End If
        
        'con.LimpiarParametro()
        'con.agregaParametro("@nombreval2", txtNombreVal2Rur.Text)
        'con.agregaParametro("@Codigo", "0216")
        'resultado = con.EjecutaNonQueryTrans("sp_ActualizaParametrosDescripcion")
        
    'End If
    
    'Confirmar transacción
    'Transac.Commit()
    
    'Guardar datos adicionales
    Call fnc_InsertDepreciacion
    Call fnc_InsertExoneraPublico
    
    MsgBox "Parámetros guardados exitosamente", vbInformation
    
    'Auditoría
    Set arrCambios = New Collection
    Call datatablemodis(arrCambios)
    
    'Cerrar y reabrir formulario
    'Dim objMe As New frmParametros
    'objMe.Show
    'Unload Me
    
    Exit Sub
    
ErrorHandler:
    'Transac.Rollback()
    MsgBox "Parámetros no se han guardado, verifique los datos", vbInformation
End Sub

Private Function ValidarTodos() As Boolean
    ValidarTodos = True
    
    'Validar campos obligatorios
    If Not IsNumeric(txtRBU.text) Or val(txtRBU.text) <= 0 Then
        MsgBox "La Remuneración Básica Unificada debe ser un valor numérico mayor a cero", vbExclamation
        txtRBU.SetFocus
        ValidarTodos = False
        Exit Function
    End If
    
    If Not IsNumeric(txtTasaMunicipalRPorc.text) Or val(txtTasaMunicipalRPorc.text) < 0 Then
        MsgBox "La Tarifa Municipal Rural debe ser un valor numérico válido", vbExclamation
        txtTasaMunicipalRPorc.SetFocus
        ValidarTodos = False
        Exit Function
    End If
    
    If Not IsNumeric(txtTasaMunicipalUPorc.text) Or val(txtTasaMunicipalUPorc.text) < 0 Then
        MsgBox "La Tarifa Municipal Urbana debe ser un valor numérico válido", vbExclamation
        txtTasaMunicipalUPorc.SetFocus
        ValidarTodos = False
        Exit Function
    End If
    
    'Validar rangos
    If CDbl(txtTasaMunicipalRPorc.text) < 0.25 Or CDbl(txtTasaMunicipalRPorc.text) > 3 Then
        MsgBox "Valor de la Tarifa Municipal Rural está fuera del rango: (0,25 a 3 por mil)", vbInformation
        ValidarTodos = False
        Exit Function
    End If
    
    If CDbl(txtTasaMunicipalUPorc.text) < 0.25 Or CDbl(txtTasaMunicipalUPorc.text) > 5 Then
        MsgBox "Valor de la Tarifa Municipal Urbana está fuera del rango: (0,25 a 5 por mil)", vbInformation
        ValidarTodos = False
        Exit Function
    End If
End Function

'=== EVENTOS KEYPRESS PARA VALIDACIÓN ===
Private Sub txtTasaMunicipalUPorc_KeyPress(KeyAscii As Integer)
    Call ValidarDecimales(KeyAscii, txtTasaMunicipalUPorc.text)
End Sub

Private Sub txtTasaAdministrativaU_KeyPress(KeyAscii As Integer)
    Call ValidarDecimales(KeyAscii, txtTasaAdministrativaU.text)
End Sub

Private Sub txtTasaAdmiRural_KeyPress(KeyAscii As Integer)
    Call ValidarDecimales(KeyAscii, txtTasaAdmiRural.text)
End Sub

Private Sub txtTasaMunicipalRPorc_KeyPress(KeyAscii As Integer)
    Call ValidarDecimales(KeyAscii, txtTasaMunicipalRPorc.text)
End Sub

Private Sub txtRBU_KeyPress(KeyAscii As Integer)
    Call ValidarDecimales(KeyAscii, txtRBU.text)
End Sub

Private Sub txtBasePrestPorc_KeyPress(KeyAscii As Integer)
    Call ValidarDecimales(KeyAscii, txtBasePrestPorc.text)
End Sub

'Función para validar solo números decimales
Private Sub ValidarDecimales(KeyAscii As Integer, TextoActual As String)
    'Permitir números, punto decimal, backspace y delete
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And _
       KeyAscii <> 46 And KeyAscii <> 8 And KeyAscii <> 127 Then
        KeyAscii = 0
    End If
    
    'Solo permitir un punto decimal
    If KeyAscii = 46 And InStr(TextoActual, ".") > 0 Then
        KeyAscii = 0
    End If
End Sub

'=== EVENTOS KEYUP PARA CÁLCULOS AUTOMÁTICOS ===
Private Sub txtTasaMunicipalUPorc_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If txtTasaMunicipalUPorc.text <> "" Then
        txtTasaMunicipalUDecimal.text = format(CDbl(txtTasaMunicipalUPorc.text) / 1000, "0.00000")
    Else
        txtTasaMunicipalUDecimal.text = "0.00000"
    End If
End Sub

Private Sub txtTasaMunicipalRPorc_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If txtTasaMunicipalRPorc.text <> "" Then
        txtTasaMunicipalRDecimal.text = format(CDbl(txtTasaMunicipalRPorc.text) / 1000, "0.00000")
    Else
        txtTasaMunicipalRDecimal.text = "0.00000"
    End If
End Sub

'=== EVENTOS CLICK PARA CHECKBOXES ===
Private Sub chkActivar1Rur_Click()
    Call fnActivarTextVal(chkActivar1Rur, txtNombreVal1Rur)
End Sub

Private Sub chkActivar2Rur_Click()
    Call fnActivarTextVal(chkActivar2Rur, txtNombreVal2Rur)
End Sub

Private Sub chkActivar1Urb_Click()
    Call fnActivarTextVal(chkActivar1Urb, txtNombreVal1Urb)
End Sub

Private Sub chkActivar2Urb_Click()
    Call fnActivarTextVal(chkActivar2Urb, txtNombreVal2Urb)
End Sub

Private Sub fnActivarTextVal(chk As CheckBox, txt As TextBox)
    If chk.Value = 1 Then
        txt.Enabled = True
    Else
        txt.Enabled = False
        txt.text = ""
    End If
End Sub

'=== EVENTOS PARA PREDIOS PÚBLICOS ===
Private Sub chkTasaAdminPublico_Click()
    If chkTasaAdminPublico.Value = 1 Then
        lblTAdminExonera.Caption = "SI"
    Else
        lblTAdminExonera.Caption = "NO"
    End If
End Sub

Private Sub chkValor1Publico_Click()
    If chkValor1Publico.Value = 1 Then
        lblValor1Exonera.Caption = "SI"
    Else
        lblValor1Exonera.Caption = "NO"
    End If
End Sub

Private Sub chkValor2Publico_Click()
    If chkValor2Publico.Value = 1 Then
        lblValor2Exonera.Caption = "SI"
    Else
        lblValor2Exonera.Caption = "NO"
    End If
End Sub

'=== EVENTOS PARA DEPRECIACIÓN ===
Private Sub optRealUrbana_Click()
    If optRealUrbana.Value = True Then
        MsgBox "Si selecciona REAL, se ejecuta el proceso de depreciación en forma definitiva", vbInformation
    End If
End Sub

Private Sub optRealRural_Click()
    If optRealRural.Value = True Then
        MsgBox "Si selecciona REAL, se ejecuta el proceso de depreciación en forma definitiva", vbInformation
    End If
End Sub

'=== FUNCIONES DE DEPRECIACIÓN ===
Private Sub fnc_InsertDepreciacion()
    On Error Resume Next
    
    Dim val As String
    
    'URBANA
    If optRealUrbana.Value = True Then
        val = "2"
    ElseIf optSimulacionUrbana.Value = True Then
        val = "1"
    End If
    
    'Aquí iría la actualización en base de datos
    'con.LimpiarParametro()
    'con.agregaParametro("@valor", Mid(anioU, 1, 4) & val)
    'con.agregaParametro("@Codigo", "0901")
    'con.EjecutaNonQueryTrans("sp_ActualizaParametros")
    
    'RURAL
    If optRealRural.Value = True Then
        val = "2"
    ElseIf optSimulacionRural.Value = True Then
        val = "1"
    End If
    
    'con.LimpiarParametro()
    'con.agregaParametro("@valor", Mid(anioR, 1, 4) & val)
    'con.agregaParametro("@Codigo", "0902")
    'con.EjecutaNonQueryTrans("sp_ActualizaParametros")
End Sub

Private Sub fnc_SelectDepreciacion()
    On Error Resume Next
    
    'Consultar parámetros de depreciación
    'anioU = ConsultarParametrogeneral("0901")
    'anioR = ConsultarParametrogeneral("0902")
    
    'Valores por defecto para prueba
    anioU = "20141"
    anioR = "20141"
    
    If (AñoTrabajo Mod 2 <> 0) Then
        lblAnioEmisionUrbana.Caption = CStr(AñoTrabajo + 1)
        lblAnioEmisionRural.Caption = CStr(AñoTrabajo + 1)
        
        'URBANO
        If Mid(anioU, 5, 1) = "1" Then
            optSimulacionUrbana.Value = True
        ElseIf Mid(anioU, 5, 1) = "2" Then
            optRealUrbana.Value = True
        ElseIf Mid(anioU, 5, 1) = "3" Then
            optRealUrbana.Value = True
            optSimulacionUrbana.Enabled = False
            optRealUrbana.Enabled = False
        End If
        
        'RURAL
        If Mid(anioR, 5, 1) = "1" Then
            optSimulacionRural.Value = True
        ElseIf Mid(anioR, 5, 1) = "2" Then
            optRealRural.Value = True
        ElseIf Mid(anioR, 5, 1) = "3" Then
            optRealRural.Value = True
            optSimulacionRural.Enabled = False
            optRealRural.Enabled = False
        End If
    Else
        lblAnioEmisionUrbana.Caption = Mid(anioU, 1, 4)
        lblAnioEmisionRural.Caption = Mid(anioR, 1, 4)
        
        optRealUrbana.Value = True
        optSimulacionUrbana.Enabled = False
        optRealUrbana.Enabled = False
        optRealRural.Value = True
        optSimulacionRural.Enabled = False
        optRealRural.Enabled = False
    End If
End Sub

'=== FUNCIONES DE EXONERACIÓN PÚBLICA ===
Private Sub fnc_SelectExoneraPublico()
    On Error Resume Next
    
    'Consultar parámetros de exoneración
    'exoPublicoValor1 = ConsultarParametrogeneral("1101")
    'exoPublicoValor2 = ConsultarParametrogeneral("1102")
    'exoPublicotAdminis = ConsultarParametrogeneral("1103")
    
    'Valores por defecto para prueba
    exoPublicoValor1 = "0"
    exoPublicoValor2 = "0"
    exoPublicotAdminis = "0"
    
    If exoPublicoValor1 = "1" Then
        chkValor1Publico.Value = 1
        lblValor1Exonera.Caption = "SI"
    Else
        chkValor1Publico.Value = 0
        lblValor1Exonera.Caption = "NO"
    End If
    
    If exoPublicoValor2 = "1" Then
        chkValor2Publico.Value = 1
        lblValor2Exonera.Caption = "SI"
    Else
        chkValor2Publico.Value = 0
        lblValor2Exonera.Caption = "NO"
    End If
    
    If exoPublicotAdminis = "1" Then
        chkTasaAdminPublico.Value = 1
        lblTAdminExonera.Caption = "SI"
    Else
        chkTasaAdminPublico.Value = 0
        lblTAdminExonera.Caption = "NO"
    End If
End Sub

Private Sub fnc_InsertExoneraPublico()
    On Error Resume Next
    
    Dim val As String
    
    'VALOR 1
    If chkValor1Publico.Value = 1 Then
        val = "1"
    Else
        val = "0"
    End If
    
    ExoneraDominioPublico = val & Mid(ExoneraDominioPublico, 2)
    
    'con.LimpiarParametro()
    'con.agregaParametro("@valor", val)
    'con.agregaParametro("@Codigo", "1101")
    'con.EjecutaNonQueryTrans("sp_ActualizaParametros")
    
    'VALOR 2
    If chkValor2Publico.Value = 1 Then
        val = "1"
    Else
        val = "0"
    End If
    
    ExoneraDominioPublico = Mid(ExoneraDominioPublico, 1, 1) & val & Mid(ExoneraDominioPublico, 3)
    
    'con.LimpiarParametro()
    'con.agregaParametro("@valor", val)
    'con.agregaParametro("@Codigo", "1102")
    'con.EjecutaNonQueryTrans("sp_ActualizaParametros")
    
    'TASA ADMINISTRATIVA
    If chkTasaAdminPublico.Value = 1 Then
        val = "1"
    Else
        val = "0"
    End If
    
    ExoneraDominioPublico = Mid(ExoneraDominioPublico, 1, 2) & val
    
    'con.LimpiarParametro()
    'con.agregaParametro("@valor", val)
    'con.agregaParametro("@Codigo", "1103")
    'con.EjecutaNonQueryTrans("sp_ActualizaParametros")
End Sub

'=== FUNCIONES AUXILIARES ===
Private Function DatoACero(valor As String) As String
    If valor = "" Or Not IsNumeric(valor) Then
        DatoACero = "0"
    Else
        DatoACero = valor
    End If
End Function

Private Sub datatablemodis(arr As Collection)
    On Error Resume Next
    
    arr.Add "Remuneración Básica Unificada: " & txtRBU.text
    arr.Add "Tarifa Mun Rural: " & txtTasaMunicipalRDecimal.text
    arr.Add "Tasa Administrativa Rural: " & txtTasaAdmiRural.text
    arr.Add "Exenciones Rurales: " & IIf(chkExcencionesRurales.Value = 1, "True", "False")
    arr.Add "Bomberos - Rural: " & IIf(chkBomberosRural.Value = 1, "True", "False")
    arr.Add "Exenciones para Bomberos Rural: " & IIf(chkExencionesBomberosRural.Value = 1, "True", "False")
    arr.Add "Tarifa Mun Urbana: " & txtTasaMunicipalUDecimal.text
    arr.Add "Tasa Administrativa Urbana: " & txtTasaAdministrativaU.text
    arr.Add "Regulación Urbana: " & IIf(chkRegulacionUrbana.Value = 1, "True", "False")
    arr.Add "Bomberos - Urbano: " & IIf(chkBomberosUrbano.Value = 1, "True", "False")
    arr.Add "Exenciones para Bomberos Urbano: " & IIf(chkExencionesBomberosUrbano.Value = 1, "True", "False")
    arr.Add "Base porcentual de préstamos: " & txtBasePrestPorc.text
End Sub

'=== FUNCIÓN PARA EXPORTAR PARÁMETROS ===
Public Function ExportarParametrosCompletos() As String
    Dim params As String
    
    params = "=== EXPORT COMPLETO DE PARÁMETROS ===" & vbCrLf & vbCrLf
    params = params & "Fecha de Exportación: " & format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf & vbCrLf
    
    params = params & "--- PARÁMETROS GENERALES ---" & vbCrLf
    params = params & "RBU: $" & txtRBU.text & vbCrLf
    params = params & "Base Porcentual Préstamos: " & txtBasePrestPorc.text & "%" & vbCrLf & vbCrLf
    
    params = params & "--- PREDIOS RURALES ---" & vbCrLf
    params = params & "Tarifa Municipal: " & txtTasaMunicipalRPorc.text & " (por mil: " & txtTasaMunicipalRDecimal.text & ")" & vbCrLf
    params = params & "Tasa Administrativa: $" & txtTasaAdmiRural.text & vbCrLf
    params = params & "Calcular Exenciones: " & IIf(chkExcencionesRurales.Value = 1, "SI", "NO") & vbCrLf
    params = params & "Convenio Bomberos: " & IIf(chkBomberosRural.Value = 1, "SI", "NO") & vbCrLf
    params = params & "Exenciones Bomberos: " & IIf(chkExencionesBomberosRural.Value = 1, "SI", "NO") & vbCrLf
    
    If chkActivar1Rur.Value = 1 Then
        params = params & "Nuevo Valor 1: " & txtNombreVal1Rur.text & vbCrLf
    End If
    If chkActivar2Rur.Value = 1 Then
        params = params & "Nuevo Valor 2: " & txtNombreVal2Rur.text & vbCrLf
    End If
    
    params = params & vbCrLf & "--- PREDIOS URBANOS ---" & vbCrLf
    params = params & "Tarifa Municipal: " & txtTasaMunicipalUPorc.text & " (por mil: " & txtTasaMunicipalUDecimal.text & ")" & vbCrLf
    params = params & "Tasa Administrativa: $" & txtTasaAdministrativaU.text & vbCrLf
    params = params & "Regulación Urbana: " & IIf(chkRegulacionUrbana.Value = 1, "SI", "NO") & vbCrLf
    params = params & "Convenio Bomberos: " & IIf(chkBomberosUrbano.Value = 1, "SI", "NO") & vbCrLf
    params = params & "Exenciones Bomberos: " & IIf(chkExencionesBomberosUrbano.Value = 1, "SI", "NO") & vbCrLf
    
    If chkActivar1Urb.Value = 1 Then
        params = params & "Nuevo Valor 1: " & txtNombreVal1Urb.text & vbCrLf
    End If
    If chkActivar2Urb.Value = 1 Then
        params = params & "Nuevo Valor 2: " & txtNombreVal2Urb.text & vbCrLf
    End If
    
    params = params & vbCrLf & "--- PREDIOS PÚBLICOS ---" & vbCrLf
    params = params & "Exoneración Tasa Administrativa: " & lblTAdminExonera.Caption & vbCrLf
    params = params & "Exoneración Valor 1: " & lblValor1Exonera.Caption & vbCrLf
    params = params & "Exoneración Valor 2: " & lblValor2Exonera.Caption & vbCrLf
    
    params = params & vbCrLf & "--- DEPRECIACIÓN ---" & vbCrLf
    params = params & "Depreciación Urbana: " & IIf(optSimulacionUrbana.Value, "Simulación", "Real") & " (" & lblAnioEmisionUrbana.Caption & ")" & vbCrLf
    params = params & "Depreciación Rural: " & IIf(optSimulacionRural.Value, "Simulación", "Real") & " (" & lblAnioEmisionRural.Caption & ")" & vbCrLf
    
    ExportarParametrosCompletos = params
End Function

