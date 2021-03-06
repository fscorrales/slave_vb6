VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form LiquidacionGanancia4ta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Retención Ganancias"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRecalcularDatos 
      BackColor       =   &H008080FF&
      Caption         =   "Recalcular Datos"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7320
      Width           =   3255
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7920
      Width           =   3255
   End
   Begin VB.Frame Frame8 
      Caption         =   "Identificación Período de Retención"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   61
      Top             =   120
      Width           =   9975
      Begin VB.CheckBox chkLiquidacionFinal 
         Alignment       =   1  'Right Justify
         Caption         =   "Liq. Anual / Final"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   7320
         TabIndex        =   68
         Top             =   360
         Width           =   2350
      End
      Begin VB.TextBox txtCodigoLiquidacion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcionPeriodo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label24 
         Caption         =   "Código Liq."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3480
         TabIndex        =   62
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "6) Importe a Retener Gcia."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4095
      Left            =   6840
      TabIndex        =   51
      Top             =   4440
      Width           =   3255
      Begin VB.TextBox txtBaseImponible 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSumaVariable 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   25
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtPorcentajeAplicable 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtAjuesteRetencion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   29
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtSubtotalRentencion 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtSumaFija 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtRetencionAcumulada 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   28
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtRetencionPeriodo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   30
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label29 
         Caption         =   "Base Imponible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Suma Variable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "% Aplicable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Ajuste"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Subtotal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Suma Fija"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Ret. Acum."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Ret. Período"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   3720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "3) Gcia. Neta Acum. (Desc. Oblig.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   3480
      TabIndex        =   47
      Top             =   2040
      Width           =   3255
      Begin VB.TextBox txtGananciaPeriodo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtGananciaNeta 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtDescuentoAcumulado 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtRentaAcumulada 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label28 
         Caption         =   "Gcia. Período"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Gcia. Neta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Desc. Acum."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Renta Acum."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "5) Deduc. Personales Acum."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   6840
      TabIndex        =   46
      Top             =   2040
      Width           =   3255
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgDeduccionesPersonales 
         Height          =   1860
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   3281
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "1) Renta Imponible"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   120
      TabIndex        =   37
      Top             =   2040
      Width           =   3255
      Begin VB.TextBox txtPluriempleo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtAjuste 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtHaberOptimo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSubtotalSueldo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "Pluriempleo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Haber Óptimo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Ajuste"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Subtotal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1800
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificación Agente a Retener"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   36
      Top             =   1080
      Width           =   9975
      Begin VB.TextBox txtDescripcionAgente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4800
         TabIndex        =   4
         Top             =   360
         Width           =   4935
      End
      Begin VB.TextBox txtPuestoLaboral 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3480
         TabIndex        =   60
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Pueto Laboral"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "4) Deduc. Generales Acum."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2775
      Left            =   3480
      TabIndex        =   35
      Top             =   4440
      Width           =   3255
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgDeduccionesGenerales 
         Height          =   2340
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   4128
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "2) Descuentos Recibo INVICO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   3255
      Begin VB.TextBox txtOtrosDescuentos 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox txtSubtotalDescuento 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox txtCuotaSindical 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtAdherente 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtSeguroObligatorio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtSeguroOptativo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtJubilacion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtObraSocial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label27 
         Caption         =   "Otros Desc."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Subtotal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Cuota Sindical"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Adherente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Seguro Oblig."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Seguro Opt."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Jubilación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Obra Social"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   1455
      End
   End
End
Attribute VB_Name = "LiquidacionGanancia4ta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    
    Dim strPL As String
    Dim strCL As String
    Dim bolLiquidacionFinal As Boolean
    
    If ValidarLiquidacionGanancias = True Then
        strPL = Me.txtPuestoLaboral.Text
        strCL = Left(Me.txtCodigoLiquidacion.Text, 4)
        bolLiquidacionFinal = Me.chkLiquidacionFinal.Value
        Call RecalcularLiquidacionGanancia4ta(strPL, strCL, False, bolLiquidacionFinal)
        GenerarLiquidacionGanancia4ta
    End If
    
    'Viejo proceso - Por qué lo diseñé así?
    'If ValidarLiquidacionGanancias = True Then
    '    If bolEditandoRetencionGanancias = False Then
    '        Call LiquidacionGanancias("PasoDos")
    '        'Call LiquidacionGanancias("PasoTres")
    '        Call LiquidacionGanancias("PasoCuatro")
    '        Call LiquidacionGanancias("PasoCinco")
    '        Call LiquidacionGanancias("PasoSiete")
    '    End If
    '    Call ControlSueldoLiquidado("Final")
    '    GenerarLiquidacionGanancia4ta
    'End If

End Sub

Private Sub cmdRecalcularDatos_Click()

    Dim strPL As String
    Dim strCL As String
    Dim bolLiquidacionFinal As Boolean
    
    
    
    If ValidarLiquidacionGanancias = True Then
        bolLiquidacionFinal = Me.chkLiquidacionFinal.Value
        strPL = Me.txtPuestoLaboral.Text
        strCL = Left(Me.txtCodigoLiquidacion.Text, 4)
        Call RecalcularLiquidacionGanancia4ta(strPL, strCL, True, bolLiquidacionFinal)
    End If
    'El viejo procedimiento es el siquiente
    'If ValidarLiquidacionGanancias = True Then
    '    Call LiquidacionGanancias("PasoDos")
    '    Call LiquidacionGanancias("PasoTres")
    '    Call LiquidacionGanancias("PasoCuatro")
    '    Call LiquidacionGanancias("PasoCinco")
    '    Call LiquidacionGanancias("PasoSiete")
    '    Call ControlSueldoLiquidado("Final")
    'End If

    strPL = ""
    strCL = ""

End Sub

Private Sub Form_Load()

    Call CenterMe(LiquidacionGanancia4ta, 10305, 8940)

End Sub

Private Sub txtPuestoLaboral_LostFocus()

    Dim SQL As String
    Dim strPL As String
    Dim strCL As String
    
    strPL = Me.txtPuestoLaboral.Text
    strCL = Left(Me.txtCodigoLiquidacion.Text, 4)
    
    If bolEditandoRetencionGanancias = False Then
        SQL = "Select * From LIQUIDACIONGANANCIAS4TACATEGORIA" _
        & " Where CODIGOLIQUIDACION = '" & strCL _
        & "' And PUESTOLABORAL = '" & strPL & "'"
        If SQLNoMatch(SQL) = True Then
            SQL = "Select * From AGENTES" _
            & " Where PUESTOLABORAL = '" & strPL & "'"
            If SQLNoMatch(SQL) = True Then
                MsgBox "Debe ingresar un Nro de Puesto Laboral válido", vbCritical + vbOKOnly, "NRO PUESTO LABORAL INEXISTENTE"
                Me.txtDescripcionAgente.Text = ""
                Me.txtPuestoLaboral.SetFocus
            Else
                Set rstBuscarSlave = New ADODB.Recordset
                rstBuscarSlave.Open SQL, dbSlave, adOpenDynamic, adLockOptimistic
                Me.txtDescripcionAgente.Text = rstBuscarSlave!NombreCompleto
                rstBuscarSlave.Close
                Set rstBuscarSlave = Nothing
                'El nuevo procedimiento es el siguiente
                Call CargaCompletaLiquidacionGanancia4ta(strPL, strCL)
                'El viejo procedimiento empezaba de la siquiente forma
                'Call LiquidacionGanancias("PasoUno", strPL, strCL)
            End If
        Else 'Raro este Else
            MsgBox "El Puesto Laboral que pretende liquidar ya posee liquidación", vbCritical + vbOKOnly, "LIQUIDACIÓN GANANCIAS DUPLICADA"
            Me.txtDescripcionAgente.Text = ""
            Me.txtPuestoLaboral.SetFocus
        End If
    End If
    
    SQL = ""
    strPL = ""
    strCL = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    bolEditandoRetencionGanancias = False

End Sub
