VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBolPag 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Emisión Boleta de Pago"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "FrmBolPag.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11400
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar Contable"
      Height          =   255
      Left            =   9510
      TabIndex        =   64
      Top             =   30
      Width           =   1845
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   375
      Left            =   7080
      TabIndex        =   48
      Top             =   1170
      Width           =   315
      _Version        =   65536
      _ExtentX        =   556
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "."
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   5055
      Left            =   10500
      TabIndex        =   41
      Top             =   2910
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   8916
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin VB.CommandButton CmdModifi 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   60
         Picture         =   "FrmBolPag.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   62
         Tag             =   "M"
         ToolTipText     =   "Modificar"
         Top             =   1770
         Width           =   1065
      End
      Begin VB.CommandButton CmdCerrar 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   60
         Picture         =   "FrmBolPag.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Cerrar"
         Top             =   4200
         Width           =   1065
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   60
         Picture         =   "FrmBolPag.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Buscar"
         Top             =   3390
         Width           =   1065
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   60
         Picture         =   "FrmBolPag.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   42
         Tag             =   "N"
         ToolTipText     =   "Registrar "
         Top             =   150
         Width           =   1065
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   60
         Picture         =   "FrmBolPag.frx":1772
         Style           =   1  'Graphical
         TabIndex        =   43
         Tag             =   "E"
         ToolTipText     =   "Eliminar"
         Top             =   960
         Width           =   1065
      End
      Begin VB.CommandButton CmdImp 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   60
         Picture         =   "FrmBolPag.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Imprimir"
         Top             =   2580
         Width           =   1065
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   8175
      Left            =   6210
      TabIndex        =   15
      Top             =   150
      Width           =   5595
      _Version        =   65536
      _ExtentX        =   9869
      _ExtentY        =   14420
      _StockProps     =   14
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      ShadowStyle     =   1
      Begin Threed.SSFrame SSFrame5 
         Height          =   765
         Left            =   2580
         TabIndex        =   56
         Top             =   1860
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   1349
         _StockProps     =   14
         Caption         =   "Forma de Pago"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         ShadowStyle     =   1
         Begin Threed.SSOption OptPro 
            Height          =   375
            Left            =   1290
            TabIndex        =   58
            Top             =   315
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "Pronto Pago"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.74
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption OptEfe 
            Height          =   375
            Left            =   90
            TabIndex        =   57
            Top             =   315
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "Efectivo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.74
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   375
         Left            =   3510
         TabIndex        =   1
         Top             =   1260
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   59441153
         CurrentDate     =   37547
      End
      Begin VB.TextBox TxtNroBol 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3510
         MaxLength       =   6
         TabIndex        =   18
         Top             =   780
         Width           =   1710
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   2595
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   2505
         _Version        =   65536
         _ExtentX        =   4419
         _ExtentY        =   4577
         _StockProps     =   14
         Caption         =   "O p c i o n e s"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Font3D          =   3
         ShadowStyle     =   1
         Begin MSDataListLib.DataCombo CmbAno 
            Height          =   360
            Left            =   120
            TabIndex        =   0
            Tag             =   "8"
            Top             =   1530
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   "Id_ano"
            BoundColumn     =   "Id_ano"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox MbNAsie 
            Height          =   350
            Left            =   120
            TabIndex        =   49
            Top             =   420
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo CmbMes1 
            Height          =   360
            Left            =   120
            TabIndex        =   50
            Tag             =   "8"
            Top             =   2145
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   "DescMes"
            BoundColumn     =   "Id_Mes"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox MbTipC 
            Height          =   345
            Left            =   120
            TabIndex        =   51
            Top             =   990
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   609
            _Version        =   393216
            ClipMode        =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "T.Cam."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   13
            Left            =   120
            TabIndex        =   54
            Top             =   780
            Width           =   750
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mes C :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   19
            Left            =   120
            TabIndex        =   53
            Top             =   1920
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nº Asto.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   14
            Left            =   120
            TabIndex        =   52
            Top             =   195
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   18
            Left            =   120
            TabIndex        =   47
            Top             =   1305
            Width           =   420
         End
      End
      Begin MSMask.MaskEdBox MbSue 
         Height          =   345
         Left            =   1740
         TabIndex        =   5
         Top             =   4890
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,##0.#0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MbCarn 
         Height          =   345
         Left            =   1740
         TabIndex        =   4
         Top             =   4350
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo CmbEmp 
         Height          =   360
         Left            =   60
         TabIndex        =   2
         Tag             =   "8"
         Top             =   2940
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "Nom"
         BoundColumn     =   "Id_Emp"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo CmbCar 
         Height          =   360
         Left            =   1740
         TabIndex        =   3
         Tag             =   "8"
         Top             =   3780
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "Descri"
         BoundColumn     =   "ID_Car"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MbDias 
         Height          =   345
         Left            =   1740
         TabIndex        =   6
         Top             =   5940
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MbFecIng 
         Height          =   345
         Left            =   1740
         TabIndex        =   20
         Top             =   6960
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MbNOrd 
         Height          =   345
         Left            =   1740
         TabIndex        =   7
         Top             =   6450
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MbFecC 
         Height          =   345
         Left            =   1740
         TabIndex        =   29
         Top             =   7440
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   405
         Left            =   1710
         TabIndex        =   55
         Top             =   3330
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "&Mostrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FrmBolPag.frx":2446
      End
      Begin MSMask.MaskEdBox MbMonCTS 
         Height          =   345
         Left            =   2340
         TabIndex        =   59
         Top             =   5430
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,##0.#0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MbPor 
         Height          =   345
         Left            =   1740
         TabIndex        =   61
         Top             =   5430
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto CTS  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   4
         Left            =   285
         TabIndex        =   60
         Top             =   5460
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ing. :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   14
         Left            =   105
         TabIndex        =   28
         Top             =   6990
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Cese :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   13
         Left            =   105
         TabIndex        =   27
         Top             =   7470
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dias. Lab.:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   12
         Left            =   120
         TabIndex        =   26
         Top             =   6030
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Orden  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   11
         Left            =   105
         TabIndex        =   25
         Top             =   6480
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carnet AFP :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   10
         Left            =   255
         TabIndex        =   24
         Top             =   4410
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sueldo  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   9
         Left            =   105
         TabIndex        =   23
         Top             =   4920
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empleado  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   8
         Left            =   90
         TabIndex        =   22
         Top             =   2670
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   7
         Left            =   105
         TabIndex        =   21
         Top             =   3840
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   5
         Left            =   2640
         TabIndex        =   19
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   3
         Left            =   2850
         TabIndex        =   17
         Top             =   810
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   2790
         Picture         =   "FrmBolPag.frx":2462
         Top             =   210
         Width           =   2385
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Left            =   2580
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   2835
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   8175
      Left            =   150
      TabIndex        =   8
      Top             =   150
      Width           =   5985
      _Version        =   65536
      _ExtentX        =   10557
      _ExtentY        =   14420
      _StockProps     =   14
      Caption         =   "C O N C E P T O S"
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ShadowStyle     =   1
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   435
         Left            =   60
         ScaleHeight     =   435
         ScaleWidth      =   5805
         TabIndex        =   38
         Top             =   7740
         Width           =   5805
         Begin MSMask.MaskEdBox MBTot_A 
            Height          =   345
            Left            =   4290
            TabIndex        =   39
            Top             =   0
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,##0.#0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBTot_A2 
            Height          =   345
            Left            =   0
            TabIndex        =   66
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,##0.#0"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Aportaciones  : "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   17
            Left            =   1680
            TabIndex        =   40
            Top             =   60
            Width           =   2625
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   60
         ScaleHeight     =   405
         ScaleWidth      =   5775
         TabIndex        =   33
         Top             =   5670
         Width           =   5775
         Begin MSMask.MaskEdBox MBTot_P 
            Height          =   345
            Left            =   4530
            TabIndex        =   34
            Top             =   0
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,##0.#0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBTot_D 
            Height          =   345
            Left            =   1140
            TabIndex        =   35
            Top             =   0
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,##0.#0"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N. Pagar :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   15
            Left            =   3420
            TabIndex        =   37
            Top             =   60
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descto. :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   16
            Left            =   150
            TabIndex        =   36
            Top             =   60
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   5775
         TabIndex        =   30
         Top             =   2670
         Width           =   5775
         Begin MSMask.MaskEdBox MBTot_R 
            Height          =   345
            Left            =   4260
            TabIndex        =   31
            Top             =   0
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,##0.#0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBTot_R2 
            Height          =   345
            Left            =   0
            TabIndex        =   65
            Top             =   0
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,##0.#0"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total de Remuneraciones  :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   6
            Left            =   1290
            TabIndex        =   32
            Top             =   60
            Width           =   2940
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   2235
         Left            =   150
         TabIndex        =   9
         Top             =   360
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   3942
         _Version        =   393216
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid Grid2 
         Height          =   2205
         Left            =   150
         TabIndex        =   10
         Top             =   3360
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   3889
         _Version        =   393216
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid Grid3 
         Height          =   1215
         Left            =   150
         TabIndex        =   11
         Top             =   6390
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   2143
         _Version        =   393216
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MbAño 
         Height          =   345
         Left            =   2010
         TabIndex        =   63
         Top             =   30
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "APORTACIONES DEL EMPLEADOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   14
         Top             =   6180
         Width           =   3510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEDUCCIONES Y DSCTO.  AL TRABAJADOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   3150
         Width           =   4485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REMUNERACIONES"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   150
         Width           =   1995
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "numerocuo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   0
      Left            =   7200
      TabIndex        =   67
      Top             =   -60
      Visible         =   0   'False
      Width           =   1155
   End
End
Attribute VB_Name = "FrmBolPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstBol As ADODB.Recordset
Dim RstEmp As ADODB.Recordset
Dim RstCar As ADODB.Recordset
Dim RstAno As ADODB.Recordset
Dim RstHH As ADODB.Recordset
Dim RstMes1 As ADODB.Recordset
Dim TotQMA  As Currency, TotSuel As Currency
Dim OPQ As Integer, Cod_Emp As Integer
Dim NroBol As Currency, TotDias As Currency, TotMtar As Currency, TotDD As Currency, TotNeto As Currency

Sub Habilita_Boton()
    CmdBuscar.Enabled = True
    CmdBuscar.Enabled = True
    CmdCerrar.Enabled = True
    CmdImp.Enabled = True
    SSFrame1.Enabled = False
End Sub
Sub DesHBot()
    CmdBuscar.Enabled = False
    CmdBuscar.Enabled = False
    CmdCerrar.Enabled = False
    CmdImp.Enabled = False
    SSFrame1.Enabled = True
End Sub
Sub Limpiar()
    TxtNroBol = ""
    'DtpFecha = Date
    CmbEmp.BoundText = ""
    CmbCar.BoundText = ""
    CmbAno.BoundText = ""
    MbCarn = ""
    MbSue = "0.00"
    MbDias = "0"
    MbNOrd = ""
    MbFecIng = ""
    MbFecC = ""
    MBTot_R = "0.00"
    MBTot_P = "0.00"
    MBTot_D = "0.00"
    MBTot_A = "0.00"
    MbMonCTS = "0.00"
    MbPor = "0"
    MbNAsie = ""
    Grid1.Rows = 1
    Grid2.Rows = 1
    Grid3.Rows = 1
End Sub
Private Sub DesEdic()
   CmdNuevo.Enabled = True
   CmdNuevo.Tag = "N"
   CmdNuevo.Caption = "&Nuevo"
   CmdNuevo.Picture = LoadPicture(StrIco1)
   CmdModifi.Enabled = True
   CmdModifi.Tag = "E"
   CmdModifi.Caption = "&Editar"
   CmdModifi.Picture = LoadPicture(StrIco2)
   CmdEliminar.Tag = "A"
   CmdEliminar.Caption = "&Anular"
   CmdEliminar.Picture = LoadPicture(StrIco3)
   CmdNuevo.SetFocus
End Sub
Private Sub ActEdic(CadS As String)
   If CadS = "N" Then
      CmdNuevo.Tag = "G"
      CmdNuevo.Caption = "&Guardar"
      CmdNuevo.Picture = LoadPicture(StrIco4)
      CmdModifi.Enabled = False
   Else
      CmdModifi.Tag = "G"
      CmdModifi.Caption = "&Guardar"
       CmdModifi.Picture = LoadPicture(StrIco4)
      CmdNuevo.Enabled = False
   End If
   CmdEliminar.Tag = "C"
   CmdEliminar.Caption = "&Cancelar"
   CmdEliminar.Picture = LoadPicture(StrIco5)
End Sub

Private Sub CmbAno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CmbMes1.SetFocus
End Sub

Private Sub CmbCar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then MbCarn.SetFocus
End Sub

Private Sub CmbEmp_Click(Area As Integer)
  If CmbEmp.BoundText <> "" Then
        Str_Sql = " SELECT SueldoBasico,Mon_CTS,Por_CTS,Tot_Remu,Sw_Dias,Tot_Dias, ID_Car, Fec_Cese, Fec_Ing, N_Orden, Carnet_AFP FROM TBEmple" & _
                  " Where TBEmple.Id_Emp =" & CmbEmp.BoundText
        Set RstBol = New ADODB.Recordset
        RstBol.Open Str_Sql, BD_Sistema, adOpenDynamic, adLockReadOnly, adCmdText
        With RstBol
           If Not (.EOF() Or .BOF()) Then
                CmbCar.BoundText = IIf(IsNull(!Id_Car), 0, !Id_Car)
                MbCarn = IIf(IsNull(!Carnet_AFP), "", !Carnet_AFP)
                Tot_Dias_Mes CmbMes1.BoundText, CmbAno.Text
                MbDias = gDiasG
                MbSue = !SueldoBasico
                MbMonCTS = !Mon_CTS
                MbPor = !Por_CTS
                MbNOrd = IIf(IsNull(!N_Orden), "", !N_Orden)
                If Not IsNull(!Fec_Cese) Then MbFecC = !Fec_Cese
                If Not IsNull(!Fec_Ing) Then MbFecIng = !Fec_Ing
           End If
           .Close
        End With
        Set RstBol = Nothing
 End If
End Sub

Private Sub CmbEmp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then CmbCar.SetFocus
End Sub

Private Sub CmbMes1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then CmbEmp.SetFocus
End Sub

Private Sub CmbMes1_LostFocus()
  If CmbMes1.BoundText <> "" Then
     Tot_Dias_Mes CmbMes1.BoundText, CmbAno.Text
     DtpFecha = Fecha_G
     MbTipC = B_TpoCmb(DtpFecha)
  End If
End Sub

Private Sub CmdBuscar_Click()
On Error GoTo Errores
    FrmBBpla.Show vbModal
     If Cancelar = 0 Then
        Set RstBol = New ADODB.Recordset
        RstBol.Open "Select * From TBBol_Pago Where No_Bol='" & CodGene & "'", BD_Sistema, adOpenDynamic, adLockReadOnly, adCmdText
        With RstBol
            If Not (.EOF() Or .BOF()) Then
                  TxtNroBol = ![No_Bol]
                  DtpFecha = ![Fecha]
                  CmbAno.BoundText = ![Id_Año]
                  CmbMes1.BoundText = ![Id_Mes]
                  CmbEmp.BoundText = ![Id_Empl]
                  CmbCar.BoundText = ![Id_Car]
                  MbCarn = IIf(IsNull(![No_Carnet]), "", ![No_Carnet])
                  MbSue = ![Sueldo]
                  MbDias = ![Dias_Lab]
                  MbNOrd = IIf(IsNull(![No_Orden]), "", ![No_Orden])
                  If Not IsNull(!Fec_Ing) Then MbFecIng = !Fec_Ing Else MbFecIng = ""
                  If Not IsNull(!Fec_Cese) Then MbFecC = !Fec_Cese Else MbFecC = ""
                  MbPor = !Por_CTS
                  MBTot_R = ![Tot_Remu]
                  MBTot_D = ![Tot_DD]
                  MBTot_P = ![Neto_Pag]
                  MBTot_A = ![Tot_Aport]
                  MbNAsie = IIf(IsNull(![NroAsto]), "", ![NroAsto])
                  CmbMes1.BoundText = Val(Right(![Mes_Ano], 2))
                  MbAño = Left(![Mes_Ano], 2)
                  MbTipC = ![TPOCMB]
                  If !Tip_Pago = 1 Then OptEfe.Value = True Else OptPro.Value = True
                  MbMonCTS = ![Mon_CTS]
                  
            End If
            .Close
        End With
        Set RstBol = Nothing
         Str_Sql = " SELECT TBD_BRem.Num_Rem AS Cod, TB_REMU.Des_Rem AS DescV, TBD_BRem.Monto AS Mon,0 AS Dscto, TB_REMU.ID_PCta, '0' AS ID_PCta2, TB_REMU.Sw_Aplic FROM TBD_BRem INNER JOIN TB_REMU ON TBD_BRem.Num_Rem = TB_REMU.Num_Rem" & _
                   " Where TBD_BRem.No_Bol='" & CodGene & "' ORDER BY TBD_BRem.Numero;"
         AddIt Grid1, 0
         
         Str_Sql = " SELECT TBD_BDDT.Num_DDT AS Cod, TB_DDT.Des_DDT AS DescV, TBD_BDDT.Monto AS Mon, TB_DDT.ID_PCta, '0' AS ID_PCta2, 0 AS Aport, TBD_BDDT.DDT_Porc AS [Dscto], TBD_BDDT.Sw_Aplic FROM TB_DDT INNER JOIN TBD_BDDT ON TB_DDT.Num_DDT = TBD_BDDT.Num_DDT" & _
                   " Where TBD_BDDT.No_Bol='" & CodGene & "' ORDER BY TBD_BDDT.Numero;"
         AddIt Grid2, 0

         Str_Sql = " SELECT TBD_BAE.Num_AE AS Cod, TB_AEMP.Des_AE AS DescV, TBD_BAE.Monto AS Mon, TB_AEMP.ID_PCta, TB_AEMP.ID_PCta1 AS ID_PCta2, TBD_BAE.Ae_Porc AS Dscto, TBD_BAE.Sw_Aplic FROM TB_AEMP INNER JOIN TBD_BAE ON TB_AEMP.Num_AE = TBD_BAE.Num_AE" & _
                   " Where TBD_BAE.No_Bol='" & CodGene & "' ORDER BY TBD_BAE.Numero;"
         AddIt Grid3, 0
      End If
Exit Sub
Errores:
  If Ctrl_Error(Err, RstBol) Then Exit Sub
End Sub

Private Sub CmdCerrar_Click()
  Unload Me
End Sub


Private Sub CmdEliminar_Click()
On Error GoTo Errores
    If CmdEliminar.Tag = "A" Then
       If Val(TxtNroBol) <> 0 Then
          If MsgBox("Estas seguro de Eliminar los Datos?...", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
             FrmClaAut.Show vbModal
             If Cancelar = 0 Then
                If (SwNAcc = 1 Or SwNAcc = 2) Then
                   BD_Sistema.Execute "Update TBBol_Pago set Anular=1,FecAnu=Date() Where No_Bol='" & TxtNroBol & "'"
                   'Eliminando los registro de contables
                   Limpiar
                   Call Mensaje(1)
                Else
                     MsgBox " Ud. no tiene autorización . . ." & Chr(13) & " Por favor, Consulte con Administración...", vbExclamation, "Aviso..."
                End If ' Final de Clave de autorizaciòn
             End If
          End If
       Else
             MsgBox "Datos Vacío ", vbExclamation, "Eliminar"
       End If
    Else
           DesEdic
           Habilita_Boton
           Limpiar
    End If
Exit Sub
Errores:
  If Ctrl_Error(Err, RstBol) Then
     Exit Sub
  End If

End Sub
Sub Font_Ne()
    With Printer
         .FontBold = True
         .FontItalic = False
         .FontSize = 10
         .FontName = "Arial"
    End With
End Sub
Sub Font_No()
    With Printer
         .FontBold = False
         .FontItalic = False
         .FontSize = 10
         .FontName = "Arial"
    End With
End Sub
Sub Cab_Izq()
    'Printer.FontBold = True
    'Printer.FontItalic = True
    'Printer.FontName = "Time New Roman"
    Font_Ne
    Printer.FontSize = 14
    Cx1 = Printer.ScaleWidth / 25
    Cy1 = Printer.ScaleHeight / 23
    Printer.PSet (Cx1, Cy1)
    Printer.Print "VANECO E.I.R.LTDA."
    'Cy1 = Printer.ScaleHeight / 17
    'Cx1 = Printer.ScaleWidth / 5.6
    'Printer.PSet (Cx1, Cy1)
    'Printer.Print "S.A."
    Cx1 = Printer.ScaleWidth / 3
    Cy1 = Printer.ScaleHeight / 26
    Printer.PSet (Cx1, Cy1)
    Printer.Print "B O L E T A"
    Cx1 = Printer.ScaleWidth / 2.9
    Cy1 = Printer.ScaleHeight / 16
    Printer.PSet (Cx1, Cy1)
    Printer.Print "Nº " & TxtNroBol
    Font_No
    Cx1 = Printer.ScaleWidth / 3.5
    Cy1 = Printer.ScaleHeight / 11.5
    Printer.PSet (Cx1, Cy1)
    Printer.Print " D.S.015-72-TR    28 - 09-72"
    
    Cx1 = Printer.ScaleWidth / 30
    Cy1 = Printer.ScaleHeight / 9.2
    Printer.PSet (Cx1, Cy1)
    Printer.Print "Reg. Pat. Seg. Soc. Del Perú N° 101195290000000"
    
    Cy1 = Printer.ScaleHeight / 7.7
    Printer.PSet (Cx1, Cy1)
    Printer.Print "Nombres :  " & CmbEmp.Text
    
    Cy2 = Printer.ScaleHeight / 6.6
    Printer.PSet (Cx1, Cy2)
    Printer.Print "Cargo :  " & CmbCar.Text
    
    Cy3 = Printer.ScaleHeight / 5.8
    Printer.PSet (Cx1, Cy3)
    Printer.Print "Sueldo :  " & Format(MbSue, "#,##0.#0")
    
    Cy4 = Printer.ScaleHeight / 5.2
    Printer.PSet (Cx1, Cy4)
    Printer.Print "Días Lab. :  " & MbDias
    
    Cx1 = Printer.ScaleWidth / 4
    Printer.PSet (Cx1, Cy2)
    Printer.Print "Carnet AFP   : " & MbCarn
    Printer.PSet (Cx1, Cy3)
    Printer.Print "Fecha Ing. : " & MbFecIng
    
    Printer.PSet (Cx1, Cy4)
    Printer.Print "Fecha Cese : " & MbFecC
    
'    Cx1 = Printer.ScaleWidth / 40
'    Cy1 = Printer.ScaleHeight / 5.5
'    Printer.Line (Cx1, Cy1)-(Cx6, Cy1)
    
    Cx1 = Printer.ScaleWidth / 16
    Cy1 = Printer.ScaleHeight / 4.45
    Printer.PSet (Cx1, Cy1)
    Printer.Print "SUELDO CORRESPONDIENTE AL MES DE " & UCase(Format(DtpFecha, "mmmm")) & " " & Format(DtpFecha, "yyyy")
End Sub


Private Sub CmdImp_Click()
   If Val(TxtNroBol) <> 0 Then
        Sw_Estado = False
        xCol = 1
        SwIgv = 0
        Cancelar = 1
        FrmPrint.Show vbModal
        If Cancelar = 0 Then
           If Sw_Estado = False Then
                 Printer.PaperSize = vbPRPSA4
                 Printer.Orientation = 2
                 Printer.ScaleMode = 3
                 Printer.DrawWidth = 1   ' Establece DrawWidth.
                 Font_No
                 Cx1 = Printer.ScaleWidth / 40
                 Cy1 = Printer.ScaleHeight / 30
                 Cx6 = Printer.ScaleWidth / 2.2
                 Printer.Line (Cx1, Cy1)-(Cx6, Cy1)
                 Cy1 = Printer.ScaleHeight / 1.08
                 Printer.Line (Cx1, Cy1)-(Cx6, Cy1)
                 Cy1 = Printer.ScaleHeight / 30
                 Cy2 = Printer.ScaleHeight / 1.08
                 Printer.Line (Cx1, Cy1)-(Cx1, Cy2)
                 Cx1 = Printer.ScaleWidth / 2.2
                 Printer.Line (Cx1, Cy1)-(Cx6, Cy2)
                 ' Cuadro Derecha
                 Cx1 = Printer.ScaleWidth / 40
                 Cy1 = Printer.ScaleHeight / 4.6
                 Printer.Line (Cx1, Cy1)-(Cx6, Cy1)
                 Cy1 = Printer.ScaleHeight / 4
                 Printer.Line (Cx1, Cy1)-(Cx6, Cy1)
                 
                 Cy1 = Printer.ScaleHeight / 30
                 Cx1 = Printer.ScaleWidth / 2
                 Cx6 = Printer.ScaleWidth / 1.08
                 Printer.Line (Cx1, Cy1)-(Cx6, Cy1) ' Linea Superior
                 Cy1 = Printer.ScaleHeight / 1.08
                 Printer.Line (Cx1, Cy1)-(Cx6, Cy1) 'Linea Inferior
                 Cy1 = Printer.ScaleHeight / 30
                 Printer.Line (Cx1, Cy1)-(Cx1, Cy2) ' Linea Izquierda
                 Cx1 = Printer.ScaleWidth / 1.08
                 Printer.Line (Cx1, Cy1)-(Cx6, Cy2) ' Linea derecha
                 'Cuadro Izquierda
                 Cx1 = Printer.ScaleWidth / 2
                 Cy1 = Printer.ScaleHeight / 4.6
                 Printer.Line (Cx1, Cy1)-(Cx6, Cy1)
                 Cy1 = Printer.ScaleHeight / 4
                 Printer.Line (Cx1, Cy1)-(Cx6, Cy1)
                 '(Cx1 = Punto Inicial
                 'Cy1) = Punto de Partida hacia Horizonatal
                 '(Cx6 = Punto Final en Vertical
                 'Cy2) = Tamaño de Linea Horizontal hacia Abajo
                 Hoj_Izq
                 Hoj_Der
                 Printer.EndDoc
           End If
       End If
 Else
    MsgBox "No hay información que Imprimir", vbExclamation, "Aviso"
 End If
End Sub
Public Sub Hoj_Izq()
    Cab_Izq
    Cx1 = Printer.ScaleWidth / 30
    Cy2 = Printer.ScaleHeight / 3.8
    StrEspac = Printer.ScaleHeight / 60

    Printer.PSet (Cx1, Cy2)
    Printer.Print "REMUNERACIONES"
    
    Cy2 = Cy2 + (StrEspac * 2)
    Cx2 = Printer.ScaleWidth / 3
    Cx3 = Printer.ScaleWidth / 2.35
    For I = 1 To Grid1.Rows - 1
        Printer.PSet (Cx1, Cy2)
        Printer.Print Grid1.TextMatrix(I, 1)

        TotGen = Cx3 - Printer.TextWidth(Format(Grid1.TextMatrix(I, 2), "##,##0.#0"))
        Printer.PSet (TotGen, Cy2)
        Printer.Print Format(Grid1.TextMatrix(I, 2), "##,##0.#0")

        Cy2 = Cy2 + StrEspac
    Next
    Printer.PSet (Cx1, Cy2)
    Printer.Print "TOTAL DE REMUNERACIONES  "

    TotGen = Cx3 - Printer.TextWidth(Format(MBTot_R, "##,##0.#0"))
    Printer.PSet (TotGen, Cy2)
    Printer.Print Format(MBTot_R, "##,##0.#0")

    Cy2 = Cy2 + (StrEspac * 2)
    Printer.PSet (Cx1, Cy2)
    Printer.Print "DEDUCCIONES Y DESCUENTOS AL TRABAJADOR"
    Cy2 = Cy2 + (StrEspac * 2)
    For I = 1 To Grid2.Rows - 1
        Printer.PSet (Cx1, Cy2)
        Printer.Print Grid2.TextMatrix(I, 1)

        TotGen = Cx2 - Printer.TextWidth(Format(Grid2.TextMatrix(I, 2), "##,##0.#0"))
        Printer.PSet (TotGen, Cy2)
        Printer.Print Format(Grid2.TextMatrix(I, 2), "##,##0.#0")
        Cy2 = Cy2 + StrEspac
    Next
    Printer.PSet (Cx1, Cy2)
    Printer.Print "TOTAL DESCUENTOS Y DEDUCCIONES"

    TotGen = Cx3 - Printer.TextWidth(Format(MBTot_D, "##,##0.#0"))
    Printer.PSet (TotGen, Cy2)
    Printer.Print Format(MBTot_D, "##,##0.#0")

    Cy2 = Cy2 + StrEspac
    Printer.PSet (Cx1, Cy2)
    Printer.Print "NETO A PAGAR  : "

    TotGen = Cx3 - Printer.TextWidth(Format(MBTot_P, "##,##0.#0"))
    Printer.PSet (TotGen, Cy2)
    Printer.Print Format(MBTot_P, "##,##0.#0")

    Cy2 = Cy2 + (StrEspac * 2)

    Printer.PSet (Cx1, Cy2)
    Printer.Print "APORTACIONES DEL EMPLEADOR"
    
    Cy2 = Cy2 + (StrEspac * 2)
    For I = 1 To Grid3.Rows - 1
        Printer.PSet (Cx1, Cy2)
        Printer.Print Grid3.TextMatrix(I, 1)

        TotGen = Cx2 - Printer.TextWidth(Format(Grid3.TextMatrix(I, 2), "##,##0.#0"))
        Printer.PSet (TotGen, Cy2)
        Printer.Print Format(Grid3.TextMatrix(I, 2), "##,##0.#0")
        Cy2 = Cy2 + StrEspac
    Next
    Printer.PSet (Cx1, Cy2)
    Printer.Print "TOTAL DE APORTACIONES "

    TotGen = Cx3 - Printer.TextWidth(Format(MBTot_A, "##,##0.#0"))
    Printer.PSet (TotGen, Cy2)
    Printer.Print Format(MBTot_A, "##,##0.#0")
    Cy2 = Cy2 + (StrEspac * 4)
    Cx1 = Printer.ScaleWidth / 30
    Printer.PSet (Cx1, Cy2)
    Printer.Print "VANECO E.I.R.LTDA. "

    Cx1 = Printer.ScaleWidth / 5
    Cx6 = Printer.ScaleWidth / 2.4
    Printer.Line (Cx1, Cy2)-(Cx6, Cy2)
    
    Cx1 = Printer.ScaleWidth / 5
    Cy2 = Cy2 + (StrEspac / 2)
    Printer.PSet (Cx1, Cy2)
    Printer.Print "Recibí Conforme"
    Cy2 = Cy2 + StrEspac
    Printer.PSet (Cx1, Cy2)
    Printer.Print CmbEmp.Text

End Sub
Sub Cab_Der()
    Font_Ne
    Printer.FontSize = 14
    Cx1 = Printer.ScaleWidth / 1.8
    Cy1 = Printer.ScaleHeight / 23
    Printer.PSet (Cx1, Cy1)
    Printer.Print "VANECO E.I.R.LTDA. "
    'Cy1 = Printer.ScaleHeight / 17
    'Cx1 = Printer.ScaleWidth / 1.38
    'Printer.PSet (Cx1, Cy1)
    'Printer.Print "S.A."
    Cx1 = Printer.ScaleWidth / 1.2
    Cy1 = Printer.ScaleHeight / 26
    Printer.PSet (Cx1, Cy1)
    Printer.Print "B O L E T A"
    Cy1 = Printer.ScaleHeight / 16
    Cx1 = Printer.ScaleWidth / 1.19
    Printer.PSet (Cx1, Cy1)
    Printer.Print "Nº " & TxtNroBol

    Font_No
    Cx1 = Printer.ScaleWidth / 1.35
    Cy1 = Printer.ScaleHeight / 11.5
    Printer.PSet (Cx1, Cy1)
    Printer.Print " D.S. N° 001-98-TR"
    
    Cx1 = Printer.ScaleWidth / 1.95
    Cy1 = Printer.ScaleHeight / 9.2
    Printer.PSet (Cx1, Cy1)
    Printer.Print "Reg. Pat. Seg. Soc. Del Perú N° 122913880000000"
    
    Cy1 = Printer.ScaleHeight / 7.7
    Printer.PSet (Cx1, Cy1)
    Printer.Print "Nombres :  " & CmbEmp.Text
    
    Cy2 = Printer.ScaleHeight / 6.6
    Printer.PSet (Cx1, Cy2)
    Printer.Print "Cargo :  " & CmbCar.Text
    
    Cy3 = Printer.ScaleHeight / 5.8
    Printer.PSet (Cx1, Cy3)
    Printer.Print "Sueldo :  " & Format(MbSue, "#,##0.#0")
    
    Cy4 = Printer.ScaleHeight / 5.2
    Printer.PSet (Cx1, Cy4)
    Printer.Print "Días Lab. :  " & MbDias
    
    Cx1 = Printer.ScaleWidth / 1.38
    Printer.PSet (Cx1, Cy2)
    Printer.Print "Carnet AFP  : " & MbCarn
    Printer.PSet (Cx1, Cy3)
    Printer.Print "Fecha Ing. : " & MbFecIng
    
    Printer.PSet (Cx1, Cy4)
    Printer.Print "Fecha Cese : " & MbFecC
    
    
    Cx1 = Printer.ScaleWidth / 1.95
    Cy1 = Printer.ScaleHeight / 4.45
    Printer.PSet (Cx1, Cy1)
    Printer.Print "SUELDO CORRESPONDIENTE AL MES DE " & UCase(Format(DtpFecha, "mmmm")) & " " & Format(DtpFecha, "yyyy")
    
End Sub

Public Sub Hoj_Der()
    Cab_Der

    Cx1 = Printer.ScaleWidth / 1.95
    Cy2 = Printer.ScaleHeight / 3.8
    StrEspac = Printer.ScaleHeight / 60

    Printer.PSet (Cx1, Cy2)
    Printer.Print "REMUNERACIONES"
    
    Cy2 = Cy2 + (StrEspac * 2)
    Cx2 = Printer.ScaleWidth / 1.25
    Cx3 = Printer.ScaleWidth / 1.1
    For I = 1 To Grid1.Rows - 1
        Printer.PSet (Cx1, Cy2)
        Printer.Print Grid1.TextMatrix(I, 1)

        TotGen = Cx3 - Printer.TextWidth(Format(Grid1.TextMatrix(I, 2), "##,##0.#0"))
        Printer.PSet (TotGen, Cy2)
        Printer.Print Format(Grid1.TextMatrix(I, 2), "##,##0.#0")

        Cy2 = Cy2 + StrEspac
    Next
    Printer.PSet (Cx1, Cy2)
    Printer.Print "TOTAL DE REMUNERACIONES  "

    TotGen = Cx3 - Printer.TextWidth(Format(MBTot_R, "##,##0.#0"))
    Printer.PSet (TotGen, Cy2)
    Printer.Print Format(MBTot_R, "##,##0.#0")

    Cy2 = Cy2 + (StrEspac * 2)
    
    Printer.PSet (Cx1, Cy2)
    Printer.Print "DEDUCCIONES Y DESCUENTOS AL TRABAJADOR"
    
    Cy2 = Cy2 + (StrEspac * 2)
    For I = 1 To Grid2.Rows - 1
        Printer.PSet (Cx1, Cy2)
        Printer.Print Grid2.TextMatrix(I, 1)

        TotGen = Cx2 - Printer.TextWidth(Format(Grid2.TextMatrix(I, 2), "##,##0.#0"))
        Printer.PSet (TotGen, Cy2)
        Printer.Print Format(Grid2.TextMatrix(I, 2), "##,##0.#0")
        Cy2 = Cy2 + StrEspac
    Next
    Printer.PSet (Cx1, Cy2)
    Printer.Print "TOTAL DESCUENTOS Y DEDUCCIONES"

    TotGen = Cx3 - Printer.TextWidth(Format(MBTot_D, "##,##0.#0"))
    Printer.PSet (TotGen, Cy2)
    Printer.Print Format(MBTot_D, "##,##0.#0")

    Cy2 = Cy2 + StrEspac
    Printer.PSet (Cx1, Cy2)
    Printer.Print "NETO A PAGAR  : "

    TotGen = Cx3 - Printer.TextWidth(Format(MBTot_P, "##,##0.#0"))
    Printer.PSet (TotGen, Cy2)
    Printer.Print Format(MBTot_P, "##,##0.#0")

    Cy2 = Cy2 + (StrEspac * 2)

    
    Printer.PSet (Cx1, Cy2)
    Printer.Print "APORTACIONES DEL EMPLEADOR"
    
    Cy2 = Cy2 + (StrEspac * 2)
    For I = 1 To Grid3.Rows - 1
        Printer.PSet (Cx1, Cy2)
        Printer.Print Grid3.TextMatrix(I, 1)

        TotGen = Cx2 - Printer.TextWidth(Format(Grid3.TextMatrix(I, 2), "##,##0.#0"))
        Printer.PSet (TotGen, Cy2)
        Printer.Print Format(Grid3.TextMatrix(I, 2), "##,##0.#0")
        Cy2 = Cy2 + StrEspac
    Next
    Printer.PSet (Cx1, Cy2)
    Printer.Print "TOTAL DE APORTACIONES "

    TotGen = Cx3 - Printer.TextWidth(Format(MBTot_A, "##,##0.#0"))
    Printer.PSet (TotGen, Cy2)
    Printer.Print Format(MBTot_A, "##,##0.#0")
    Cy2 = Cy2 + (StrEspac * 4)
    Cx1 = Printer.ScaleWidth / 1.95
    Printer.PSet (Cx1, Cy2)
    Printer.Print "VANECO E.I.R.LTDA."
    
    Cx1 = Printer.ScaleWidth / 1.5
    Cx6 = Printer.ScaleWidth / 1.1
    Printer.Line (Cx1, Cy2)-(Cx6, Cy2)
    
    Cx1 = Printer.ScaleWidth / 1.35
    Cy2 = Cy2 + (StrEspac / 2)
    Cx1 = Printer.ScaleWidth / 1.5
    Printer.PSet (Cx1, Cy2)
    Printer.Print "Recibí Conforme"
    Cy2 = Cy2 + StrEspac
    Printer.PSet (Cx1, Cy2)
    Printer.Print CmbEmp.Text

End Sub

Private Sub CmdModifi_Click()
On Error GoTo Errores
  Dim TotIpSS, TotAFP As Currency
  If CmdModifi.Tag = "E" Then
         If Val(TxtNroBol) <> 0 Then
             ActEdic ("G")
             DesHBot
             CmbMes1.SetFocus
             CmdEliminar.Enabled = True
             CmbEmp.Enabled = False
             'SSCommand1.Enabled = False
         Else
             MsgBox "Muestrelo los datos para Modificar... ", vbInformation, "Modificar"
         End If
      Else
        If Val_Año <> CmbAno.BoundText Then
           MsgBox "Año de Proceso no es Correcto !!!!!!", vbExclamation, "Aviso"
           Exit Sub
        End If
        Est_MesC CmbMes1.BoundText, Val_Año
        If Eval = 1 Then Exit Sub
        If Trim(MbFecC) <> "" Then
            If Not IsDate(MbFecC) Then
               MsgBox "Fecha Incorrecta ...", vbExclamation, "Aviso"
               MbFecC = ""
               Exit Sub
            End If
        End If
         If Val(TxtNroBol) <> 0 And CmbEmp.BoundText <> "" And CmbMes1.BoundText <> "" And Grid1.Rows >= 2 And Grid2.Rows >= 2 And Grid3.Rows >= 2 And CmbAno.BoundText <> "" Then
             If MsgBox("Estas seguro modificar los datos ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                StrCad = MbAño & Right("0" & CmbMes1.BoundText, 2)
                Val_Mes = CmbMes1.BoundText
                TotIpSS = 0
                TotAFP = 0
                If Grid2.Rows >= 2 Then
                   For I = 1 To Grid2.Rows - 1 'IPSS
                       If Val(Grid2.TextMatrix(I, 0)) = 1 Then
                          If Val(Grid2.TextMatrix(I, 2)) <> 0 Then
                             TotIpSS = CDbl(Grid2.TextMatrix(I, 2))
                          End If
                          Exit For
                       End If
                   Next
                   For I = 1 To Grid2.Rows - 1 'AFP
                       If (Val(Grid2.TextMatrix(I, 0)) = 3) Or (Val(Grid2.TextMatrix(I, 0)) = 4) Or (Val(Grid2.TextMatrix(I, 0)) = 5) Then
                          If Val(Grid2.TextMatrix(I, 2)) <> 0 Then
                             TotAFP = TotAFP + CDbl(Grid2.TextMatrix(I, 2))
                          End If
                       End If
                   Next
                End If
                BD_Sistema.Execute "Delete No_Bol From TBBol_Pago Where No_Bol='" & TxtNroBol & "'"
                If IsDate(MbFecC) Then
                    Str_Sql = "Insert Into TBBol_Pago (No_Bol,Id_Empl,Id_Car,Dias_Lab,Tot_Remu,Tot_DD,Neto_Pag,Tot_Aport,Sueldo,Fecha,Fec_Ing,Fec_Cese,No_Orden,No_Carnet,ID_Mes,Id_Año," & IIf(OptEfe.Value, "", "Saldo,") & " NroAsto,Mes_Ano,TPOCMB,Mon_CTS,Sal_CTS,Tip_Pago,Mon_Ipss,Sal_Ipss,Mon_Afp,Sal_AFP,Mon_Seg,Sal_Seg,Por_CTS) Values('" & TxtNroBol & "'," & _
                    CmbEmp.BoundText & "," & CmbCar.BoundText & "," & Val(MbDias) & "," & CDbl(MBTot_R) & "," & CDbl(MBTot_D) & "," & CDbl(MBTot_P) & "," & CDbl(MBTot_A) & "," & CDbl(MbSue) & ",#" & Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(MbFecIng, "mm/dd/yyyy") & _
                    "#,#" & Format(MbFecC, "mm/dd/yyyy") & "#,'" & IIf(Trim(MbNOrd) = "", Chr(32), Trim(MbNOrd)) & "','" & IIf(Trim(MbCarn) = "", Chr(32), Trim(MbCarn)) & "'," & CmbMes1.BoundText & "," & CmbAno.BoundText & "," & IIf(OptEfe.Value, "", CDbl(MBTot_P) & ",") & "'" & MbNAsie & "','" & StrCad & "'," & Val(MbTipC) & "," & CDbl(MbMonCTS) & "," & CDbl(MbMonCTS) & "," & IIf(OptEfe.Value, 1, 2) & _
                    "," & CDbl(TotIpSS) & "," & CDbl(TotIpSS) & "," & CDbl(TotAFP) & "," & CDbl(TotAFP) & "," & CDbl(MBTot_A) & "," & CDbl(MBTot_A) & "," & Val(MbPor) & ")"
                Else
                    Str_Sql = "Insert Into TBBol_Pago (No_Bol,Id_Empl,Id_Car,Dias_Lab,Tot_Remu,Tot_DD,Neto_Pag,Tot_Aport,Sueldo,Fecha,Fec_Ing,No_Orden,No_Carnet,ID_Mes,Id_Año," & IIf(OptEfe.Value, "", "Saldo,") & " NroAsto,Mes_Ano,TPOCMB,Mon_CTS,Sal_CTS,Tip_Pago,Mon_Ipss,Sal_Ipss,Mon_Afp,Sal_AFP,Mon_Seg,Sal_Seg,Por_CTS) Values('" & TxtNroBol & "'," & _
                    CmbEmp.BoundText & "," & CmbCar.BoundText & "," & Val(MbDias) & "," & CDbl(MBTot_R) & "," & CDbl(MBTot_D) & "," & CDbl(MBTot_P) & "," & CDbl(MBTot_A) & "," & CDbl(MbSue) & ",#" & Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(MbFecIng, "mm/dd/yyyy") & _
                    "#,'" & IIf(Trim(MbNOrd) = "", Chr(32), Trim(MbNOrd)) & "','" & IIf(Trim(MbCarn) = "", Chr(32), Trim(MbCarn)) & "'," & CmbMes1.BoundText & "," & CmbAno.BoundText & "," & IIf(OptEfe.Value, "", CDbl(MBTot_P) & ",") & "'" & MbNAsie & "','" & StrCad & "'," & Val(MbTipC) & "," & CDbl(MbMonCTS) & "," & CDbl(MbMonCTS) & "," & IIf(OptEfe.Value, 1, 2) & _
                    "," & CDbl(TotIpSS) & "," & CDbl(TotIpSS) & "," & CDbl(TotAFP) & "," & CDbl(TotAFP) & "," & CDbl(MBTot_A) & "," & CDbl(MBTot_A) & "," & Val(MbPor) & ")"
                End If
                
                BD_Sistema.Execute Str_Sql
                BD_Sistema.Execute "Delete No_Bol From TBD_BRem  Where No_Bol='" & TxtNroBol & "'"
                For I = 1 To Grid1.Rows - 1 ' Remuneraciones
                   BD_Sistema.Execute "Insert Into TBD_BRem (No_Bol,Num_Rem,Monto,Rem_Porc,Sw_Aplic,ID_PCta)Values ('" & TxtNroBol & "'," & _
                   Val(Grid1.TextMatrix(I, 0)) & "," & CDbl(Grid1.TextMatrix(I, 2)) & "," & Val(Grid1.TextMatrix(I, 4)) & "," & Val(Grid1.TextMatrix(I, 5)) & ",'" & IIf(Trim(Grid1.TextMatrix(I, 6)) = "", Chr(32), Trim(Grid1.TextMatrix(I, 6))) & "')"
                Next
                
                BD_Sistema.Execute "Delete No_Bol From TBD_BDDT  Where No_Bol='" & TxtNroBol & "'"
                For I = 1 To Grid2.Rows - 1 ' Descuento y Reducciones del trabajador
                   BD_Sistema.Execute "Insert Into TBD_BDDT (No_Bol,Num_DDT,Monto,DDt_Porc,Sw_Aplic,ID_PCta,Saldo)Values ('" & TxtNroBol & _
                   "'," & Val(Grid2.TextMatrix(I, 0)) & "," & CDbl(Grid2.TextMatrix(I, 2)) & "," & Val(Grid2.TextMatrix(I, 4)) & _
                   "," & Val(Grid2.TextMatrix(I, 5)) & ",'" & IIf(Trim(Grid2.TextMatrix(I, 6)) = "", Chr(32), Trim(Grid2.TextMatrix(I, 6))) & "'," & CDbl(Grid2.TextMatrix(I, 2)) & ")"
                Next
                
                BD_Sistema.Execute "Delete No_Bol From TBD_BAE  Where No_Bol='" & TxtNroBol & "'"
                For I = 1 To Grid3.Rows - 1 ' Aportaciones del Empleador
                   BD_Sistema.Execute "Insert Into TBD_BAE (No_Bol,Num_AE,Monto,Saldo,AE_Porc,Sw_Aplic,ID_PCta,ID_PCta1)Values ('" & TxtNroBol & _
                   "'," & Val(Grid3.TextMatrix(I, 0)) & "," & CDbl(Grid3.TextMatrix(I, 2)) & "," & CDbl(Grid3.TextMatrix(I, 2)) & "," & Val(Grid3.TextMatrix(I, 4)) & "," & Val(Grid3.TextMatrix(I, 5)) & ",'" & _
                   IIf(Trim(Grid3.TextMatrix(I, 6)) = "", Chr(32), Trim(Grid3.TextMatrix(I, 6))) & "','" & IIf(Trim(Grid3.TextMatrix(I, 7)) = "", Chr(32), Trim(Grid3.TextMatrix(I, 7))) & "')"
                Next
                
                DesEdic
                Habilita_Boton
                Call Mensaje(6)
                
            End If
         Else
                If Val(TxtNroBol) = 0 Then MsgBox "Falta Nro. de Boleta de Pago....", vbExclamation, "Error": Exit Sub
                If CmbEmp.BoundText = "" Then MsgBox "Seleccione Empleado , por favor... ", vbExclamation, "Error": CmbEmp.SetFocus: Exit Sub
                If CmbMes1.BoundText = "" Then MsgBox "Seleccione La Mes , por favor... ", vbExclamation, "Error": CmbMes1.SetFocus: Exit Sub
                If CmbAno.BoundText = "" Then MsgBox "Seleccione La Año, por favor... ", vbExclamation, "Error": CmbAno.SetFocus: Exit Sub
                If Grid1.Rows = 1 Then MsgBox "Falta los Coceptos de  REMUNERACIONES ... ", vbExclamation, "Error": Exit Sub
                If Grid2.Rows = 1 Then MsgBox "Falta los Coceptos de REDUCCIONES Y DESCUENTOS AL TRBAJADOR... ", vbExclamation, "Error": Exit Sub
                If Grid3.Rows = 1 Then MsgBox "Falta los Coceptos de  APORTACIONES DEL EMPLEADOR... ", vbExclamation, "Error": Exit Sub
         End If
   End If
Exit Sub
Errores:
   If Ctrl_Error(Err, RstBol) Then Exit Sub
End Sub

Private Sub CmdNuevo_Click()
Dim TotIpSS, TotAFP As Currency
On Error GoTo Errores
 If CmdNuevo.Tag = "N" Then
        ActEdic ("N")
        DesHBot
        Limpiar
        
        Set RstBol = New ADODB.Recordset
        RstBol.Open "Select * From TBCorrela Where Numero=33;", BD_Sistema, adOpenDynamic, adLockReadOnly, adCmdText
        TxtNroBol = Right("0000" & (RstBol!NUME) + 1, 5)
        RstBol.Close
        Set RstBol = Nothing
        CmbMes1.BoundText = Month(DtpFecha)
        CmbAno.BoundText = Val_Año
        CmbEmp.Enabled = True
        SSCommand1.Enabled = True
        CmbAno.SetFocus
 ElseIf CmdNuevo.Tag = "G" Then
    If Val_Año <> CmbAno.BoundText Then
       MsgBox "Año de Proceso no es Correcto !!!!!!", vbExclamation, "Aviso"
       Exit Sub
    End If
    Est_MesC CmbMes1.BoundText, Val_Año
    If Eval = 1 Then Exit Sub
    If Trim(MbFecC) <> "" Then
        If Not IsDate(MbFecC) Then
           MsgBox "Fecha Incorrecta ...", vbExclamation, "Aviso"
           MbFecC = ""
           Exit Sub
        End If
    End If
    If Val(TxtNroBol) <> 0 And CmbEmp.BoundText <> "" And CmbMes1.BoundText <> "" And Grid1.Rows >= 2 And Grid2.Rows >= 2 And Grid3.Rows >= 2 And CmbAno.BoundText <> "" Then
        If MsgBox(" Estas Seguro registrar BOLETA DE PAGO?", vbQuestion + vbYesNo, "Guardar") = vbYes Then
        
            Set RstBol = New ADODB.Recordset
            RstBol.Open "Select Tipo From TBBol_Pago Where Id_Mes=" & Val(CmbMes1.BoundText) & " And id_Año=" & CmbAno.Text & " And Anular =0 And  Id_Empl=" & CmbEmp.BoundText, BD_Sistema, adOpenDynamic, adLockReadOnly, adCmdText
            If Not (RstBol.EOF() Or RstBol.BOF()) Then
               Str_Sql = "BOLETA DE PAGO "
               MsgBox Str_Sql & " Ya está Registrado, Por favor verifique ", vbExclamation, "Aviso"
               RstBol.Close
               Set RstBol = Nothing
               Exit Sub
               CmbEmp.SetFocus
            End If
            RstBol.Close
            Set RstBol = Nothing
            
            StrCad = Right(Val_Año, 2) & Right("0" & CmbMes1.BoundText, 2)
            MbNAsie = StrCad & Condom(CmbMes1.BoundText, 33)
            TotIpSS = 0
            TotAFP = 0
            If Grid2.Rows >= 2 Then
               For I = 1 To Grid2.Rows - 1 'IPSS
                   If Val(Grid2.TextMatrix(I, 0)) = 1 Then
                      If Val(Grid2.TextMatrix(I, 2)) <> 0 Then
                         TotIpSS = CDbl(Grid2.TextMatrix(I, 2))
                      End If
                      Exit For
                   End If
               Next
               For I = 1 To Grid2.Rows - 1 'AFP
                   If (Val(Grid2.TextMatrix(I, 0)) = 3) Or (Val(Grid2.TextMatrix(I, 0)) = 4) Or (Val(Grid2.TextMatrix(I, 0)) = 5) Then
                      If Val(Grid2.TextMatrix(I, 2)) <> 0 Then
                         TotAFP = TotAFP + CDbl(Grid2.TextMatrix(I, 2))
                      End If
                   End If
               Next
            End If
            Str_Sql = "Insert Into TBBol_Pago (No_Bol,Id_Empl,Id_Car,Dias_Lab,Tot_Remu,Tot_DD,Neto_Pag,Tot_Aport,Sueldo,Fecha,Fec_Ing" & IIf(IsDate(MbFecC), ",Fec_Cese", "") & ",No_Orden,No_Carnet,ID_Mes,Id_Año," & IIf(OptEfe.Value, "", "Saldo,") & " NroAsto,Mes_Ano,TPOCMB,Mon_CTS,Sal_CTS,Tip_Pago,Mon_Ipss,Sal_Ipss,Mon_Afp,Sal_AFP,Mon_Seg,Sal_Seg,Por_CTS) Values('" & TxtNroBol & "'," & _
            CmbEmp.BoundText & "," & CmbCar.BoundText & "," & Val(MbDias) & "," & CDbl(MBTot_R) & "," & CDbl(MBTot_D) & "," & CDbl(MBTot_P) & "," & CDbl(MBTot_A) & "," & CDbl(MbSue) & ",#" & Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(MbFecIng, "mm/dd/yyyy") & _
            "#" & IIf(IsDate(MbFecC), ",#" & Format(MbFecC, "mm/dd/yyyy") & "#", "") & ",'" & IIf(Trim(MbNOrd) = "", Chr(32), Trim(MbNOrd)) & "','" & IIf(Trim(MbCarn) = "", Chr(32), Trim(MbCarn)) & "'," & CmbMes1.BoundText & "," & CmbAno.BoundText & "," & IIf(OptEfe.Value, "", CDbl(MBTot_P) & ",") & "'" & MbNAsie & "','" & StrCad & "'," & Val(MbTipC) & "," & CDbl(MbMonCTS) & "," & CDbl(MbMonCTS) & "," & IIf(OptEfe.Value, 1, 2) & _
            "," & CDbl(TotIpSS) & "," & CDbl(TotIpSS) & "," & CDbl(TotAFP) & "," & CDbl(TotAFP) & "," & CDbl(MBTot_A) & "," & CDbl(MBTot_A) & "," & Val(MbPor) & ")"
            BD_Sistema.Execute Str_Sql
            BD_Sistema.Execute "Update TBCorrela Set Nume=Nume + 1 Where Numero=33;"
            
            For I = 1 To Grid1.Rows - 1 ' Remuneraciones
               BD_Sistema.Execute "Insert Into TBD_BRem (No_Bol,Num_Rem,Monto,Rem_Porc,Sw_Aplic,ID_PCta)Values ('" & TxtNroBol & "'," & _
               Val(Grid1.TextMatrix(I, 0)) & "," & CDbl(Grid1.TextMatrix(I, 2)) & "," & Val(Grid1.TextMatrix(I, 4)) & "," & Val(Grid1.TextMatrix(I, 5)) & ",'" & IIf(Trim(Grid1.TextMatrix(I, 6)) = "", Chr(32), Trim(Grid1.TextMatrix(I, 6))) & "')"
            Next
            
            For I = 1 To Grid2.Rows - 1 ' Descuento y Reducciones del trabajador
               BD_Sistema.Execute "Insert Into TBD_BDDT (No_Bol,Num_DDT,Monto,Saldo,DDt_Porc,Sw_Aplic,ID_PCta)Values ('" & TxtNroBol & _
               "'," & Val(Grid2.TextMatrix(I, 0)) & "," & CDbl(Grid2.TextMatrix(I, 2)) & "," & CDbl(Grid2.TextMatrix(I, 2)) & "," & Val(Grid2.TextMatrix(I, 4)) & _
               "," & Val(Grid2.TextMatrix(I, 5)) & ",'" & IIf(Trim(Grid2.TextMatrix(I, 6)) = "", Chr(32), Trim(Grid2.TextMatrix(I, 6))) & "')"
            Next
            
            For I = 1 To Grid3.Rows - 1 ' Aportaciones del Empleador
               BD_Sistema.Execute "Insert Into TBD_BAE (No_Bol,Num_AE,Monto,Saldo,AE_Porc,Sw_Aplic,ID_PCta,ID_PCta1)Values ('" & TxtNroBol & _
               "'," & Val(Grid3.TextMatrix(I, 0)) & "," & CDbl(Grid3.TextMatrix(I, 2)) & "," & CDbl(Grid3.TextMatrix(I, 2)) & "," & Val(Grid3.TextMatrix(I, 4)) & "," & Val(Grid3.TextMatrix(I, 5)) & ",'" & _
               IIf(Trim(Grid3.TextMatrix(I, 6)) = "", Chr(32), Trim(Grid3.TextMatrix(I, 6))) & "','" & IIf(Trim(Grid3.TextMatrix(I, 7)) = "", Chr(32), Trim(Grid3.TextMatrix(I, 7))) & "')"
            Next
            
            Call GuaCond(CmbMes1.BoundText, 33)
            Habilita_Boton
            DesEdic
            Call Mensaje(1)
        End If
     Else
            If Val(TxtNroBol) = 0 Then MsgBox "Falta Nro. de Boleta de Pago....", vbExclamation, "Error": Exit Sub
            If CmbEmp.BoundText = "" Then MsgBox "Seleccione Empleado , por favor... ", vbExclamation, "Error": CmbEmp.SetFocus: Exit Sub
            If CmbMes1.BoundText = "" Then MsgBox "Seleccione La Mes , por favor... ", vbExclamation, "Error": CmbMes1.SetFocus: Exit Sub
            If CmbAno.BoundText = "" Then MsgBox "Seleccione La Año, por favor... ", vbExclamation, "Error": CmbAno.SetFocus: Exit Sub
            If Grid1.Rows = 1 Then MsgBox "Falta los Coceptos de  REMUNERACIONES ... ", vbExclamation, "Error": Exit Sub
            If Grid2.Rows = 1 Then MsgBox "Falta los Coceptos de REDUCCIONES Y DESCUENTOS AL TRBAJADOR... ", vbExclamation, "Error": Exit Sub
            If Grid3.Rows = 1 Then MsgBox "Falta los Coceptos de  APORTACIONES DEL EMPLEADOR... ", vbExclamation, "Error": Exit Sub
     End If
  End If
Exit Sub
Errores:
 If Ctrl_Error(Err, RstBol) Then Exit Sub
End Sub
Sub Gua_Asie()
    ' Remuneraciones del Emepleado
    Y = 1
    TotImp = 34
    For I = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(I, 2)) <> 0 Then
            CodGene = Trim(Grid1.TextMatrix(I, 6))
            TotGen = Format(CDbl(Grid1.TextMatrix(I, 2)) / Val(MbTipC.Text), "#0.#00")
            TotGenS = CDbl(Grid1.TextMatrix(I, 2))
            
            Str_Sql = " Values ('" & CodGene & "','" & MbNAsie & "','" & StrCad & "'," & Y & ",34," & CmbEmp.BoundText & ",2,'" & TxtNroBol & "',#" & _
            Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(DtpFecha, "mm/dd/yyyy") & "#,1,3,1," & CDbl(TotGen) & "," & CDbl(TotGenS) & "," & Val(MbTipC.Text) & ",43,43,'Planilla'," & Val(Label9(0).Caption) & ",'','" & CmbEmp.Text & "')"
            BD_CONTA.Execute "Insert Into TCDeDGen (ID_PCta,NroAsto,MesAno,NroItem,Nro_OCon,ID_Emp,Id_Aux,Nro_Doc,Fec_doc,Fec_Vence,Id_Mon,Afecto,TPOMOV,Tot_usa,Tot_Nac,Tip_Cam,Id_Doc,Id_Doc1,Glosa,numerocuo,numasocia,nombreasocia) " & Str_Sql
            
            BAAutom
        End If
        Y = Y + 1
    Next
    For I = 1 To Grid3.Rows - 1 ' Seguros Debe
        If Val(Grid3.TextMatrix(I, 2)) <> 0 Then
            'TotGen = CDbl(Grid3.TextMatrix(I, 2))
            TotGen = Format(CDbl(Grid3.TextMatrix(I, 2)) / Val(MbTipC.Text), "#0.#00")
            TotGenS = CDbl(Grid3.TextMatrix(I, 2))
            Str_Sql = " Values ('" & Trim(Grid3.TextMatrix(I, 6)) & "','" & MbNAsie & "','" & StrCad & "'," & Y & ",34," & CmbEmp.BoundText & ",2,'" & TxtNroBol & "',#" & _
            Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(DtpFecha, "mm/dd/yyyy") & "#,1,3,1," & CDbl(TotGen) & "," & CDbl(TotGenS) & "," & Val(MbTipC.Text) & ",43,43,'Planilla'," & Val(Label9(0).Caption) & ",'','" & CmbEmp.Text & "')"
            BD_CONTA.Execute "Insert Into TCDeDGen (ID_PCta,NroAsto,MesAno,NroItem,Nro_OCon,ID_Emp,Id_Aux,Nro_Doc,Fec_doc,Fec_Vence,Id_Mon,Afecto,TPOMOV,Tot_usa,Tot_Nac,Tip_Cam,Id_Doc,Id_Doc1,Glosa,numerocuo,numasocia,nombreasocia) " & Str_Sql
            CodGene = Trim(Grid3.TextMatrix(I, 6))
            BAAutom
        End If
        Y = Y + 1
    Next
    For I = 1 To Grid3.Rows - 1 ' Seguros Haber
        If Val(Grid3.TextMatrix(I, 2)) <> 0 Then
            'TotGen = CDbl(Grid3.TextMatrix(I, 2))
            TotGen = Format(CDbl(Grid3.TextMatrix(I, 2)) / Val(MbTipC.Text), "#0.#00")
            TotGenS = CDbl(Grid3.TextMatrix(I, 2))
            Str_Sql = " Values ('" & Trim(Grid3.TextMatrix(I, 7)) & "','" & MbNAsie & "','" & StrCad & "'," & Y & ",34," & CmbEmp.BoundText & ",2,'" & TxtNroBol & "',#" & _
            Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(DtpFecha, "mm/dd/yyyy") & "#,1,3,2," & CDbl(TotGen) & "," & CDbl(TotGenS) & "," & Val(MbTipC.Text) & ",43,43,'Planilla'," & Val(Label9(0).Caption) & ",'','" & CmbEmp.Text & "')"
            BD_CONTA.Execute "Insert Into TCDeDGen (ID_PCta,NroAsto,MesAno,NroItem,Nro_OCon,ID_Emp,Id_Aux,Nro_Doc,Fec_doc,Fec_Vence,Id_Mon,Afecto,TPOMOV,Tot_usa,Tot_Nac,Tip_Cam,Id_Doc,Id_Doc1,Glosa,numerocuo,numasocia,nombreasocia) " & Str_Sql
            CodGene = Trim(Grid3.TextMatrix(I, 7))
            BAAutom
        End If
        Y = Y + 1
    Next
    For I = 1 To Grid2.Rows - 1 ' Descuentos y Deducciones al Trabajador
        If Val(Grid2.TextMatrix(I, 2)) <> 0 Then
            'TotGen = CDbl(Grid2.TextMatrix(I, 2))
            TotGen = Format(CDbl(Grid2.TextMatrix(I, 2)) / Val(MbTipC.Text), "#0.#00")
            TotGenS = CDbl(Grid2.TextMatrix(I, 2))
            Str_Sql = " Values ('" & Trim(Grid2.TextMatrix(I, 6)) & "','" & MbNAsie & "','" & StrCad & "'," & Y & ",34," & CmbEmp.BoundText & ",2,'" & TxtNroBol & "',#" & _
            Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(DtpFecha, "mm/dd/yyyy") & "#,1,3,2," & CDbl(TotGen) & "," & CDbl(TotGenS) & "," & Val(MbTipC.Text) & ",43,43,'Planilla'," & Val(Label9(0).Caption) & ",'','" & CmbEmp.Text & "')"
            BD_CONTA.Execute "Insert Into TCDeDGen (ID_PCta,NroAsto,MesAno,NroItem,Nro_OCon,ID_Emp,Id_Aux,Nro_Doc,Fec_doc,Fec_Vence,Id_Mon,Afecto,TPOMOV,Tot_usa,Tot_Nac,Tip_Cam,Id_Doc,Id_Doc1,Glosa,numerocuo,numasocia,nombreasocia) " & Str_Sql
            CodGene = Trim(Grid2.TextMatrix(I, 6))
            BAAutom
        End If
        Y = Y + 1
    Next
    'Neto a pagar
    'TotGen = CDbl(MBTot_P)
    TotGen = Format(CDbl(MBTot_P) / Val(MbTipC.Text), "#0.#00")
    TotGenS = CDbl(MBTot_P)
    Y = Y + 1
    Str_Sql = " Values ('41101001','" & MbNAsie & "','" & StrCad & "'," & Y & ",34," & CmbEmp.BoundText & ",2,'" & TxtNroBol & "',#" & _
    Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(DtpFecha, "mm/dd/yyyy") & "#,1,3,2," & CDbl(TotGen) & "," & CDbl(TotGenS) & "," & Val(MbTipC.Text) & ",43,43,'Planilla'," & Val(Label9(0).Caption) & ",'','" & CmbEmp.Text & "')"
    BD_CONTA.Execute "Insert Into TCDeDGen (ID_PCta,NroAsto,MesAno,NroItem,Nro_OCon,ID_Emp,Id_Aux,Nro_Doc,Fec_doc,Fec_Vence,Id_Mon,Afecto,TPOMOV,Tot_usa,Tot_Nac,Tip_Cam,Id_Doc,Id_Doc1,Glosa,numerocuo,numasocia,nombreasocia) " & Str_Sql
    CodGene = "41101001"
    BAAutom

    If OptEfe.Value Then
        Y = Y + 1
        Str_Sql = " Values ('10101001','41101001','" & MbNAsie & "','" & StrCad & "'," & Y & ",6," & CmbEmp.BoundText & ",2,'" & TxtNroBol & "',#" & _
        Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(DtpFecha, "mm/dd/yyyy") & "#,1,3,1," & CDbl(TotGen) & "," & CDbl(TotGenS) & "," & Val(MbTipC.Text) & ",43,43,'Planilla'," & Val(Label9(0).Caption) & ",'','" & CmbEmp.Text & "')"
        BD_CONTA.Execute "Insert Into TCDeDGen (cuentacab,ID_PCta,NroAsto,MesAno,NroItem,Nro_OCon,ID_Emp,Id_Aux,Nro_Doc,Fec_doc,Fec_Vence,Id_Mon,Afecto,TPOMOV,Tot_usa,Tot_Nac,Tip_Cam,Id_Doc,Id_Doc1,Glosa,numerocuo,numasocia,nombreasocia) " & Str_Sql
        CodGene = "41101001"
        TotImp = 6
        BAAutom
        
        Y = Y + 1
        Str_Sql = " Values ('10101001','10101001','" & MbNAsie & "','" & StrCad & "'," & Y & ",6," & CmbEmp.BoundText & ",2,'" & TxtNroBol & "',#" & _
        Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(DtpFecha, "mm/dd/yyyy") & "#,1,3,2," & CDbl(TotGen) & "," & CDbl(TotGenS) & "," & Val(MbTipC.Text) & ",43,43,'Planilla'," & Val(Label9(0).Caption) & ",'','" & CmbEmp.Text & "')"
        BD_CONTA.Execute "Insert Into TCDeDGen (cuentacab,ID_PCta,NroAsto,MesAno,NroItem,Nro_OCon,ID_Emp,Id_Aux,Nro_Doc,Fec_doc,Fec_Vence,Id_Mon,Afecto,TPOMOV,Tot_usa,Tot_Nac,Tip_Cam,Id_Doc,Id_Doc1,Glosa,numerocuo,numasocia,nombreasocia) " & Str_Sql
        CodGene = "10101001"
        BAAutom
        
    End If
    'TotGen = CDbl(MbMonCTS)
    TotGen = Format(CDbl(MbMonCTS) / Val(MbTipC.Text), "#0.#00")
    TotGenS = CDbl(MbMonCTS)
    'Provicionando CTS Debe
    Y = Y + 1
    TotImp = 36
    Str_Sql = " Values ('68601001','" & MbNAsie & "','" & StrCad & "'," & Y & ",36," & CmbEmp.BoundText & ",2,'" & TxtNroBol & "',#" & _
    Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(DtpFecha, "mm/dd/yyyy") & "#,1,3,1," & CDbl(TotGen) & "," & CDbl(TotGenS) & "," & Val(MbTipC.Text) & ",43,43,'Planilla'," & Val(Label9(0).Caption) & ",'','" & CmbEmp.Text & "')"
    BD_CONTA.Execute "Insert Into TCDeDGen (ID_PCta,NroAsto,MesAno,NroItem,Nro_OCon,ID_Emp,Id_Aux,Nro_Doc,Fec_doc,Fec_Vence,Id_Mon,Afecto,TPOMOV,Tot_usa,Tot_Nac,Tip_Cam,Id_Doc,Id_Doc1,Glosa,numerocuo,numasocia,nombreasocia) " & Str_Sql
    CodGene = "68601001"
    BAAutom
    'Provicionando CTS haber
    Y = Y + 1
    Str_Sql = " Values ('47101001','" & MbNAsie & "','" & StrCad & "'," & Y & ",36," & CmbEmp.BoundText & ",2,'" & TxtNroBol & "',#" & _
    Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(DtpFecha, "mm/dd/yyyy") & "#,1,3,2," & CDbl(TotGen) & "," & CDbl(TotGenS) & "," & Val(MbTipC.Text) & ",43,43,'Planilla'," & Val(Label9(0).Caption) & ",'','" & CmbEmp.Text & "')"
    BD_CONTA.Execute "Insert Into TCDeDGen (ID_PCta,NroAsto,MesAno,NroItem,Nro_OCon,ID_Emp,Id_Aux,Nro_Doc,Fec_doc,Fec_Vence,Id_Mon,Afecto,TPOMOV,Tot_usa,Tot_Nac,Tip_Cam,Id_Doc,Id_Doc1,Glosa,numerocuo,numasocia,nombreasocia) " & Str_Sql
    CodGene = "47101001"
    BAAutom
End Sub
Sub BAAutom() ' Buscando los Asientos Automaticos
    Y = Y + 1
    Set RstBol = New ADODB.Recordset
    RstBol.Open "Select CodCta1,CodCta2,GCAuto From TCPlaCta Where ID_PCta='" & CodGene & "'", BD_Sistema, adOpenDynamic, adLockReadOnly, adCmdText
    With RstBol
        If Not (.EOF() Or .BOF()) Then
           If !GCAuto = 1 Then
               If Trim(!CodCta1) <> "" Then
                    Str_Sql = " Values ('" & !CodCta1 & "','" & MbNAsie & "','" & StrCad & "'," & Y & "," & TotImp & "," & CmbEmp.BoundText & ",2,'" & TxtNroBol & "',#" & _
                    Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(DtpFecha, "mm/dd/yyyy") & "#,1,3,1," & CDbl(TotGen) & "," & CDbl(TotGenS) & "," & Val(MbTipC.Text) & ",43,43,'Planilla'," & Val(Label9(0).Caption) & ",'','" & CmbEmp.Text & "',1)"
                    BD_CONTA.Execute "Insert Into TCDeDGen (ID_PCta,NroAsto,MesAno,NroItem,Nro_OCon,ID_Emp,Id_Aux,Nro_Doc,Fec_doc,Fec_Vence,Id_Mon,Afecto,TPOMOV,Tot_usa,Tot_Nac,Tip_Cam,Id_Doc,Id_Doc1,Glosa,numerocuo,numasocia,nombreasocia,cuentaauto) " & Str_Sql
                    Y = Y + 1
                    Str_Sql = " Values ('" & !CodCta2 & "','" & MbNAsie & "','" & StrCad & "'," & Y & "," & TotImp & "," & CmbEmp.BoundText & ",2,'" & TxtNroBol & "',#" & _
                    Format(DtpFecha, "mm/dd/yyyy") & "#,#" & Format(DtpFecha, "mm/dd/yyyy") & "#,1,3,2," & CDbl(TotGen) & "," & CDbl(TotGenS) & "," & Val(MbTipC.Text) & ",43,43,'Planilla'," & Val(Label9(0).Caption) & ",'','" & CmbEmp.Text & "',1)"
                    BD_CONTA.Execute "Insert Into TCDeDGen (ID_PCta,NroAsto,MesAno,NroItem,Nro_OCon,ID_Emp,Id_Aux,Nro_Doc,Fec_doc,Fec_Vence,Id_Mon,Afecto,TPOMOV,Tot_usa,Tot_Nac,Tip_Cam,Id_Doc,Id_Doc1,Glosa,numerocuo,numasocia,nombreasocia,cuentaauto) " & Str_Sql
               End If
           End If
        End If
    End With
    RstBol.Close
    Set RstBol = Nothing
End Sub
Public Sub Cal_Porc()
 If Grid1.Rows >= 2 And Grid2.Rows >= 2 And Grid3.Rows >= 3 And CDbl(MbSue) <> 0 And CmbEmp.BoundText <> "" Then
     TotGen = 0
     'TotSuel = CDbl(MBTot_R)
     TotSuel = CDbl(MBTot_R2)
     
     For I = 1 To Grid2.Rows - 1
        If Grid2.TextMatrix(I, 4) > 0 And Val(Grid2.TextMatrix(I, 5)) = 1 Then
           Grid2.TextMatrix(I, 2) = Format((CDbl(TotSuel) * (Val(Grid2.TextMatrix(I, 4)) / 100)), "#0.#0")
        End If
        TotGen = TotGen + CDbl(Grid2.TextMatrix(I, 2))
     Next
     MBTot_D = TotGen
     MBTot_P = MBTot_R - TotGen
     TotGen = 0
     For I = 1 To Grid3.Rows - 1
        If Grid3.TextMatrix(I, 4) > 0 And Val(Grid3.TextMatrix(I, 5)) = 1 Then
           If Val(Grid3.TextMatrix(I, 0)) <> 1 Then
              Grid3.TextMatrix(I, 2) = Format((CDbl(MbSue) * (Val(Grid3.TextMatrix(I, 4)) / 100)), "#0.#0")
           Else
              Grid3.TextMatrix(I, 2) = Format((CDbl(TotSuel) * (Val(Grid3.TextMatrix(I, 4)) / 100)), "#0.#0")
           End If
        End If
        TotGen = TotGen + CDbl(Grid3.TextMatrix(I, 2))
     Next
     MBTot_A = TotGen
  End If
End Sub


Private Sub Command1_Click()
Dim RstTMP As ADODB.Recordset
 If MsgBox("Estas Seguro generar los Asientos Contables?", vbYesNo, "Aviso") = vbYes Then
    'Eliminado los tablas temporales contabilidad
    Str_Sql = "DELETE * FROM TCDeDGen;"
    BD_CONTA.Execute Str_Sql
     
    Str_Sql = "DELETE * FROM TCDeDTMP;"
    BD_CONTA.Execute Str_Sql
    
    Str_Sql = "DELETE * FROM TCDeDGen;"
    BD_CONTA.Execute Str_Sql
    
    Str_Sql = "DELETE * FROM TCDeTMP2;"
    BD_CONTA.Execute Str_Sql
    
    Str_Sql = "DELETE * FROM TCTMP_AD;"
    BD_CONTA.Execute Str_Sql
    
    Str_Sql = "DELETE * FROM TCTMPBCO;"
    BD_CONTA.Execute Str_Sql
    
    Str_Sql = "DELETE * FROM TCDeDCom;"
    BD_CONTA.Execute Str_Sql
    
    RstEmp.Close
    Set RstEmp = New ADODB.Recordset
    RstEmp.Open " SELECT TBEmple.Id_Emp, [ApePat] & ' ' & [ApeMat] & ', ' & [Nombre]  AS Nom From TBEmple ORDER BY [ApePat],[Nombre]", BD_Sistema, adOpenKeyset, adLockOptimistic
    Set CmbEmp.RowSource = RstEmp
    RstEmp.Requery
 
    Set RstTMP = New ADODB.Recordset
    RstTMP.Open "SELECT * From TBBol_Pago Where Left(Mes_Ano,2) ='" & Right(Val_Año, 2) & "' And Anular=0 Order By NroAsto", BD_Sistema, adOpenDynamic, adLockReadOnly, adCmdText
    If Not (RstTMP.EOF() Or RstTMP.BOF()) Then
       BD_CONTA.Execute "Delete Sw_Anu From TCDeDGen Where Id_Doc = 43 And Id_Doc1 =43 And left(MesAno,2)='" & Right(Val_Año, 2) & "'"
       Do While Not RstTMP.EOF()
            Label9(0).Caption = RstTMP!numerocuo
            CodGene = RstTMP!No_Bol
            Set RstBol = New ADODB.Recordset
            RstBol.Open "Select * From TBBol_Pago Where No_Bol='" & CodGene & "'", BD_Sistema, adOpenDynamic, adLockReadOnly, adCmdText
            With RstBol
                If Not (.EOF() Or .BOF()) Then
                    TxtNroBol = CodGene
                    DtpFecha = ![Fecha]
                    CmbAno.BoundText = ![Id_Año]
                    CmbMes1.BoundText = ![Id_Mes]
                    CmbEmp.BoundText = ![Id_Empl]
                    MbSue = ![Sueldo]
                    MbDias = ![Dias_Lab]
                    MBTot_R = ![Tot_Remu]
                    MBTot_D = ![Tot_DD]
                    MBTot_P = ![Neto_Pag]
                    MBTot_A = ![Tot_Aport]
                    MbNAsie = IIf(IsNull(![NroAsto]), "", ![NroAsto])
                    CmbMes1.BoundText = Val(Right(![Mes_Ano], 2))
                    MbAño = Left(![Mes_Ano], 2)
                    MbTipC = ![TPOCMB]
                    If !Tip_Pago = 1 Then OptEfe.Value = True Else OptPro.Value = True
                    MbMonCTS = ![Mon_CTS]
                End If
                .Close
            End With
            Set RstBol = Nothing
             Str_Sql = " SELECT TBD_BRem.Num_Rem AS Cod, TB_REMU.Des_Rem AS DescV, TBD_BRem.Monto AS Mon,0 AS Dscto, TB_REMU.ID_PCta, '0' AS ID_PCta2, TB_REMU.Sw_Aplic FROM TBD_BRem INNER JOIN TB_REMU ON TBD_BRem.Num_Rem = TB_REMU.Num_Rem" & _
                       " Where TBD_BRem.No_Bol='" & CodGene & "' ORDER BY TBD_BRem.Numero;"
             AddIt Grid1, 0
             
             Str_Sql = " SELECT TBD_BDDT.Num_DDT AS Cod, TB_DDT.Des_DDT AS DescV, TBD_BDDT.Monto AS Mon, TB_DDT.ID_PCta, '0' AS ID_PCta2, 0 AS Aport, TBD_BDDT.DDT_Porc AS [Dscto], TBD_BDDT.Sw_Aplic FROM TB_DDT INNER JOIN TBD_BDDT ON TB_DDT.Num_DDT = TBD_BDDT.Num_DDT" & _
                       " Where TBD_BDDT.No_Bol='" & CodGene & "' ORDER BY TBD_BDDT.Numero;"
             AddIt Grid2, 0
    
             Str_Sql = " SELECT TBD_BAE.Num_AE AS Cod, TB_AEMP.Des_AE AS DescV, TBD_BAE.Monto AS Mon, TB_AEMP.ID_PCta, TB_AEMP.ID_PCta1 AS ID_PCta2, TBD_BAE.Ae_Porc AS Dscto, TBD_BAE.Sw_Aplic FROM TB_AEMP INNER JOIN TBD_BAE ON TB_AEMP.Num_AE = TBD_BAE.Num_AE" & _
                       " Where TBD_BAE.No_Bol='" & CodGene & "' ORDER BY TBD_BAE.Numero;"
             AddIt Grid3, 0
             Val_Mes = CmbMes1.BoundText
             StrCad = Right(Val_Año, 2) & Right("0" & Val_Mes, 2)
             'Buscando nro. de plan contable
             Gua_Asie
             RstTMP.MoveNext
       Loop
       RstTMP.Close
       Set RstTMP = Nothing
       
    RstEmp.Close
    Set RstEmp = New ADODB.Recordset
    RstEmp.Open " SELECT TBEmple.Id_Emp, [ApePat] & ' ' & [ApeMat] & ', ' & [Nombre]  AS Nom From TBEmple Where Sw_Tipo=0 And Est_Emp=0 ORDER BY [ApePat],[Nombre]", BD_Sistema, adOpenKeyset, adLockOptimistic
    Set CmbEmp.RowSource = RstEmp
    RstEmp.Requery
            
    End If
    Call Mensaje(1)
 Else
   MsgBox "Se canceló correctamente !!!", vbExclamation, "Aviso"
 End If
End Sub

Private Sub Form_Load()
    Set RstCar = New ADODB.Recordset
    RstCar.Open "Select * From TBCargo;", BD_Sistema, adOpenKeyset, adLockOptimistic
    Set CmbCar.RowSource = RstCar
    
    Set RstAno = New ADODB.Recordset
    RstAno.Open "Select * From TBAno;", BD_Sistema, adOpenKeyset, adLockOptimistic
    Set CmbAno.RowSource = RstAno
    
    Set RstEmp = New ADODB.Recordset
    RstEmp.Open " SELECT TBEmple.Id_Emp, ([ApePat] & ' ' & [ApeMat] & ', ' & [Nombre]) AS Nom From TBEmple Where Sw_Tipo=0 And Est_Emp=0 ORDER BY [ApePat],[ApeMat],[Nombre]", BD_Sistema, adOpenKeyset, adLockOptimistic
    Set CmbEmp.RowSource = RstEmp
    
    ActEdic ("N")
    DesHBot
    MbSue = "0.00"
    MBTot_R = "0.00"
    MBTot_P = "0.00"
    MBTot_D = "0.00"
    MBTot_A = "0.00"
    MbDias = "0"
    Set RstBol = New ADODB.Recordset
    RstBol.Open "Select * From TBCorrela Where Numero=33;", BD_Sistema, adOpenDynamic, adLockReadOnly, adCmdText
    TxtNroBol = Right("0000" & (RstBol!NUME) + 1, 5)
    RstBol.Close
    Set RstBol = Nothing
    DtpFecha = Date
    CmbAno.BoundText = Val_Año
    
    Grid1.Row = 0
    Grid1.Cols = 8
    Grid1.Col = 0: Grid1.ColWidth(0) = 450: Grid1.CellAlignment = 4: Grid1.Text = "Cód."
    Grid1.Col = 1: Grid1.ColWidth(1) = 3600: Grid1.CellAlignment = 4: Grid1.Text = "Desc. de Concepto"
    Grid1.Col = 2: Grid1.ColWidth(2) = 1000: Grid1.CellAlignment = 4: Grid1.Text = "Monto"
    Grid1.Col = 3: Grid1.ColWidth(3) = 0 ' No sirve
    Grid1.Col = 4: Grid1.ColWidth(4) = 600 ' Dscto. Porcentaje
    Grid1.Col = 5: Grid1.ColWidth(5) = 200 ' Sw Descto
    Grid1.Col = 6: Grid1.ColWidth(6) = 1500: Grid1.Text = "N° Cuenta" ' Plan de Cuenta
    Grid1.ColWidth(7) = 0 ' Plan de Cuenta
    Grid1.Rows = 1
    
    Grid2.Row = 0
    Grid2.Cols = 8
    Grid2.Col = 0: Grid2.ColWidth(0) = 450: Grid2.CellAlignment = 4: Grid2.Text = "Cód."
    Grid2.Col = 1: Grid2.ColWidth(1) = 3500: Grid2.CellAlignment = 4: Grid2.Text = "Desc. de Concepto"
    Grid2.Col = 2: Grid2.ColWidth(2) = 1000: Grid2.CellAlignment = 4: Grid2.Text = "Monto"
    Grid2.Col = 3: Grid2.ColWidth(3) = 900 ' No sirve
    Grid2.Col = 4: Grid2.ColWidth(4) = 900 ' Dscto. Porcentaje
    Grid2.Col = 5: Grid2.ColWidth(5) = 900 ' Sw Descto
    Grid2.Col = 6: Grid2.ColWidth(6) = 1500: Grid2.Text = "N° Cuenta" ' Plan de Cuenta
    Grid2.ColWidth(7) = 0 ' Plan de Cuenta
    Grid2.Rows = 1
    
    Grid3.Row = 0
    Grid3.Cols = 8
    Grid3.Col = 0: Grid3.ColWidth(0) = 450: Grid3.CellAlignment = 4: Grid3.Text = "Cód."
    Grid3.Col = 1: Grid3.ColWidth(1) = 3600: Grid3.CellAlignment = 4: Grid3.Text = "Desc. de Concepto"
    Grid3.Col = 2: Grid3.ColWidth(2) = 1200: Grid3.CellAlignment = 4: Grid3.Text = "Monto"
    Grid3.Col = 3: Grid3.ColWidth(3) = 0 ' No sirve
    Grid3.Col = 4: Grid3.ColWidth(4) = 600 ' Dscto. Porcentaje
    Grid3.Col = 5: Grid3.ColWidth(5) = 200 ' Sw Descto
    Grid3.Col = 6: Grid3.ColWidth(6) = 1500: Grid3.Text = "N° Cta. 1"  ' Plan de Cuenta
    Grid3.Col = 7: Grid3.ColWidth(7) = 1500: Grid3.Text = "N° Cta. 2" ' Plan de Cuenta
    Grid3.Rows = 1
    
    Set RstMes1 = New ADODB.Recordset
    RstMes1.Open " SELECT * From TbMes Order By Id_Mes;", BD_Sistema, adOpenKeyset, adLockOptimistic
    Set CmbMes1.RowSource = RstMes1
    CmbMes1.BoundText = DtpFecha.Month
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    RstCar.Close
    RstEmp.Close
    RstAno.Close
    RstMes1.Close
    Set RstAno = Nothing
    Set RstCar = Nothing
    Set RstEmp = Nothing
    Set RstMes1 = Nothing
End Sub


Private Sub Grid1_KeyPress(KeyAscii As Integer)
    XFila = Grid1.Row
    xCol = Grid1.Col
    If xCol = 2 And (CmdNuevo.Tag = "G" Or CmdModifi.Tag = "G") Then
        If KeyAscii = 8 Then
           Grid1.Text = ""
           Cal_Rem
        End If
        If KeyAscii = 46 Then
           If InStr(Grid1.Text, ".") Then Exit Sub
        End If
        If Len(Grid1.Text) > 8 Then Exit Sub
        If KeyAscii = 46 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
           Grid1.Text = Grid1.TextMatrix(XFila, xCol) & Chr(KeyAscii)
           Cal_Rem
           MbMonCTS = Format(CDbl(MBTot_R) * (Val(MbPor) / 100), "#0.#0")
        End If
        Cal_Porc
        'Recalculando Descuento del Empleado
        Cal_DDT
        'Recalculando Aportaciones del empleador
        Cal_AE
    End If
End Sub

Private Sub Grid2_KeyPress(KeyAscii As Integer)
    XFila = Grid2.Row
    xCol = Grid2.Col
    If xCol = 2 And (CmdNuevo.Tag = "G" Or CmdModifi.Tag = "G") Then
        If KeyAscii = 8 Then
           Grid2.Text = ""
           Cal_DDT
        End If
        If KeyAscii = 46 Then
           If InStr(Grid2.Text, ".") Then Exit Sub
        End If
        If Len(Grid2.Text) > 8 Then Exit Sub
        If KeyAscii = 46 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
           Grid2.Text = Grid2.TextMatrix(XFila, xCol) & Chr(KeyAscii)
           Cal_DDT
        End If
    End If
End Sub

Private Sub Grid3_KeyPress(KeyAscii As Integer)
    XFila = Grid3.Row
    xCol = Grid3.Col
    If xCol = 2 And (CmdNuevo.Tag = "G" Or CmdModifi.Tag = "G") Then
        If KeyAscii = 8 Then
           Grid3.Text = ""
           Cal_AE
        End If
        If KeyAscii = 46 Then
           If InStr(Grid3.Text, ".") Then Exit Sub
        End If
        If Len(Grid3.Text) > 8 Then Exit Sub
        If KeyAscii = 46 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
           Grid3.Text = Grid3.TextMatrix(XFila, xCol) & Chr(KeyAscii)
           Cal_AE
        End If
    End If

End Sub


Private Sub MbCarn_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then MbSue.SetFocus
End Sub

Private Sub MbDias_Change()
   If (MbDias.Text) = "" Then
      MbDias.Text = "0.00"
      MbDias.SelStart = 0
      MbDias.SelLength = Len(MbDias)
   End If
End Sub
Private Sub MbDias_GotFocus()
    MbDias.SelStart = 0
    MbDias.SelLength = Len(MbDias)
End Sub
Private Sub MbDias_KeyPress(KeyAscii As Integer)
KeyAscii = Key_Numero(KeyAscii)
If KeyAscii = 13 Then MbNOrd.SetFocus
End Sub

Private Sub MbFecC_KeyPress(KeyAscii As Integer)
Call Key_Fecha(KeyAscii)
End Sub

Private Sub MbFecC_LostFocus()
    If Trim(MbFecC) <> "" Then
        If Not IsDate(MbFecC) Then
           MsgBox "Fecha Incorrecta ...", vbExclamation, "Aviso"
           MbFecC = ""
        End If
    End If

End Sub

Private Sub MbFecIng_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then MbFecC.SetFocus
End Sub

Private Sub MbNOrd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then MbFecIng.SetFocus
End Sub

Private Sub MbSue_Change()
   If (MbSue.Text) = "" Then
      MbSue.Text = "0.00"
      MbSue.SelStart = 0
      MbSue.SelLength = Len(MbSue)
   End If
End Sub
Private Sub MbSue_GotFocus()
    MbSue.SelStart = 0
    MbSue.SelLength = Len(MbSue)
End Sub
Private Sub MbSue_KeyPress(KeyAscii As Integer)
KeyAscii = Key_Precio(KeyAscii, MbSue)
    If KeyAscii = 13 Then MbDias.SetFocus
End Sub

Private Sub SSCommand1_Click()
   If CmbEmp.BoundText <> "" Then
         Str_Sql = " SELECT TBH_E_Rem.Num_Rem As Cod,TBH_E_Rem.Remu_Por As Dscto, TB_REMU.Des_Rem As DescV,[Monto] AS Mon,TBH_E_Rem.Sw_Aplic,TB_REMU.ID_PCta,'0' As ID_PCta2 FROM TBH_E_Rem INNER JOIN TB_REMU ON TBH_E_Rem.Num_Rem = TB_REMU.Num_Rem" & _
                   " Where TBH_E_Rem.Id_Empl=" & CmbEmp.BoundText & " ORDER BY TBH_E_Rem.Num_Rem;"
         AddIt Grid1, 0
         'If Grid1.Rows >= 1 Then
         '   MbSue = Format(CDbl(Grid1.TextMatrix(1, 2)) * 2, "##,##0.#0")
         'End If
         Cal_Rem
         Str_Sql = " SELECT TBH_E_DDT.Num_DDt AS Cod, TB_DDT.Des_DDT AS DescV,TBH_E_DDT.Dscto, [TBH_E_DDT].[Monto] AS Mon,TBH_E_DDT.Sw_Aplic,TB_DDT.ID_PCta,'0' As ID_PCta2 FROM TB_DDT INNER JOIN TBH_E_DDT ON TB_DDT.Num_DDT = TBH_E_DDT.Num_DDt" & _
                   " Where TBH_E_DDT.Id_Empl=" & CmbEmp.BoundText & " ORDER BY TBH_E_DDT.Num_DDt;"
         AddIt Grid2, 0
         
         'Mon_Ant
         Cal_DDT
         Str_Sql = " SELECT TBH_E_AE.Num_AE As Cod, TB_AEMP.Des_AE As DescV,TBH_E_AE.Aport As Dscto, [Monto] AS Mon,TBH_E_AE.Sw_Aplic,TB_AEMP.ID_PCta,TB_AEMP.ID_PCta1 As ID_PCta2  FROM TB_AEMP INNER JOIN TBH_E_AE ON TB_AEMP.Num_AE = TBH_E_AE.Num_AE" & _
                   " Where TBH_E_AE.ID_Empl=" & CmbEmp.BoundText & " ORDER BY TBH_E_AE.Num_AE;"
         AddIt Grid3, 0
         Cal_AE
   Else
      MsgBox "Seleccione Empleado, Por favor ....", vbExclamation, "Aviso..."
      CmbEmp.SetFocus
   End If
End Sub

Public Sub AddIt(Ob_Grid As MSFlexGrid, Val_Q As Byte)
  Ob_Grid.Rows = 1
  Set RstBol = New ADODB.Recordset
  RstBol.Open Str_Sql, BD_Sistema, adOpenDynamic, adLockReadOnly, adCmdText
    With RstBol
       If Not (.EOF() Or .BOF()) Then
           XFila = 1
           Ob_Grid.Col = 2
           Do While Not .EOF()
                Ob_Grid.Rows = XFila + 1
                Ob_Grid.Col = 2
                Ob_Grid.TextMatrix(XFila, 0) = "  " & ![Cod]
                Ob_Grid.TextMatrix(XFila, 1) = ![DescV]
                Ob_Grid.TextMatrix(XFila, 2) = Format(![Mon], "##0.#0")
                Ob_Grid.TextMatrix(XFila, 4) = ![Dscto]
                Ob_Grid.TextMatrix(XFila, 5) = ![Sw_Aplic]
                Ob_Grid.TextMatrix(XFila, 6) = ![ID_PCta]
                Ob_Grid.TextMatrix(XFila, 7) = ![ID_PCta2]
                XFila = XFila + 1
                .MoveNext
           Loop
       End If
       .Close
    End With
    Set RstBol = Nothing
End Sub

Public Sub Cal_Rem()
  TotGen = 0: TotImp = 0
  For I = 1 To Grid1.Rows - 1
     If Trim(Grid1.TextMatrix(I, 2)) <> "" Then
        TotGen = TotGen + Val(Grid1.TextMatrix(I, 2))
     End If
     If Val(Grid1.TextMatrix(I, 0)) <> 10 And Val(Grid1.TextMatrix(I, 0)) <> 12 Then
        TotImp = TotImp + Val(Grid1.TextMatrix(I, 2))
     End If
  Next
   MBTot_R2 = Format(TotImp, "##,##0.#0")
   
   MBTot_R = Format(TotGen, "##,##0.#0")
   MBTot_P = Format(MBTot_R - MBTot_D, "##,##0.#0")
End Sub

Public Sub Cal_DDT()
    TotGen = 0
    For I = 1 To Grid2.Rows - 1
        If Val(Grid2.TextMatrix(I, 4)) > 0 Then
           'Grid2.TextMatrix(I, 2) = Format((CDbl(MBTot_R) * (Val(Grid2.TextMatrix(I, 4)) / 100)), "##,##0.#0")
           TotGen = TotGen + Val(Grid2.TextMatrix(I, 2))
        Else
           TotGen = TotGen + Val(Grid2.TextMatrix(I, 2))
        End If
    Next
    MBTot_D = Format(TotGen, "##,##0.#0")
    MBTot_P = Format(MBTot_R - TotGen, "##,##0.#0")
End Sub
Public Sub Cal_AE()
    TotGen = 0
    For I = 1 To Grid3.Rows - 1
        If Val(Grid3.TextMatrix(I, 4)) > 0 And Val(Grid3.TextMatrix(I, 5)) = 1 Then
           'Grid3.TextMatrix(I, 2) = Format((CDbl(MBTot_R) * (Val(Grid3.TextMatrix(I, 4)) / 100)), "##,##0.#0")
           TotGen = TotGen + Val(Grid3.TextMatrix(I, 2))
        End If
    Next
   MBTot_A = Format(TotGen, "##,##0.#0")
End Sub

Private Sub SSCommand2_Click()
    If CmdCerrar.Enabled = True Then
       If Val(MbNAsie) <> 0 Then
          CodGene = Right(CmbAno.Text, 2) & Right("0" & CmbMes1.BoundText, 2) '     Format(DtpFecha, "yy") 'Año y Mes
          StrAsie = MbNAsie  ' Numero de Asiento
          TotGenD = 43 ' tipo de documentos
          TotImpD = 34 ' Tipo de Contables
          Rptldia2.Show
       End If
    End If
End Sub
