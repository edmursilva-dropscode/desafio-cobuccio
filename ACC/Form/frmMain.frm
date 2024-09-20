VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Administradora de Cartões de Crédito"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   19305
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMain"
   ScaleHeight     =   10125
   ScaleWidth      =   19305
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox mainBorder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   30
      ScaleHeight     =   8655
      ScaleWidth      =   19965
      TabIndex        =   0
      Top             =   375
      Width           =   19965
      Begin VB.PictureBox MenuItem 
         Height          =   7455
         Index           =   1
         Left            =   4830
         ScaleHeight     =   7395
         ScaleWidth      =   2925
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   2985
         Begin VB.Frame fmeTransacao 
            Height          =   600
            Left            =   75
            TabIndex        =   8
            Top             =   270
            Width           =   16230
            Begin VB.Frame fmeData 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   8475
               TabIndex        =   18
               Top             =   135
               Visible         =   0   'False
               Width           =   4575
               Begin VB.CommandButton cmdFiltrar 
                  Caption         =   "&Filtrar"
                  Height          =   315
                  Left            =   3435
                  TabIndex        =   21
                  Top             =   30
                  Width           =   705
               End
               Begin MSMask.MaskEdBox mkbDataInicial 
                  Height          =   315
                  Left            =   870
                  TabIndex        =   11
                  Top             =   45
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mkbDataFinal 
                  Height          =   315
                  Left            =   2310
                  TabIndex        =   12
                  Top             =   45
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptChar      =   "_"
               End
               Begin VB.Label lblData 
                  AutoSize        =   -1  'True
                  Caption         =   "Data:"
                  Height          =   195
                  Left            =   390
                  TabIndex        =   20
                  Top             =   120
                  Width           =   390
               End
               Begin VB.Label lblAte 
                  AutoSize        =   -1  'True
                  Caption         =   "até"
                  Height          =   195
                  Left            =   1935
                  TabIndex        =   19
                  Top             =   120
                  Width           =   225
               End
            End
            Begin VB.ComboBox cmbLocalizarTransacao 
               Height          =   315
               ItemData        =   "frmMain.frx":0000
               Left            =   3915
               List            =   "frmMain.frx":0010
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   180
               Width           =   1680
            End
            Begin VB.TextBox txtLocalizarTransacao 
               Height          =   315
               Left            =   840
               MaxLength       =   20
               TabIndex        =   9
               Top             =   180
               Width           =   2475
            End
            Begin MSAdodcLib.Adodc adoTransacao 
               Height          =   330
               Left            =   6075
               Top             =   195
               Visible         =   0   'False
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   582
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   1
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AdministradoraCC"
               OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AdministradoraCC"
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   $"frmMain.frx":003C
               Caption         =   "adoTransacao"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin VB.Label lblTipo 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   3450
               TabIndex        =   14
               Top             =   255
               Width           =   360
            End
            Begin VB.Label lblLocalizar 
               AutoSize        =   -1  'True
               Caption         =   "Localizar:"
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   255
               Width           =   675
            End
         End
         Begin MSDataGridLib.DataGrid dtgTransacao 
            Bindings        =   "frmMain.frx":016B
            Height          =   6435
            Left            =   60
            TabIndex        =   15
            Top             =   900
            Width           =   16215
            _ExtentX        =   28601
            _ExtentY        =   11351
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "ID_Transacao"
               Caption         =   "ID_Transacao"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "ID_Cliente"
               Caption         =   "ID_Cliente"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Nome_Cliente"
               Caption         =   "Cliente"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Numero_Cartao"
               Caption         =   "            No do cartão"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Valor_Transacao"
               Caption         =   "            Valor"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "Data_Transacao"
               Caption         =   "        Data"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "Descricao"
               Caption         =   "Descrição"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "Categoria"
               Caption         =   "    Categoria"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1035,213
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   3734,929
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   2190,047
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1590,236
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column06 
                  Locked          =   -1  'True
                  ColumnWidth     =   3734,929
               EndProperty
               BeginProperty Column07 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1140,095
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc adoComandas 
            Height          =   330
            Left            =   16470
            Top             =   1290
            Visible         =   0   'False
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TESTE_VB6"
            OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TESTE_VB6"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   $"frmMain.frx":0186
            Caption         =   "adoComandas"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Image imgItem 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   1
            Left            =   1245
            Picture         =   "frmMain.frx":0638
            Top             =   120
            Width           =   210
         End
         Begin VB.Label lblTitulo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transações"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   1
            Left            =   105
            TabIndex        =   16
            Top             =   60
            Width           =   1020
         End
      End
      Begin VB.PictureBox MenuItem 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Index           =   2
         Left            =   8220
         ScaleHeight     =   7395
         ScaleWidth      =   3030
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   3090
         Begin MSDataGridLib.DataGrid dtgEstatisticas 
            Bindings        =   "frmMain.frx":098A
            Height          =   6435
            Left            =   60
            TabIndex        =   33
            Top             =   885
            Width           =   16215
            _ExtentX        =   28601
            _ExtentY        =   11351
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "Numero_Cartao"
               Caption         =   "No do cartão"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Nome_Cliente"
               Caption         =   "Cliente"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Valor_Total"
               Caption         =   "Valor Total"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Quantidade_Transacoes"
               Caption         =   "Qtde. de transações"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   1590,236
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   3734,929
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1769,953
               EndProperty
            EndProperty
         End
         Begin VB.Frame fmeEstatisticas 
            Height          =   600
            Left            =   60
            TabIndex        =   22
            Top             =   255
            Width           =   16230
            Begin VB.ComboBox cmbLocalizarEstatisticas 
               Height          =   315
               ItemData        =   "frmMain.frx":09A8
               Left            =   4080
               List            =   "frmMain.frx":09B2
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   180
               Width           =   1680
            End
            Begin MSAdodcLib.Adodc adoEstatisticas 
               Height          =   330
               Left            =   10365
               Top             =   180
               Visible         =   0   'False
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   582
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   1
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AdministradoraCC"
               OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=AdministradoraCC"
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   $"frmMain.frx":09C5
               Caption         =   "adoEstatisticas"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin VB.TextBox txtLocalizarEstatisticas 
               Height          =   315
               Left            =   840
               MaxLength       =   20
               TabIndex        =   27
               Top             =   180
               Width           =   2475
            End
            Begin VB.Frame fmeDataEstatistica 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   5940
               TabIndex        =   23
               Top             =   135
               Visible         =   0   'False
               Width           =   4575
               Begin VB.CommandButton cmdFiltrarEstatisticas 
                  Caption         =   "&Filtrar"
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   24
                  Top             =   30
                  Width           =   705
               End
               Begin MSMask.MaskEdBox mkbDataInicialEstatistica 
                  Height          =   315
                  Left            =   870
                  TabIndex        =   30
                  Top             =   45
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mkbDataFinalEstatistica 
                  Height          =   315
                  Left            =   2310
                  TabIndex        =   32
                  Top             =   45
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptChar      =   "_"
               End
               Begin VB.Label lblAtéEstatisticas 
                  AutoSize        =   -1  'True
                  Caption         =   "até"
                  Height          =   195
                  Left            =   1935
                  TabIndex        =   26
                  Top             =   120
                  Width           =   225
               End
               Begin VB.Label lblDataEstatisticas 
                  AutoSize        =   -1  'True
                  Caption         =   "Data:"
                  Height          =   195
                  Left            =   390
                  TabIndex        =   25
                  Top             =   120
                  Width           =   390
               End
            End
            Begin VB.Label lblTipoEstatistica 
               AutoSize        =   -1  'True
               Caption         =   "Tipo:"
               Height          =   195
               Left            =   3570
               TabIndex        =   34
               Top             =   225
               Width           =   360
            End
            Begin VB.Label lblLocalizarEstatisticas 
               AutoSize        =   -1  'True
               Caption         =   "Localizar:"
               Height          =   195
               Left            =   105
               TabIndex        =   29
               Top             =   255
               Width           =   675
            End
         End
         Begin MSAdodcLib.Adodc adoCozinhaItens 
            Height          =   330
            Left            =   18555
            Top             =   6225
            Visible         =   0   'False
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TESTE_VB6"
            OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TESTE_VB6"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   $"frmMain.frx":0AE8
            Caption         =   "adoCozinhaItens"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Image imgItem 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   2
            Left            =   1215
            Picture         =   "frmMain.frx":10BE
            Top             =   105
            Width           =   210
         End
         Begin VB.Label lblTitulo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estatísticas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   75
            TabIndex        =   31
            Top             =   60
            Width           =   1005
         End
      End
      Begin VB.PictureBox MenuItem 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Index           =   0
         Left            =   1785
         ScaleHeight     =   7395
         ScaleWidth      =   2505
         TabIndex        =   4
         Top             =   0
         Width           =   2565
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   17385
            Left            =   0
            Picture         =   "frmMain.frx":1410
            ScaleHeight     =   17385
            ScaleWidth      =   29085
            TabIndex        =   5
            Top             =   -60
            Width           =   29085
         End
      End
      Begin VB.CommandButton cmdEstatisticas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estatisticas"
         Height          =   1290
         Left            =   120
         Picture         =   "frmMain.frx":37B8E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1530
         Width           =   1275
      End
      Begin VB.CommandButton cmdTransacao 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Transações"
         Height          =   1290
         Left            =   135
         Picture         =   "frmMain.frx":38291
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   135
         Width           =   1275
      End
   End
   Begin MSComctlLib.ImageList imlBotões 
      Left            =   13620
      Top             =   9240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38A13
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38B6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38CCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38E27
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38F83
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":390DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":391F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3934F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":394AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39607
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39BA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39CFF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Barra 
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   21195
      _ExtentX        =   37386
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlBotões"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoNovo"
            Object.ToolTipText     =   "Novo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoBarra01"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoImprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoEditar"
            Object.ToolTipText     =   "Editar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoExcluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoBarra02"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoAgrupar"
            Object.ToolTipText     =   "Agrupar"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoDetalhes"
            Object.ToolTipText     =   "Exibir Detalhes"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btoAtualizar"
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoClassificar"
            Object.ToolTipText     =   "Classificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btoCalculadora"
            Object.ToolTipText     =   "Calculadora"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoConfig"
            Object.ToolTipText     =   "Configurações"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "btoLogin"
            Object.ToolTipText     =   "Login"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "btoAjuda"
            Object.ToolTipText     =   "Ajuda"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   9855
      Width           =   19305
      _ExtentX        =   34052
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25347
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "19/09/2024"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "19:47"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuSairSistema 
         Caption         =   "&Sair do Sistema"
      End
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuCliente 
         Caption         =   "&Cliente"
      End
   End
   Begin VB.Menu mnuPopTransacao 
      Caption         =   "mnuPopTransacao"
      Visible         =   0   'False
      Begin VB.Menu mnuNovaTransacao 
         Caption         =   "&Nova transação"
      End
      Begin VB.Menu mnuRelatorioTransacao 
         Caption         =   "&Relatório Excel"
      End
   End
   Begin VB.Menu mnuPopEstatistica 
      Caption         =   "mnuPopEstatistica"
      Visible         =   0   'False
      Begin VB.Menu mnuRelatorioEstatistica 
         Caption         =   "&Relatório Excel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variável de acesso as classes
Private vop_TransacaoNegocios As New clsTransacaoNegocios
Private vop_EstatisticaNegocios As New clsEstatisticaNegocios
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'Variaveis de controle do form
Private vil_IdTransicao As Long                             'Identificador da Transicao
Private vil_NumeroCartao As String                          ' Variavel de controle da estatistica
Private bMovingProgrammatically As Boolean                  ' Variável de controle para evitar movimentação dupla
Private vil_LastBookmark As Variant                         ' Variável de módulo para armazenar o último Bookmark




'Form
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo TrataErros

    'Tecla de atalho da calculadora
    If KeyCode = vbKeyF7 Then
        KeyCode = 0
        Exit Sub
    End If
    'Tecla de sair do form
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
TrataErros:
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo TrataErros
    If MsgBox(MSG01, Style10, Title01) = vbYes Then
       Set frmMain = Nothing
       End
    Else
      Cancel = 1
    End If
TrataErros:
    If Err.Number = 3420 Then End
   
End Sub

Private Sub Form_Resize()

   Call HabilitaMenuItem
    
End Sub

Private Sub mnuSairSistema_Click()
On Error GoTo TrataErros
    If MsgBox(MSG01, Style10, Title01) = vbYes Then
       Set frmMain = Nothing
       End
    End If
TrataErros:
    If Err.Number = 3420 Then End
End Sub

Private Sub imgItem_Click(Index As Integer)
    Select Case Index
        Case 1
            PopupMenu mnuPopTransacao, , Screen.TwipsPerPixelX + 1600, Screen.TwipsPerPixelY + 740
        Case 2
            PopupMenu mnuPopEstatistica, , Screen.TwipsPerPixelX + 1600, Screen.TwipsPerPixelY + 740
    End Select
End Sub

Private Sub Barra_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
         Case "btoAtualizar"
            If MenuItem(1).Visible = True Then
               Call CarregarGridTransacao
            End If
        Case "btoCalculadora"
            Call Calculadora
    End Select
End Sub





'Cliente
Private Sub mnuCliente_Click()
On Error GoTo TrataErros

    DoEvents
    Load frmListaClientes
    DoEvents
    frmListaClientes.Show vbModal
    
TrataErros:
    If Err.Number <> 0 Then TrataErros
    
End Sub





'Transacao
Private Sub mnuNovaTransacao_Click()
On Error GoTo TrataErro
    DoEvents
    Load frmTransacao
    DoEvents
    frmTransacao.Show vbModal
    '
    txtLocalizarTransacao.text = Empty
    cmbLocalizarTransacao.ListIndex = 0
    'Carregar Grid
    Call CarregarGridTransacao
TrataErro:
    If Err.Number <> 0 Then TrataErros
End Sub

Private Sub mnuRelatorioTransacao_Click()
On Error GoTo TrataErro
    
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    Dim i As Long, j As Long
    Dim caminho As String
    
    ' Abrir diálogo para escolher o local de salvamento
    caminho = ExibirDialogoSalvar()
    If caminho = "" Then Exit Sub ' Usuário cancelou
    
    ' Criar nova instância do Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets(1)
    
    ' Exportar cabeçalhos
    For j = 0 To dtgTransacao.Columns.Count - 1
        xlWorksheet.Cells(1, j + 1).Value = dtgTransacao.Columns(j).Caption
    Next j
    
    ' Exportar dados
    dtgTransacao.Row = 0
    For i = 0 To dtgTransacao.VisibleRows - 1
        dtgTransacao.Row = i
        For j = 0 To dtgTransacao.Columns.Count - 1
            dtgTransacao.Col = j
            xlWorksheet.Cells(i + 2, j + 1).Value = dtgTransacao.text
        Next j
    Next i
    
    ' Salvar e fechar
    xlWorkbook.SaveAs caminho
    xlWorkbook.Close
    xlApp.Quit
    
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    
    MsgBox "Exportação concluída com sucesso!", vbInformation
    
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then TrataErros
End Sub

Private Sub mkbDataInicial_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8  ' Backspace
            ' Permite o backspace
        Case 13  ' Enter
            KeyAscii = 0  ' Cancela o "beep"
            Sendkeys "{TAB}"  ' Simula pressionar a tecla Tab
        Case 48 To 57  ' Dígitos de 0 a 9
            ' Permite dígitos
        Case Else
            KeyAscii = 0  ' Cancela qualquer outro caractere
    End Select
End Sub

Private Sub mkbDataInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then
        Dim cursorPos As Integer
        Dim newPos As Integer
        Dim i As Integer
        
        cursorPos = mkbDataInicial.SelStart
        
        ' Se o cursor estiver no início, não faz nada
        If cursorPos = 0 Then Exit Sub
        
        ' Encontra a próxima posição não underscore à esquerda
        For i = cursorPos - 1 To 0 Step -1
            If Mid$(mkbDataInicial.text, i + 1, 1) <> "_" Then
                newPos = i
                Exit For
            End If
        Next i
        
        ' Se encontrou uma posição válida, apaga até lá
        If newPos >= 0 Then
            mkbDataInicial.SelStart = newPos
            mkbDataInicial.SelLength = cursorPos - newPos
            mkbDataInicial.SelText = String$(cursorPos - newPos, "_")
            mkbDataInicial.SelStart = newPos
        End If
        
        ' Cancela o keypress padrão
        KeyCode = 0
    End If
End Sub

Private Sub mkbDataInicial_Change()
    Dim cursorPos As Integer
    Dim cleanInput As String
    Dim formattedDate As String
    Dim i As Integer
    
    ' Guarda a posição atual do cursor
    cursorPos = mkbDataInicial.SelStart
    
    ' Remove todos os caracteres não numéricos
    cleanInput = ""
    For i = 1 To Len(mkbDataInicial.text)
        If IsNumeric(Mid$(mkbDataInicial.text, i, 1)) Then
            cleanInput = cleanInput & Mid$(mkbDataInicial.text, i, 1)
        End If
    Next i
    
    ' Formata a data
    formattedDate = ""
    For i = 1 To Len(cleanInput)
        formattedDate = formattedDate & Mid$(cleanInput, i, 1)
        If i = 2 Or i = 4 Then
            formattedDate = formattedDate & "/"
        End If
    Next i
    
    ' Preenche o restante com underscores
    While Len(formattedDate) < 10
        formattedDate = formattedDate & "_"
    Wend
    
    ' Atualiza o texto do TextBox
    mkbDataInicial.text = formattedDate
    
    ' Ajusta a posição do cursor
    If cursorPos > 0 Then
        If Mid$(formattedDate, cursorPos, 1) = "/" Then
            cursorPos = cursorPos + 1
        End If
        If cursorPos > Len(formattedDate) Then
            cursorPos = Len(formattedDate)
        End If
    End If
    mkbDataInicial.SelStart = cursorPos
End Sub

Private Sub mkbDataInicial_Validate(Cancel As Boolean)
    If Not IsDate(mkbDataInicial.text) And mkbDataInicial.text <> "__/__/____" Then
        If Len(Replace(mkbDataInicial.text, "_", "")) > 0 Then
           MsgBox "Data inválida. Por favor, verifique.", vbExclamation
           Cancel = True
        End If
    End If
End Sub

Private Sub mkbDataFinal_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8  ' Backspace
            ' Permite o backspace
        Case 13  ' Enter
            KeyAscii = 0  ' Cancela o "beep"
            Sendkeys "{TAB}"  ' Simula pressionar a tecla Tab
        Case 48 To 57  ' Dígitos de 0 a 9
            ' Permite dígitos
        Case Else
            KeyAscii = 0  ' Cancela qualquer outro caractere
    End Select
End Sub

Private Sub mkbDataFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then
        Dim cursorPos As Integer
        Dim newPos As Integer
        Dim i As Integer
        
        cursorPos = mkbDataFinal.SelStart
        
        ' Se o cursor estiver no início, não faz nada
        If cursorPos = 0 Then Exit Sub
        
        ' Encontra a próxima posição não underscore à esquerda
        For i = cursorPos - 1 To 0 Step -1
            If Mid$(mkbDataFinal.text, i + 1, 1) <> "_" Then
                newPos = i
                Exit For
            End If
        Next i
        
        ' Se encontrou uma posição válida, apaga até lá
        If newPos >= 0 Then
            mkbDataFinal.SelStart = newPos
            mkbDataFinal.SelLength = cursorPos - newPos
            mkbDataFinal.SelText = String$(cursorPos - newPos, "_")
            mkbDataFinal.SelStart = newPos
        End If
        
        ' Cancela o keypress padrão
        KeyCode = 0
    End If
End Sub

Private Sub mkbDataFinal_Change()
    Dim cursorPos As Integer
    Dim cleanInput As String
    Dim formattedDate As String
    Dim i As Integer
    
    ' Guarda a posição atual do cursor
    cursorPos = mkbDataFinal.SelStart
    
    ' Remove todos os caracteres não numéricos
    cleanInput = ""
    For i = 1 To Len(mkbDataFinal.text)
        If IsNumeric(Mid$(mkbDataFinal.text, i, 1)) Then
            cleanInput = cleanInput & Mid$(mkbDataFinal.text, i, 1)
        End If
    Next i
    
    ' Formata a data
    formattedDate = ""
    For i = 1 To Len(cleanInput)
        formattedDate = formattedDate & Mid$(cleanInput, i, 1)
        If i = 2 Or i = 4 Then
            formattedDate = formattedDate & "/"
        End If
    Next i
    
    ' Preenche o restante com underscores
    While Len(formattedDate) < 10
        formattedDate = formattedDate & "_"
    Wend
    
    ' Atualiza o texto do TextBox
    mkbDataFinal.text = formattedDate
    
    ' Ajusta a posição do cursor
    If cursorPos > 0 Then
        If Mid$(formattedDate, cursorPos, 1) = "/" Then
            cursorPos = cursorPos + 1
        End If
        If cursorPos > Len(formattedDate) Then
            cursorPos = Len(formattedDate)
        End If
    End If
    mkbDataFinal.SelStart = cursorPos
End Sub

Private Sub mkbDataFinal_Validate(Cancel As Boolean)
    If Not IsDate(mkbDataFinal.text) And mkbDataFinal.text <> "__/__/____" Then
        If Len(Replace(mkbDataFinal.text, "_", "")) > 0 Then
           MsgBox "Data inválida. Por favor, verifique.", vbExclamation
           Cancel = True
        End If
    End If
End Sub

Private Sub cmdFiltrar_Click()
   Call CarregarGridTransacao
End Sub

Private Sub txtLocalizarTransacao_Change()

On Error GoTo TrataErros

   Set vop_TransacaoNegocios = New clsTransacaoNegocios
       Call vop_TransacaoNegocios.LocalizarTransacao(dtgTransacao, adoTransacao, txtLocalizarTransacao.text, cmbLocalizarTransacao.ListIndex, Replace(mkbDataInicial.text, "_", ""), Replace(mkbDataFinal.text, "_", ""))
   Set vop_TransacaoNegocios = Nothing
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_TransacaoNegocios = Nothing
       Exit Sub
    End If
End Sub

Private Sub cmbLocalizarTransacao_Click()

On Error GoTo TrataErros

   If cmbLocalizarTransacao.ListIndex = 3 Then
      fmeData.Visible = True
   Else
      fmeData.Visible = False
      mkbDataInicial.text = Empty
      mkbDataFinal.text = Empty
   End If
   
   Set vop_TransacaoNegocios = New clsTransacaoNegocios
       Call vop_TransacaoNegocios.LocalizarTransacao(dtgTransacao, adoTransacao, txtLocalizarTransacao.text, cmbLocalizarTransacao.ListIndex, Replace(mkbDataInicial.text, "_", ""), Replace(mkbDataFinal.text, "_", ""))
   Set vop_TransacaoNegocios = Nothing
    
   txtLocalizarTransacao.text = Empty
   
TrataErros:
    If Err.Number <> 0 Then
       Set vop_TransacaoNegocios = Nothing
       Exit Sub
    End If
End Sub

Private Sub cmdTransacao_Click()
   '
   Call CarregarGridTransacao
   cmbLocalizarTransacao.ListIndex = 0
   'Transacao
   Call HabilitaMenuItem
   MenuItem(0).Visible = False
   MenuItem(1).Visible = True
   MenuItem(2).Visible = False
   fmeTransacao.Width = mainBorder.Width - 1630 - 90
   dtgTransacao.Width = mainBorder.Width - 1630 - 110
   dtgTransacao.Height = mainBorder.Height - 960
   
   ' Inicia foco
   txtLocalizarTransacao.SetFocus
   
End Sub

Private Sub dtgTransacao_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' Só atualiza se não estiver movendo programaticamente
    If Not bMovingProgrammatically Then
        If dtgTransacao.Bookmark > 0 Then
            If dtgTransacao.SelBookmarks.Count > 0 Then
                dtgTransacao.SelBookmarks.Remove 0
            End If
            dtgTransacao.SelBookmarks.Add dtgTransacao.Bookmark
            If IsNull(dtgTransacao.Columns(0).text) Or dtgTransacao.Columns(0).text = "" Then
                vil_IdTransicao = 0
            Else
                vil_IdTransicao = CInt(dtgTransacao.Columns(0).text)
            End If
        End If
    End If
End Sub

Private Sub dtgTransacao_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            MoveRecordT adoTransacao.Recordset, False
            'Valida fim do DataGrid
        Case vbKeyUp
            MoveRecordT adoTransacao.Recordset, True
    End Select
End Sub

Private Sub MoveRecordT(rs As ADODB.Recordset, MovePrevious As Boolean)
    If rs.RecordCount > 0 Then
        bMovingProgrammatically = True
        
        If dtgTransacao.SelBookmarks.Count > 0 Then
            dtgTransacao.SelBookmarks.Remove 0
        End If
        
        If MovePrevious Then
            rs.MovePrevious
            If rs.BOF Then
                rs.MoveFirst
                dtgTransacao.SelBookmarks.Add dtgTransacao.Bookmark
            Else
                'rs.MoveNext
                dtgTransacao.SelBookmarks.Add dtgTransacao.Bookmark
            End If
        Else
            rs.MoveNext
            If rs.EOF Then
                rs.MoveLast
                dtgTransacao.SelBookmarks.Add dtgTransacao.Bookmark
            Else
               'rs.MovePrevious
               dtgTransacao.SelBookmarks.Add dtgTransacao.Bookmark
            End If
        End If
        
        dtgTransacao.Bookmark = rs.Bookmark
        
        bMovingProgrammatically = False ' Reseta o flag
    End If
End Sub

Private Sub dtgTransacao_DblClick()
Dim vvl_BookMark As Variant
Dim vil_RowIndex As Long

On Error GoTo TrataErros

    vil_RowIndex = dtgTransacao.Row
    vvl_BookMark = dtgTransacao.RowBookmark(vil_RowIndex)
    If vvl_BookMark = Empty Then Exit Sub

    Call frmTransacao.Form_Load
    Call frmTransacao.Editar(vil_IdTransicao)

    dtgTransacao.Bookmark = vvl_BookMark
    dtgTransacao.Scroll 0, dtgTransacao.RowContaining(vvl_BookMark)
    dtgTransacao.SelBookmarks.Add vvl_BookMark
    dtgTransacao.Refresh

TrataErros:
    If Err.Number <> 0 Then Exit Sub
End Sub

Private Sub mnuTransacao_Click()
On Error GoTo TrataErros
   
    'Inicia form
    frmTransacao.Show vbModal
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_TransacaoNegocios = Nothing
       Exit Sub
    End If
End Sub




'Estatiscica
Private Sub mnuRelatorioEstatistica_Click()
On Error GoTo TrataErro
    
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    Dim i As Long, j As Long
    Dim caminho As String
    
    ' Abrir diálogo para escolher o local de salvamento
    caminho = ExibirDialogoSalvar()
    If caminho = "" Then Exit Sub ' Usuário cancelou
    
    ' Criar nova instância do Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets(1)
    
    ' Exportar cabeçalhos
    For j = 0 To dtgEstatisticas.Columns.Count - 1
        xlWorksheet.Cells(1, j + 1).Value = dtgEstatisticas.Columns(j).Caption
    Next j
    
    ' Exportar dados
    dtgEstatisticas.Row = 0
    For i = 0 To dtgEstatisticas.VisibleRows - 1
        dtgEstatisticas.Row = i
        For j = 0 To dtgEstatisticas.Columns.Count - 1
            dtgEstatisticas.Col = j
            xlWorksheet.Cells(i + 2, j + 1).Value = dtgEstatisticas.text
        Next j
    Next i
    
    ' Salvar e fechar
    xlWorkbook.SaveAs caminho
    xlWorkbook.Close
    xlApp.Quit
    
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    
    MsgBox "Exportação concluída com sucesso!", vbInformation
    
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then TrataErros
End Sub

Private Sub cmdEstatisticas_Click()
   
   '
   Call CarregarGridEstatistica
   cmbLocalizarEstatisticas.ListIndex = 0
   'Estatistica
   Call HabilitaMenuItem
   MenuItem(0).Visible = False
   MenuItem(1).Visible = False
   MenuItem(2).Visible = True
   fmeEstatisticas.Width = mainBorder.Width - 1630 - 90
   dtgEstatisticas.Width = mainBorder.Width - 1630 - 110
   dtgEstatisticas.Height = mainBorder.Height - 960
   
   ' Inicia foco
   txtLocalizarEstatisticas.SetFocus
   
   
End Sub

Private Sub dtgEstatisticas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' Só atualiza se não estiver movendo programaticamente
    If Not bMovingProgrammatically Then
        If dtgEstatisticas.Bookmark > 0 Then
            If dtgEstatisticas.SelBookmarks.Count > 0 Then
                dtgEstatisticas.SelBookmarks.Remove 0
            End If
            dtgEstatisticas.SelBookmarks.Add dtgEstatisticas.Bookmark
            If IsNull(dtgEstatisticas.Columns(0).text) Or dtgEstatisticas.Columns(0).text = "" Then
                vil_NumeroCartao = "0"
            Else
                vil_NumeroCartao = dtgEstatisticas.Columns(0).text
            End If
        End If
    End If
End Sub

Private Sub dtgEstatisticas_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            MoveRecordE adoTransacao.Recordset, False
            'Valida fim do DataGrid
        Case vbKeyUp
            MoveRecordE adoTransacao.Recordset, True
    End Select
End Sub

Private Sub MoveRecordE(rs As ADODB.Recordset, MovePrevious As Boolean)
    'If rs.RecordCount > 0 Then
        bMovingProgrammatically = True
        
        If dtgEstatisticas.SelBookmarks.Count > 0 Then
            dtgEstatisticas.SelBookmarks.Remove 0
        End If
        
        If MovePrevious Then
            'rs.MovePrevious
            'If rs.BOF Then
                'rs.MoveFirst
                dtgEstatisticas.SelBookmarks.Add dtgEstatisticas.Bookmark
            'Else
                'rs.MoveNext
                'dtgEstatisticas.SelBookmarks.Add dtgEstatisticas.Bookmark
            'End If
        Else
            'rs.MoveNext
            'If rs.EOF Then
                'rs.MoveLast
                dtgEstatisticas.SelBookmarks.Add dtgEstatisticas.Bookmark
            'Else
               'rs.MovePrevious
               dtgEstatisticas.SelBookmarks.Add dtgEstatisticas.Bookmark
            'End If
        End If
        
        'dtgEstatisticas.Bookmark = rs.Bookmark
        '
        bMovingProgrammatically = False ' Reseta o flag
    'End If
End Sub

Private Sub txtLocalizarEstatisticas_Change()

On Error GoTo TrataErros

   Set vop_EstatisticaNegocios = New clsEstatisticaNegocios
       Call vop_EstatisticaNegocios.LocalizarEstatisticas(dtgEstatisticas, adoEstatisticas, txtLocalizarEstatisticas.text, cmbLocalizarEstatisticas.ListIndex, Replace(mkbDataInicialEstatistica.text, "_", ""), Replace(mkbDataFinalEstatistica.text, "_", ""))
   Set vop_EstatisticaNegocios = Nothing
       
TrataErros:
    If Err.Number <> 0 Then
       Set vop_TransacaoNegocios = Nothing
       Exit Sub
    End If

End Sub

Private Sub cmbLocalizarEstatisticas_Click()

On Error GoTo TrataErros

   If cmbLocalizarEstatisticas.ListIndex = 1 Then
      fmeDataEstatistica.Visible = True
   Else
      fmeDataEstatistica.Visible = False
      mkbDataInicialEstatistica.text = Empty
      mkbDataFinalEstatistica.text = Empty
   End If
   
   Set vop_EstatisticaNegocios = New clsEstatisticaNegocios
       Call vop_EstatisticaNegocios.LocalizarEstatisticas(dtgEstatisticas, adoEstatisticas, txtLocalizarEstatisticas.text, cmbLocalizarEstatisticas.ListIndex, Replace(mkbDataInicialEstatistica.text, "_", ""), Replace(mkbDataFinalEstatistica.text, "_", ""))
   Set vop_EstatisticaNegocios = Nothing
    
   txtLocalizarEstatisticas.text = Empty
   
TrataErros:
    If Err.Number <> 0 Then
       Set vop_EstatisticaNegocios = Nothing
       Exit Sub
    End If

End Sub

Private Sub mkbDataInicialEstatistica_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8  ' Backspace
            ' Permite o backspace
        Case 13  ' Enter
            KeyAscii = 0  ' Cancela o "beep"
            Sendkeys "{TAB}"  ' Simula pressionar a tecla Tab
        Case 48 To 57  ' Dígitos de 0 a 9
            ' Permite dígitos
        Case Else
            KeyAscii = 0  ' Cancela qualquer outro caractere
    End Select
End Sub

Private Sub mkbDataInicialEstatistica_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then
        Dim cursorPos As Integer
        Dim newPos As Integer
        Dim i As Integer
        
        cursorPos = mkbDataInicialEstatistica.SelStart
        
        ' Se o cursor estiver no início, não faz nada
        If cursorPos = 0 Then Exit Sub
        
        ' Encontra a próxima posição não underscore à esquerda
        For i = cursorPos - 1 To 0 Step -1
            If Mid$(mkbDataInicialEstatistica.text, i + 1, 1) <> "_" Then
                newPos = i
                Exit For
            End If
        Next i
        
        ' Se encontrou uma posição válida, apaga até lá
        If newPos >= 0 Then
            mkbDataInicialEstatistica.SelStart = newPos
            mkbDataInicialEstatistica.SelLength = cursorPos - newPos
            mkbDataInicialEstatistica.SelText = String$(cursorPos - newPos, "_")
            mkbDataInicialEstatistica.SelStart = newPos
        End If
        
        ' Cancela o keypress padrão
        KeyCode = 0
    End If
End Sub

Private Sub mkbDataInicialEstatistica_Change()
    Dim cursorPos As Integer
    Dim cleanInput As String
    Dim formattedDate As String
    Dim i As Integer
    
    ' Guarda a posição atual do cursor
    cursorPos = mkbDataInicialEstatistica.SelStart
    
    ' Remove todos os caracteres não numéricos
    cleanInput = ""
    For i = 1 To Len(mkbDataInicialEstatistica.text)
        If IsNumeric(Mid$(mkbDataInicialEstatistica.text, i, 1)) Then
            cleanInput = cleanInput & Mid$(mkbDataInicialEstatistica.text, i, 1)
        End If
    Next i
    
    ' Formata a data
    formattedDate = ""
    For i = 1 To Len(cleanInput)
        formattedDate = formattedDate & Mid$(cleanInput, i, 1)
        If i = 2 Or i = 4 Then
            formattedDate = formattedDate & "/"
        End If
    Next i
    
    ' Preenche o restante com underscores
    While Len(formattedDate) < 10
        formattedDate = formattedDate & "_"
    Wend
    
    ' Atualiza o texto do TextBox
    mkbDataInicialEstatistica.text = formattedDate
    
    ' Ajusta a posição do cursor
    If cursorPos > 0 Then
        If Mid$(formattedDate, cursorPos, 1) = "/" Then
            cursorPos = cursorPos + 1
        End If
        If cursorPos > Len(formattedDate) Then
            cursorPos = Len(formattedDate)
        End If
    End If
    mkbDataInicialEstatistica.SelStart = cursorPos
End Sub

Private Sub mkbDataInicialEstatistica_Validate(Cancel As Boolean)
    If Not IsDate(mkbDataInicialEstatistica.text) And mkbDataInicialEstatistica.text <> "__/__/____" Then
        If Len(Replace(mkbDataInicialEstatistica.text, "_", "")) > 0 Then
           MsgBox "Data inválida. Por favor, verifique.", vbExclamation
           Cancel = True
        End If
    End If
End Sub

Private Sub mkbDataFinalEstatistica_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8  ' Backspace
            ' Permite o backspace
        Case 13  ' Enter
            KeyAscii = 0  ' Cancela o "beep"
            Sendkeys "{TAB}"  ' Simula pressionar a tecla Tab
        Case 48 To 57  ' Dígitos de 0 a 9
            ' Permite dígitos
        Case Else
            KeyAscii = 0  ' Cancela qualquer outro caractere
    End Select
End Sub

Private Sub mkbDataFinalEstatistica_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then
        Dim cursorPos As Integer
        Dim newPos As Integer
        Dim i As Integer
        
        cursorPos = mkbDataFinalEstatistica.SelStart
        
        ' Se o cursor estiver no início, não faz nada
        If cursorPos = 0 Then Exit Sub
        
        ' Encontra a próxima posição não underscore à esquerda
        For i = cursorPos - 1 To 0 Step -1
            If Mid$(mkbDataFinalEstatistica.text, i + 1, 1) <> "_" Then
                newPos = i
                Exit For
            End If
        Next i
        
        ' Se encontrou uma posição válida, apaga até lá
        If newPos >= 0 Then
            mkbDataFinalEstatistica.SelStart = newPos
            mkbDataFinalEstatistica.SelLength = cursorPos - newPos
            mkbDataFinalEstatistica.SelText = String$(cursorPos - newPos, "_")
            mkbDataFinalEstatistica.SelStart = newPos
        End If
        
        ' Cancela o keypress padrão
        KeyCode = 0
    End If
End Sub

Private Sub mkbDataFinalEstatistica_Change()
    Dim cursorPos As Integer
    Dim cleanInput As String
    Dim formattedDate As String
    Dim i As Integer
    
    ' Guarda a posição atual do cursor
    cursorPos = mkbDataFinalEstatistica.SelStart
    
    ' Remove todos os caracteres não numéricos
    cleanInput = ""
    For i = 1 To Len(mkbDataFinalEstatistica.text)
        If IsNumeric(Mid$(mkbDataFinalEstatistica.text, i, 1)) Then
            cleanInput = cleanInput & Mid$(mkbDataFinalEstatistica.text, i, 1)
        End If
    Next i
    
    ' Formata a data
    formattedDate = ""
    For i = 1 To Len(cleanInput)
        formattedDate = formattedDate & Mid$(cleanInput, i, 1)
        If i = 2 Or i = 4 Then
            formattedDate = formattedDate & "/"
        End If
    Next i
    
    ' Preenche o restante com underscores
    While Len(formattedDate) < 10
        formattedDate = formattedDate & "_"
    Wend
    
    ' Atualiza o texto do TextBox
    mkbDataFinalEstatistica.text = formattedDate
    
    ' Ajusta a posição do cursor
    If cursorPos > 0 Then
        If Mid$(formattedDate, cursorPos, 1) = "/" Then
            cursorPos = cursorPos + 1
        End If
        If cursorPos > Len(formattedDate) Then
            cursorPos = Len(formattedDate)
        End If
    End If
    mkbDataFinalEstatistica.SelStart = cursorPos
End Sub

Private Sub mkbDataFinalEstatistica_Validate(Cancel As Boolean)
    If Not IsDate(mkbDataFinalEstatistica.text) And mkbDataFinalEstatistica.text <> "__/__/____" Then
        If Len(Replace(mkbDataFinalEstatistica.text, "_", "")) > 0 Then
           MsgBox "Data inválida. Por favor, verifique.", vbExclamation
           Cancel = True
        End If
    End If
End Sub

Private Sub cmdFiltrarEstatisticas_Click()
   Call CarregarGridEstatistica
End Sub









'Function
Private Function HabilitaMenuItem()
Dim vil_Count As Integer

    If Me.WindowState = vbMaximized Then
        Barra.Refresh
        mainBorder.Height = Me.Height - 1330
        mainBorder.Width = Me.Width
        Barra.Width = Me.Width
        For vil_Count = 0 To MenuItem.Count - 1
            MenuItem(vil_Count).Visible = False
            MenuItem(vil_Count).Left = 1560
            MenuItem(vil_Count).Top = 0
            MenuItem(vil_Count).Height = mainBorder.Height
            MenuItem(vil_Count).Width = mainBorder.Width - 1630
        Next
        MenuItem(0).Visible = True
    End If

End Function

Private Function ExibeMenuCheck(ByVal mnuBar, ByRef pmnuBar) As Integer

     mnuTodosOsItens.Checked = IIf(mnuBar = 0, True, False)
     mnuItemAguardandoEnvio.Checked = IIf(mnuBar = 1, True, False)
     mnuItemAguardandoProcessamento.Checked = IIf(mnuBar = 2, True, False)
     mnuItemSendoPreparado.Checked = IIf(mnuBar = 3, True, False)
     mnuItemParaEntrega.Checked = IIf(mnuBar = 4, True, False)
     mnuItemCancelado.Checked = IIf(mnuBar = 5, True, False)
     mnuItemEntregue.Checked = IIf(mnuBar = 6, True, False)
     
     lblFiltro.Caption = IIf(mnuBar = 0, mnuTodosOsItens.Caption, IIf(mnuBar = 1, mnuItemAguardandoEnvio.Caption, IIf(mnuBar = 2, mnuItemAguardandoProcessamento.Caption, IIf(mnuBar = 3, mnuItemSendoPreparado.Caption, IIf(mnuBar = 4, mnuItemParaEntrega.Caption, IIf(mnuBar = 4, mnuItemCancelado.Caption, mnuItemEntregue.Caption))))))
     ExibeMenuCheck = mnuBar
     pmnuBar = mnuBar
        
End Function

Public Function CarregarGridTransacao()
Dim vbl_Carregar As Boolean
    
On Error GoTo TrataErros

    Set vop_TransacaoNegocios = New clsTransacaoNegocios
                
        vbl_Carregar = vop_TransacaoNegocios.CarregarGridTransacaoMainRS(dtgTransacao, adoTransacao, txtLocalizarTransacao.text, IIf(cmbLocalizarTransacao.ListIndex < 0, 0, cmbLocalizarTransacao.ListIndex), Replace(mkbDataInicial.text, "_", ""), Replace(mkbDataFinal.text, "_", ""))
        If vbl_Carregar = True And adoTransacao.Recordset.Bookmark = 0 Then
           dtgTransacao.Refresh
        End If
               
    Set vop_TransacaoNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_TransacaoNegocios = Nothing
       Exit Function
    End If

End Function

Public Function CarregarGridEstatistica()
Dim vbl_Carregar As Boolean
    
On Error GoTo TrataErros

    Set vop_EstatisticaNegocios = New clsEstatisticaNegocios
                
        vbl_Carregar = vop_EstatisticaNegocios.CarregarGridEstatisticaMainRS(dtgEstatisticas, adoEstatisticas, txtLocalizarEstatisticas.text, IIf(cmbLocalizarEstatisticas.ListIndex < 0, 0, cmbLocalizarTransacao.ListIndex), Replace(mkbDataInicialEstatistica.text, "_", ""), Replace(mkbDataFinalEstatistica.text, "_", ""))
        If vbl_Carregar = True And adoEstatisticas.Recordset.Bookmark = 0 Then
           dtgEstatisticas.Refresh
        End If
               
    Set vop_EstatisticaNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_EstatisticaNegocios = Nothing
       Exit Function
    End If

End Function

Private Function ExibirDialogoSalvar() As String
    Dim ofn As OPENFILENAME
    Dim lReturn As Long
    Dim sFilter As String
    
    sFilter = "Arquivos Excel (*.xlsx)" & Chr(0) & "*.xlsx" & Chr(0) & "Todos os arquivos (*.*)" & Chr(0) & "*.*" & Chr(0)
    
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = Me.hWnd  ' Se estiver usando em um módulo, remova esta linha
        .hInstance = App.hInstance
        .lpstrFilter = sFilter
        .nFilterIndex = 1
        .lpstrFile = String(257, 0)
        .nMaxFile = Len(.lpstrFile) - 1
        .lpstrFileTitle = .lpstrFile
        .nMaxFileTitle = .nMaxFile
        .lpstrInitialDir = CurDir
        .lpstrTitle = "Salvar arquivo Excel"
        .Flags = 0
        .lpstrDefExt = "xlsx"
    End With
    
    lReturn = GetSaveFileName(ofn)
    
    If lReturn <> 0 Then
        ExibirDialogoSalvar = Trim$(ofn.lpstrFile)
    Else
        ExibirDialogoSalvar = ""
    End If
End Function
