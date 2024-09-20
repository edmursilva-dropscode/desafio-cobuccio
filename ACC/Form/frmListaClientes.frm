VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmListaClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Clientes"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7305
   LinkTopic       =   "frmListaClentes"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cmbLocalizar 
      Height          =   315
      ItemData        =   "frmListaClientes.frx":0000
      Left            =   4110
      List            =   "frmListaClientes.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   825
      Width           =   1410
   End
   Begin VB.TextBox txtLocalizar 
      Height          =   315
      Left            =   870
      MaxLength       =   40
      TabIndex        =   2
      Top             =   825
      Width           =   2475
   End
   Begin MSAdodcLib.Adodc adoClientes 
      Height          =   330
      Left            =   1245
      Top             =   4710
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
      RecordSource    =   "SELECT ID_Cliente, Nome_Cliente, Numero_Cartao FROM Clientes ORDER BY ID_Cliente"
      Caption         =   "adoClientes"
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
   Begin MSDataGridLib.DataGrid dtgClientes 
      Bindings        =   "frmListaClientes.frx":001C
      Height          =   3345
      Left            =   210
      TabIndex        =   5
      Top             =   1245
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   5900
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "ID_Cliente"
         Caption         =   "   Código"
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
         Caption         =   "Nome"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   3344,882
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2115,213
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   360
      Left            =   6180
      TabIndex        =   8
      Top             =   4680
      Width           =   1005
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   360
      Left            =   5145
      TabIndex        =   7
      Top             =   4695
      Width           =   1005
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   360
      Left            =   4110
      TabIndex        =   6
      Top             =   4695
      Width           =   1005
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      Height          =   195
      Index           =   2
      Left            =   3660
      TabIndex        =   3
      Top             =   885
      Width           =   360
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localizar:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   885
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   12
      Left            =   105
      Picture         =   "frmListaClientes.frx":0036
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   11
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17955
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Index           =   10
      Left            =   -1605
      Picture         =   "frmListaClientes.frx":0900
      Top             =   675
      Width           =   10740
   End
End
Attribute VB_Name = "frmListaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_ClientesNegocios As New clsClientesNegocios

'Variaveis de controle do form
Private vil_IdCliente As Long                               'Identificador da Cliente
Private bMovingProgrammatically As Boolean                  ' Variável de controle para evitar movimentação dupla
Private vil_LastBookmark As Variant                         ' Variável de módulo para armazenar o último Bookmark



'Eventos
Private Sub Form_Activate()
   
    Me.Refresh
    
End Sub

Private Sub Form_Load()
    
    'Carrega clientes
    Call CarregarGrid
    cmbLocalizar.ListIndex = 0
    
End Sub

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

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frmListaClientes = Nothing
   
End Sub

Private Sub cmdExcluir_Click()
    If vil_IdCliente = 0 Then Exit Sub

    If MsgBox("Confirma a Exclusão ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      Set vop_ClientesNegocios = New clsClientesNegocios
          vop_ClientesNegocios.IdCliente = vil_IdCliente
          If vop_ClientesNegocios.ExcluirCliente() = True Then
             MsgBox "Cliente excluído com sucesso !", vbExclamation, "Cliente"
             'Limpa variável
             txtLocalizar.text = Empty
             Call CarregarGrid
          End If
      Set vop_ClientesNegocios = Nothing
    End If
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdNovo_Click()

On Error GoTo TrataErros
   
    'Inicia form
    frmCliente.Show vbModal
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_ClientesNegocios = Nothing
       Exit Sub
    End If
   
End Sub


Private Sub cmbLocalizar_Click()

On Error GoTo TrataErros

   Set vop_ClientesNegocios = New clsClientesNegocios
       Call vop_ClientesNegocios.LocalizarCliente(dtgClientes, adoClientes, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_ClientesNegocios = Nothing

   txtLocalizar.text = Empty

TrataErros:
    If Err.Number <> 0 Then
       Set vop_ClientesNegocios = Nothing
       Exit Sub
    End If

End Sub

Private Sub dtgClientes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' Só atualiza se não estiver movendo programaticamente
    If Not bMovingProgrammatically Then
        If dtgClientes.Bookmark > 0 Then
            If dtgClientes.SelBookmarks.Count > 0 Then
                dtgClientes.SelBookmarks.Remove 0
            End If
            dtgClientes.SelBookmarks.Add dtgClientes.Bookmark
            If IsNull(dtgClientes.Columns(0).text) Or dtgClientes.Columns(0).text = "" Then
                vil_IdCliente = 0
            Else
                vil_IdCliente = CInt(dtgClientes.Columns(0).text)
            End If
        End If
    End If
End Sub

Private Sub dtgClientes_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            MoveRecord adoClientes.Recordset, False
            'Valida fim do DataGrid
        Case vbKeyUp
            MoveRecord adoClientes.Recordset, True
    End Select
End Sub

Private Sub MoveRecord(rs As ADODB.Recordset, MovePrevious As Boolean)
    If rs.RecordCount > 0 Then
        bMovingProgrammatically = True
        
        If dtgClientes.SelBookmarks.Count > 0 Then
            dtgClientes.SelBookmarks.Remove 0
        End If
        
        If MovePrevious Then
            rs.MovePrevious
            If rs.BOF Then
                rs.MoveFirst
                dtgClientes.SelBookmarks.Add dtgClientes.Bookmark
            Else
                rs.MoveNext
            End If
        Else
            rs.MoveNext
            If rs.EOF Then
                rs.MoveLast
                dtgClientes.SelBookmarks.Add dtgClientes.Bookmark
            Else
               rs.MovePrevious
            End If
        End If
        
        dtgClientes.Bookmark = rs.Bookmark
        
        bMovingProgrammatically = False ' Reseta o flag
    End If
End Sub

Private Sub dtgClientes_DblClick()
Dim vvl_BookMark As Variant
Dim vil_RowIndex As Long

On Error GoTo TrataErros

    vil_RowIndex = dtgClientes.Row
    vvl_BookMark = dtgClientes.RowBookmark(vil_RowIndex)
    If vvl_BookMark = Empty Then Exit Sub

    Call frmCliente.Form_Load
    Call frmCliente.Editar(vil_IdCliente)

    dtgClientes.Bookmark = vvl_BookMark
    dtgClientes.Scroll 0, dtgClientes.RowContaining(vvl_BookMark)
    dtgClientes.SelBookmarks.Add vvl_BookMark
    dtgClientes.Refresh

TrataErros:
    If Err.Number <> 0 Then Exit Sub

End Sub

Private Sub txtLocalizar_Change()

On Error GoTo TrataErros

   Set vop_ClientesNegocios = New clsClientesNegocios
       Call vop_ClientesNegocios.LocalizarCliente(dtgClientes, adoClientes, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_ClientesNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_ClientesNegocios = Nothing
       Exit Sub
    End If

End Sub

'Medotos
Public Sub CarregarGrid()
Dim vbl_Carregar As Boolean '

On Error GoTo TrataErros '

    Set vop_ClientesNegocios = New clsClientesNegocios '

        vbl_Carregar = vop_ClientesNegocios.CarregarGridClienteRS(dtgClientes, adoClientes, cmbLocalizar.ListIndex)
        If vbl_Carregar = True And adoClientes.Recordset.RecordCount > 0 Then
           cmdNovo.Left = 4050
           cmdExcluir.Left = 5115
           cmdExcluir.Visible = True
        Else
           dtgClientes.Refresh
           cmdNovo.Left = 5115
           cmdExcluir.Visible = False
        End If

    Set vop_ClientesNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_ClientesNegocios = Nothing
       Exit Sub
    End If

End Sub
