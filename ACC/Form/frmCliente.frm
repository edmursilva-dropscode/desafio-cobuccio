VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{DC81D4AD-48D8-4DD6-A8B5-228CB11C1826}#1.0#0"; "prjXTab.ocx"
Begin VB.Form frmCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6630
   LinkTopic       =   "frmCliente"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   360
      Left            =   4410
      TabIndex        =   7
      Top             =   2250
      Width           =   1005
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   360
      Left            =   5490
      TabIndex        =   8
      Top             =   2235
      Width           =   1005
   End
   Begin prjXTab.XTab xtbCliente 
      Height          =   1320
      Left            =   75
      TabIndex        =   0
      Top             =   780
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   2328
      TabCount        =   1
      TabCaption(0)   =   "  Cliente "
      TabContCtrlCnt(0)=   2
      Tab(0)ContCtrlCap(1)=   "fraTab1"
      Tab(0)ContCtrlCap(2)=   "lblCodigo"
      TabStyle        =   1
      TabTheme        =   1
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin VB.Frame fraTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   1
         Left            =   165
         TabIndex        =   11
         Top             =   450
         Width           =   6180
         Begin VB.TextBox txtNome 
            Height          =   315
            Left            =   1185
            MaxLength       =   40
            TabIndex        =   3
            Tag             =   "0"
            Top             =   0
            Width           =   4935
         End
         Begin VB.TextBox txtNumeroCartao 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1185
            MaxLength       =   16
            TabIndex        =   5
            Tag             =   "0"
            Text            =   "0"
            Top             =   405
            Width           =   2055
         End
         Begin VB.Label lblNome 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   2
            Top             =   75
            Width           =   570
         End
         Begin VB.Label lblNumeroCartao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. cartão:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   45
            TabIndex        =   4
            Top             =   465
            Width           =   960
         End
      End
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   -74940
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   5080
      End
      Begin VB.Label lblCodigo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1020
         TabIndex        =   1
         Top             =   15
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lvwCliente 
      Height          =   360
      Left            =   2520
      TabIndex        =   6
      Top             =   2220
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   635
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IdCliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NumeroCartao"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   12
      Left            =   60
      Picture         =   "frmCliente.frx":0000
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   17955
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -1605
      Picture         =   "frmCliente.frx":08CA
      Top             =   675
      Width           =   10740
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_ClientesNegocios As New clsClientesNegocios

'Variaveis de controle do form
Private vbp_Cliente As Boolean                             'Verifica uma inclusao ou alteracao


'Eventos
Private Sub Form_Activate()
   
   Me.Refresh
   
End Sub

Public Sub Form_Load()
    
    vbp_Cliente = False              'Verifica uma inclusao ou alteracao
    
    'Inicializa entrada e saida
    Call InicializaCliente
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo TrataErros

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
   
   Set frmCliente = Nothing
   
End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtNumeroCartao_Change()
    
    If VerNumeros = False Then Exit Sub
    
End Sub

Private Sub txtNumeroCartao_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtNumeroCartao_KeyPress(KeyAscii As Integer)
    
    'Permite Backspace e Enter
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
    'Permite apenas números e formato de Moedas
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
            
End Sub

Private Sub cmdGravar_Click()
Dim vsp_Mensagem As String

   'Valida mensagem
   If vbp_Cliente = False Then
      vsp_Mensagem = "Confirma a Inclusão ?"
   Else
      vsp_Mensagem = "Confirma a Alteração ?"
   End If

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   
   If MsgBox(vsp_Mensagem, vbQuestion + vbYesNo, "Confirme !") = vbYes Then
   
      Set vop_ClientesNegocios = New clsClientesNegocios
          vop_ClientesNegocios.IdCliente = lblCodigo.Caption
          vop_ClientesNegocios.Nome = txtNome.text
          vop_ClientesNegocios.NumeroCartao = txtNumeroCartao.text
          If vbp_Cliente = False Then
             If vop_ClientesNegocios.IncluirCliente(lvwCliente) = True Then
                MsgBox "Cliente cadastrado com sucesso !", vbExclamation, "Cliente"
             End If
          Else
             If vop_ClientesNegocios.AlterarCliente(lvwCliente) = True Then
                MsgBox "Cliente alterado com sucesso !", vbExclamation, "Cliente"
             End If
          End If
      Set vop_ClientesNegocios = Nothing
      
   End If
   
   'Atualiza grid
   Call frmListaClientes.CarregarGrid
   
   'Inicializa entrada e saida
   Call InicializaCliente
   
   'Valida entrada de dados
   If vbp_Cliente = True Then
      Call cmdFechar_Click
   End If
   
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub



'Funcoes
Function Editar(ByVal pIdCliente As Integer) As Boolean
    
    'Verifica uma inclusao ou alteracao do cliente
    vbp_Cliente = True
    'Controle de exibicao
    lblCodigo.Visible = True
    
    Set vop_ClientesNegocios = New clsClientesNegocios
    
        If vop_ClientesNegocios.PesquisarCliente(lvwCliente, pIdCliente, 0) = True Then
           lblCodigo.Caption = pIdCliente
           txtNome.text = vop_ClientesNegocios.Nome
           txtNumeroCartao.text = vop_ClientesNegocios.NumeroCartao
        Else
            MsgBox "Não foi possível encontrar o Cliente !", vbCritical, "Cliente"
        End If
          
    Set vop_ClientesNegocios = Nothing
    
    Me.Show vbModal

End Function

Function VerCampos() As Boolean
    
    If Trim$(txtNome.text) = Empty Then
        MsgBox "Informe o nome do Cliente !", vbExclamation, "Cliente"
        'txtNome.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtNumeroCartao.text) = Empty Or Trim$(txtNumeroCartao.text) = "0" Then
        MsgBox "Informe o numero do cartão do Cliente !", vbExclamation, "Cliente"
        'txtNumeroCartao.SetFocus
        VerCampos = False
        Exit Function
    End If
    
    VerCampos = True

End Function

Private Function VerNumeros() As Boolean

    If IsNumeric(txtNumeroCartao.text) = False Then
       'txtTempoDePreparo.SetFocus
       VerNumeros = False
       Exit Function
    End If

    VerNumeros = True

End Function

Private Function DefaultCampos() As BookmarkEnum

    txtNumeroCartao.text = "0"

End Function

Private Function InicializaCliente() As Boolean

    'Limpa entrada de dados
    Call LimpaCampos(Me)
    
    'Inicializa entrada de dados
    Call DefaultCampos

End Function
