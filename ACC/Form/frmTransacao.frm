VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{DC81D4AD-48D8-4DD6-A8B5-228CB11C1826}#1.0#0"; "prjXTab.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmTransacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Transação"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9465
   LinkTopic       =   "frmTransacao"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   360
      Left            =   60
      TabIndex        =   14
      Top             =   3105
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   360
      Left            =   8370
      TabIndex        =   16
      Top             =   3090
      Width           =   1005
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   360
      Left            =   7320
      TabIndex        =   15
      Top             =   3105
      Width           =   1005
   End
   Begin MSACAL.Calendar calData 
      Height          =   315
      Left            =   3255
      TabIndex        =   8
      Top             =   1665
      Width           =   2280
      _Version        =   524288
      _ExtentX        =   4022
      _ExtentY        =   556
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2024
      Month           =   9
      Day             =   18
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   2
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.01
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjXTab.XTab xtbCliente 
      Height          =   2175
      Left            =   60
      TabIndex        =   0
      Top             =   795
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   3836
      TabCount        =   1
      TabCaption(0)   =   "  Transação  "
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
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   -74940
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   5080
      End
      Begin VB.Frame fraTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   1680
         Index           =   1
         Left            =   225
         TabIndex        =   18
         Top             =   450
         Width           =   8970
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1080
            MaxLength       =   9
            TabIndex        =   10
            Text            =   "0,00"
            Top             =   870
            Width           =   1500
         End
         Begin VB.CommandButton cmdPesquisar 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3225
            TabIndex        =   4
            Top             =   15
            Width           =   345
         End
         Begin VB.TextBox txtNumeroCartao 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1095
            MaxLength       =   16
            TabIndex        =   3
            Tag             =   "0"
            Text            =   "0"
            Top             =   45
            Width           =   2055
         End
         Begin VB.TextBox txtDescricao 
            Height          =   315
            Left            =   1065
            MaxLength       =   60
            TabIndex        =   12
            Tag             =   "0"
            Top             =   1275
            Width           =   7905
         End
         Begin VB.Label lblDataFinal 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1095
            TabIndex        =   7
            Top             =   450
            Width           =   1695
         End
         Begin VB.Label lblDescricao 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição:"
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
            Index           =   2
            Left            =   30
            TabIndex        =   11
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label lblValor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
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
            Index           =   1
            Left            =   45
            TabIndex        =   9
            Top             =   915
            Width           =   525
         End
         Begin VB.Label lblNomeCliente 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3675
            TabIndex        =   5
            Top             =   30
            Width           =   5295
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
            Left            =   30
            TabIndex        =   2
            Top             =   105
            Width           =   960
         End
         Begin VB.Label lblData 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data:"
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
            TabIndex        =   6
            Top             =   510
            Width           =   480
         End
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
         Left            =   1215
         TabIndex        =   1
         Top             =   15
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lvwTransacao 
      Height          =   360
      Left            =   3480
      TabIndex        =   13
      Top             =   3120
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IdTransacao"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "NumeroCartao"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Descricao"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwCliente 
      Height          =   360
      Left            =   2160
      TabIndex        =   20
      Top             =   3120
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
   Begin VB.Image imgIcone 
      Height          =   555
      Left            =   105
      Picture         =   "frmTransacao.frx":0000
      Top             =   75
      Width           =   555
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   17955
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -1830
      Picture         =   "frmTransacao.frx":0540
      Top             =   675
      Width           =   10740
   End
End
Attribute VB_Name = "frmTransacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Private vop_TransacaoNegocios As New clsTransacaoNegocios
Private vop_ClientesNegocios As New clsClientesNegocios
Private vop_System As New clsSystem

'Variaveis de controle do form
Private vbp_Transacao As Boolean                             'Verifica uma inclusao ou alteracao



'Eventos
Private Sub Form_Activate()
   
   Me.Refresh
   
End Sub

Public Sub Form_Load()
    
    vbp_Transacao = False              'Verifica uma inclusao ou alteracao
    
    'Inicializa entrada e saida
    Call InicializaTransacao
    
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
   
   Set frmTransacao = Nothing
   
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

Private Sub txtNumeroCartao_LostFocus()

    If VerCartao() = False Then Exit Sub
    
End Sub

Private Sub cmdPesquisar_Click()

On Error GoTo TrataErros

   Set vop_System = New clsSystem
       vop_System.Coluna_01 = "No dp cartão:"
       vop_System.Coluna_02 = "Cliente:"
       vop_System.FormataColuna_01 = "000"
       vop_System.TamanhoColuna_01 = "2200"
       If vop_System.FindLista("Clientes", "Numero_Cartao", "Nome_Cliente", frmFindLista.lvwLista) = True Then
          frmFindLista.Caption = "Cliente"
          frmFindLista.optSort01.Caption = "No. do cartão"
          frmFindLista.optSort02.Caption = "Cliente"
          frmFindLista.Show vbModal
       Else
          txtNumeroCartao.SetFocus
          Set vop_System = Nothing
          Exit Sub
       End If
    Set vop_System = Nothing
    
    If Trim(frmFindLista.lvwLista.SelectedItem.text) <> Empty Then
       txtNumeroCartao.text = Trim(frmFindLista.lvwLista.SelectedItem)
       lblNomeCliente.Caption = frmFindLista.lvwLista.ListItems(frmFindLista.lvwLista.SelectedItem.Index).SubItems(1)
       Unload frmFindLista

    Else
       Unload frmFindLista
       txtNumeroCartao.SetFocus
    End If

TrataErros:
    If Err.Number <> 0 Then
       MsgBox "Não foi possível listar as Áreas !", vbCritical
       Err.Clear
    End If

End Sub

Private Sub calData_Click()
    lblDataFinal.Caption = calData.Value
End Sub

Private Sub fraTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Converter as coordenadas do mouse de twips para pixels
    Dim mouseX As Single
    Dim mouseY As Single
    
    mouseX = X / Screen.TwipsPerPixelX
    mouseY = Y / Screen.TwipsPerPixelY
    
    If ((mouseX >= 199 And mouseX <= 355) And (mouseY >= 25 And mouseY <= 50)) Then
       calData.Height = 1890
    Else
       calData.Height = 315
       lblDataFinal.Caption = calData.Value
    End If
    
End Sub

Private Sub txtValor_GotFocus()
    
    txtValor.SelStart = 0
    txtValor.SelLength = Len(txtValor.text)
    
End Sub

Private Sub txtValor_Change()
   
   If VerNumeros = False Then Exit Sub
   
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)

    If VerNumeros = False Then Exit Sub
    
    'Permite Backspace e Enter
    If KeyAscii = vbKeyBack Then Exit Sub
    'Permite apenas números e formato de Moedas
    If KeyAscii <> 13 Then
       If InStr("0123456789.,", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtDescricao_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub

Private Sub cmdExcluir_Click()
    If Trim$(lblCodigo.Caption) = Empty Or Trim$(lblCodigo.Caption) = "0" Then Exit Sub

    If MsgBox("Confirma a Exclusão ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      Set vop_TransacaoNegocios = New clsTransacaoNegocios
          vop_TransacaoNegocios.IdTransacao = lblCodigo.Caption
          If vop_TransacaoNegocios.ExcluirTransacao() = True Then
             MsgBox "Transação excluída com sucesso !", vbExclamation, "Transacao"
             'Atualiza grid
             Call frmMain.CarregarGridTransacao
            
             'Inicializa entrada e saida
             Call InicializaTransacao
            
             'Valida entrada de dados
             If vbp_Transacao = True Then
                Call cmdFechar_Click
             End If
          End If
      Set vop_TransacaoNegocios = Nothing
    End If
End Sub

Private Sub cmdGravar_Click()
Dim vsp_Mensagem As String

   'Valida mensagem
   If vbp_Transacao = False Then
      vsp_Mensagem = "Confirma a Inclusão ?"
   Else
      vsp_Mensagem = "Confirma a Alteração ?"
   End If

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   If MsgBox(vsp_Mensagem, vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      
      Set vop_TransacaoNegocios = New clsTransacaoNegocios
          vop_TransacaoNegocios.IdTransacao = lblCodigo.Caption
          vop_TransacaoNegocios.Numero_Cartao = txtNumeroCartao.text
          vop_TransacaoNegocios.Valor_Transacao = txtValor.text
          vop_TransacaoNegocios.Data_Transacao = lblDataFinal.Caption
          vop_TransacaoNegocios.Descricao = txtDescricao.text
          If vbp_Transacao = False Then
             If vop_TransacaoNegocios.IncluirTransacao(lvwTransacao) = True Then
                MsgBox "Transacao cadastrada com sucesso !", vbExclamation, "Transacao"
             End If
          Else
             If vop_TransacaoNegocios.AlterarTransacao(lvwTransacao) = True Then
                MsgBox "Transacao alterada com sucesso !", vbExclamation, "Transacao"
             End If
          End If
      Set vop_TransacaoNegocios = Nothing
      
   End If
   
   'Atualiza grid
   Call frmMain.CarregarGridTransacao
   
   'Inicializa entrada e saida
   Call InicializaTransacao
   
   'Valida entrada de dados
   If vbp_Transacao = True Then
      Call cmdFechar_Click
   End If
   
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub




'Funcoes
Function Editar(ByVal pIdTransacao As Integer) As Boolean
    
    'Verifica uma inclusao ou alteracao da transacao
    vbp_Transacao = True
    'Controle de exibicao
    lblCodigo.Visible = True
    cmdExcluir.Visible = True
    
    Set vop_TransacaoNegocios = New clsTransacaoNegocios
    
        If vop_TransacaoNegocios.PesquisarTransacao(lvwTransacao, pIdTransacao, 0) = True Then
           lblCodigo.Caption = pIdTransacao
           txtNumeroCartao.text = vop_TransacaoNegocios.Numero_Cartao
           txtValor.text = Format(vop_TransacaoNegocios.Valor_Transacao, "#,##0.00")
           lblDataFinal.Caption = vop_TransacaoNegocios.Data_Transacao
           txtDescricao.text = vop_TransacaoNegocios.Descricao
           'Nome do cliente
           Set vop_ClientesNegocios = New clsClientesNegocios
              If vop_ClientesNegocios.PesquisarCliente(lvwCliente, txtNumeroCartao.text, 2) = True Then
                 lblNomeCliente.Caption = lvwCliente.ListItems(lvwCliente.ListItems.Count).SubItems(1)
              End If
           Set vop_ClientesNegocios = Nothing
           'Limpa lista
           lvwTransacao.ListItems.Clear
           lvwCliente.ListItems.Clear
        Else
            MsgBox "Não foi possível encontrar a Transação !", vbCritical, "Transação"
        End If
          
    Set vop_TransacaoNegocios = Nothing
    
    Me.Show vbModal

End Function

Private Function VerCartao() As Boolean

    If Trim$(txtNumeroCartao.text) = "0" Then
       MsgBox "Cartão existe !", vbCritical, "Transação"
       lblNomeCliente.Caption = Empty
       txtNumeroCartao.SetFocus
       VerCartao = False
       Exit Function
    ElseIf Trim$(txtNumeroCartao.text) <> Empty Then
       Set vop_ClientesNegocios = New clsClientesNegocios
          If vop_ClientesNegocios.PesquisarCliente(lvwCliente, txtNumeroCartao, 2) = False Then
             MsgBox "Cartão não existe !", vbCritical, "Transação"
             lblNomeCliente.Caption = Empty
             txtNumeroCartao.SetFocus
             VerCartao = False
             Exit Function
          End If
       Set vop_ClientesNegocios = Nothing
    Else
       lblNomeCliente.Caption = Empty
    End If

    VerCartao = True

End Function

Private Function VerNumeros() As Boolean

    If IsNumeric(txtNumeroCartao.text) = False Then
       If txtNumeroCartao.text <> Empty Then txtNumeroCartao.SetFocus
       VerNumeros = False
       Exit Function
    ElseIf IsNumeric(txtValor) = False Then
       If txtValor <> Empty Then txtValor.SetFocus
       VerNumeros = False
       Exit Function
    End If
       
    VerNumeros = True

End Function

Private Function DefaultCampos() As BookmarkEnum

    txtNumeroCartao.text = "0"
    txtValor.text = "0,00"
    lblDataFinal.Caption = calData.Value

End Function

Private Function InicializaTransacao() As Boolean

    'Limpa entrada de dados
    Call LimpaCampos(Me)
    
    'Inicializa entrada de dados
    Call DefaultCampos

End Function

Function VerCampos() As Boolean
    
    If Trim$(txtNumeroCartao.text) = Empty Or Trim$(txtNumeroCartao.text) = "0" Then
        MsgBox "Informe o numero do cartão do Cliente !", vbExclamation, "Transacao"
        'txtNumeroCartao.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtValor.text) = Empty Or Trim$(txtValor.text) = "0,00" Then
        MsgBox "Informe o valor da Transacao !", vbExclamation, "Transacao"
        'txtValor.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(txtDescricao.text) = Empty Then
        MsgBox "Informe a descrição da Transacao !", vbExclamation, "Transacao"
        'txtDescricao.SetFocus
        VerCampos = False
        Exit Function
    End If
        
    VerCampos = True

End Function
