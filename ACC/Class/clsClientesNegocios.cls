VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientesNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Inicializa variáveis da classe
Private vop_ClientesDados As New clsClientesDados      'Variável de acesso classe clsClienteDados

'Variáveis da classe
Private vsp_Query As String                'Variável usada nas consultas sql
Private vip_Count As Integer               'Variável utilizadas nos loops
Private vbp_RollBack As Boolean            'Define que uma conexão está aberta e necessita de um rollback

'Variáveis das Propriedades
Private mvarIdCliente As Long              'Identificador da Clientes
Private mvarNome As String                 'Descrição da Clientes
Private mvarNumeroCartao As String         'NumeroCartao da Clientes




'Propriedades
'IdCliente
Public Property Let IdCliente(ByVal vData As Long)
    mvarIdCliente = vData
End Property

Public Property Get IdCliente() As Long
    IdCliente = mvarIdCliente
End Property

'Nome
Public Property Let Nome(ByVal vData As String)
    mvarNome = vData
End Property

Public Property Get Nome() As String
    Nome = mvarNome
End Property

'NumeroCartao
Public Property Let NumeroCartao(ByVal vData As String)
    mvarNumeroCartao = vData
End Property

Public Property Get NumeroCartao() As String
    NumeroCartao = mvarNumeroCartao
End Property



'Metodos
Public Function PesquisarCliente(ByRef plvwClientes As ListView, ByVal pCliente As String, ByVal pBusca As Integer) As Boolean

    PesquisarCliente = False
    Screen.MousePointer = vbHourglass
      If vop_ClientesDados.PesquisarCliente(plvwClientes, pCliente, pBusca) = True Then
         With plvwClientes
            IdCliente = Trim$(.ListItems(plvwClientes.ListItems.Count).text)
            Nome = Trim$(.ListItems(plvwClientes.ListItems.Count).SubItems(1))
            NumeroCartao = Trim$(.ListItems(plvwClientes.ListItems.Count).SubItems(2))
         End With
         PesquisarCliente = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function LocalizarCliente(ByRef pdtgClientes As DataGrid, ByRef padoClientes As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean

    LocalizarCliente = False
    Screen.MousePointer = vbHourglass
      If vop_ClientesDados.LocalizarCliente(pdtgClientes, padoClientes, ptxtLocalizar, pcmbLocalizar) = True Then
         LocalizarCliente = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function CarregarGridClienteRS(ByRef pdtgClientes As DataGrid, ByRef padoClientes As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridClienteRS = vop_ClientesDados.CarregarGridClienteRS(pdtgClientes, padoClientes, pcmbLocalizar)
    Screen.MousePointer = vbDefault
    
End Function

Public Function IncluirCliente(ByRef plvwCliente As ListView) As Boolean
Dim ListItem As ListItem

On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      ' Para adicionar dados na lista Cliente
      Set ListItem = plvwCliente.ListItems.Add(, , IdCliente)
      ListItem.SubItems(1) = Nome
      ListItem.SubItems(2) = NumeroCartao
      '
      IncluirCliente = vop_ClientesDados.IncluirCliente(plvwCliente)
      'Limpa lista
      plvwCliente.ListItems.Clear
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function AlterarCliente(ByRef plvwCliente As ListView) As Boolean
Dim ListItem As ListItem

On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      ' Para adicionar dados na lista Cliente
      Set ListItem = plvwCliente.ListItems.Add(, , IdCliente)
      ListItem.SubItems(1) = Nome
      ListItem.SubItems(2) = NumeroCartao
      '
      AlterarCliente = vop_ClientesDados.AlterarCliente(plvwCliente)
      'Limpa lista
      plvwCliente.ListItems.Clear
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function ExcluirCliente() As Boolean
   Screen.MousePointer = vbHourglass
      ExcluirCliente = vop_ClientesDados.ExcluirCliente(IdCliente)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
End Function















