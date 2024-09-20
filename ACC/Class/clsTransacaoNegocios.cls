VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransacaoNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Inicializa variáveis da classe
Private vop_TransacaoDados As New clsTransacaoDados      'Variável de acesso classe clsTransicaoDados

'Variáveis da classe
Private vsp_Query As String                'Variável usada nas consultas sql
Private vip_Count As Integer               'Variável utilizadas nos loops
Private vbp_RollBack As Boolean            'Define que uma conexão está aberta e necessita de um rollback

'Variáveis das Propriedades
Private mvarIdTransacao As Long               'Identificador da Transacao
Private mvarNumero_Cartao As String           'Identificador do Numero_Cartao
Private mvarValor_Transacao As Double         'Valor_Transacao
Private mvarData_Transacao As Date            'Data_Transacao
Private mvarDescricao As String               'Descricao


'Propriedades
'IdTransacao
Public Property Let IdTransacao(ByVal vData As Long)
    mvarIdTransacao = vData
End Property

Public Property Get IdTransacao() As Long
    IdTransacao = mvarIdTransacao
End Property

'Numero_Cartao
Public Property Let Numero_Cartao(ByVal vData As String)
    mvarNumero_Cartao = vData
End Property

Public Property Get Numero_Cartao() As String
    Numero_Cartao = mvarNumero_Cartao
End Property

'Valor_Transacao
Public Property Let Valor_Transacao(ByVal vData As Double)
    mvarValor_Transacao = vData
End Property

Public Property Get Valor_Transacao() As Double
    Valor_Transacao = mvarValor_Transacao
End Property

'Data_Transacao
Public Property Let Data_Transacao(ByVal vData As Date)
    mvarData_Transacao = vData
End Property

Public Property Get Data_Transacao() As Date
    Data_Transacao = mvarData_Transacao
End Property

'Descricao
Public Property Let Descricao(ByVal vData As String)
    mvarDescricao = vData
End Property

Public Property Get Descricao() As String
    Descricao = mvarDescricao
End Property



'Metodos
Public Function PesquisarTransacao(ByRef plvwTransacao As ListView, ByVal pTransacao As String, ByVal pBusca As Integer) As Boolean

    PesquisarTransacao = False
    Screen.MousePointer = vbHourglass
      If vop_TransacaoDados.PesquisarTransacao(plvwTransacao, pTransacao, pBusca) = True Then
         With plvwTransacao
            IdTransacao = Trim$(.ListItems(plvwTransacao.ListItems.Count).text)
            Numero_Cartao = Trim$(.ListItems(plvwTransacao.ListItems.Count).SubItems(1))
            Data_Transacao = Trim$(.ListItems(plvwTransacao.ListItems.Count).SubItems(2))
            Valor_Transacao = Trim$(.ListItems(plvwTransacao.ListItems.Count).SubItems(3))
            Descricao = Trim$(.ListItems(plvwTransacao.ListItems.Count).SubItems(4))
         End With
         PesquisarTransacao = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function LocalizarTransacao(ByRef pdtgTransacao As DataGrid, ByRef padoTransacao As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer, ByVal pDataInicial As String, ByVal pDataFinal As String) As Boolean

    LocalizarTransacao = False
    Screen.MousePointer = vbHourglass
      If vop_TransacaoDados.LocalizarTransacao(pdtgTransacao, padoTransacao, ptxtLocalizar, pcmbLocalizar, pDataInicial, pDataFinal) = True Then
         LocalizarTransacao = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function CarregarGridTransacaoMainRS(ByRef pdtgTransacao As DataGrid, ByRef padoTransacao As Adodc, ByVal ptxtLocalizarTransacao As String, ByVal pcmbLocalizarTransacao As Integer, ByVal pDataInicial As String, ByVal pDataFinal As String) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridTransacaoMainRS = vop_TransacaoDados.CarregarGridTransacaoMainRS(pdtgTransacao, padoTransacao, ptxtLocalizarTransacao, pcmbLocalizarTransacao, pDataInicial, pDataFinal)
    Screen.MousePointer = vbDefault
    
End Function

Public Function IncluirTransacao(ByVal plvwTransacao As ListView) As Boolean
Dim ListItem As ListItem

On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      ' Para adicionar dados na lista Transacao
      Set ListItem = plvwTransacao.ListItems.Add(, , IdTransacao)
      ListItem.SubItems(1) = Numero_Cartao
      ListItem.SubItems(2) = Data_Transacao
      ListItem.SubItems(3) = Valor_Transacao
      ListItem.SubItems(4) = Descricao
      '
      IncluirTransacao = vop_TransacaoDados.IncluirTransacao(plvwTransacao)
      'Limpa lista
      plvwTransacao.ListItems.Clear
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function AlterarTransacao(ByVal plvwTransacao As ListView) As Boolean
Dim ListItem As ListItem

On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      ' Para adicionar dados na lista Transacao
      Set ListItem = plvwTransacao.ListItems.Add(, , IdTransacao)
      ListItem.SubItems(1) = Numero_Cartao
      ListItem.SubItems(2) = Data_Transacao
      ListItem.SubItems(3) = Valor_Transacao
      ListItem.SubItems(4) = Descricao
      '
      AlterarTransacao = vop_TransacaoDados.AlterarTransacao(plvwTransacao)
      'Limpa lista
      plvwTransacao.ListItems.Clear
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
End Function

Public Function ExcluirTransacao() As Boolean
   Screen.MousePointer = vbHourglass
      ExcluirTransacao = vop_TransacaoDados.ExcluirTransacao(IdTransacao)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
End Function



















