VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientesDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variáveis do ADO
Private vol_Conexao As New clsConexao
Private vol_Command As New ADODB.Command


Public Property Set Conexao(ByVal vData As ADODB.Connection)
    Set cConexao = vData
End Property


'Metodos
Public Function CarregarGridClienteRS(ByRef pdtgClientes As DataGrid, ByRef padoClientes As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
On Error GoTo TrataErros
     
   CarregarGridClienteRS = True
   vol_Conexao.AbrirConexao
   
      padoClientes.ConnectionString = vol_Conexao.cConexao
      padoClientes.RecordSource = "SELECT ID_Cliente, Nome_Cliente, Numero_Cartao FROM Clientes (NOLOCK) ORDER BY " + IIf(pcmbLocalizar <= 0, "ID_Cliente", "Nome_Cliente") + " "
      padoClientes.Refresh
      
     ' Conectar o DataGrid ao ADODC
      Set pdtgClientes.DataSource = padoClientes
      
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CarregarGridClienteRS = False
        MsgBox "Não foi possível carregar informações do Cliente !", vbExclamation
    End If
    
    
End Function

Public Function PesquisarCliente(ByRef plvwClientes As ListView, ByVal pCliente As String, ByVal pBusca As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   PesquisarCliente = True
   vol_Conexao.AbrirConexao
      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT ID_Cliente, Nome_Cliente, Numero_Cartao FROM Clientes (NOLOCK) WHERE " + IIf(pBusca = 0, " ID_Cliente = ", IIf(pBusca = 1, " Nome_Cliente = ", "Numero_Cartao = ")) + "'" + pCliente + "'")
      With vol_Recordset
         If .RecordCount <> 0 Then
            Do Until .EOF
               plvwClientes.ListItems.Add , , .Fields(0).Value
               plvwClientes.ListItems(plvwClientes.ListItems.Count).SubItems(1) = .Fields(1).Value
               plvwClientes.ListItems(plvwClientes.ListItems.Count).SubItems(2) = .Fields(2).Value
               .MoveNext
            Loop
         End If
      End With
      Set vol_Recordset = Nothing
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        PesquisarCliente = False
        MsgBox "Não foi possível encontrar informações da Cliente !", vbExclamation
    End If
End Function

Public Function LocalizarCliente(ByRef pdtgClientes As DataGrid, ByRef padoClientes As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim I As Integer
On Error GoTo TrataErros
     
   LocalizarCliente = True
   vol_Conexao.AbrirConexao
   
      padoClientes.ConnectionString = vol_Conexao.cConexao
      If Trim$(ptxtLocalizar) = Empty Then
         padoClientes.RecordSource = "SELECT ID_Cliente, Nome_Cliente, Numero_Cartao FROM Clientes (NOLOCK) ORDER BY " + IIf(pcmbLocalizar = 0, "ID_Cliente", "Nome_Cliente") + " "
      Else
         padoClientes.RecordSource = "SELECT ID_Cliente, Nome_Cliente, Numero_Cartao FROM Clientes (NOLOCK) WHERE " + IIf(pcmbLocalizar = 0, "ID_Cliente", "Nome_Cliente") + " LIKE '%" + Trim$(ptxtLocalizar) + "%' ORDER BY " + IIf(pcmbLocalizar = 0, "ID_Cliente", "Nome_Cliente") + " "
      End If
      padoClientes.Refresh
      
     ' Conectar o DataGrid ao ADODC
      Set pdtgClientes.DataSource = padoClientes
      
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarCliente = False
        MsgBox "Não foi possível encontrar informações da Cliente !", vbExclamation
    End If
End Function

Public Function IncluirCliente(ByRef plvwCliente As ListView) As Boolean
Dim Incluir As Boolean
Dim Item As ListItem

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
         
         IncluirCliente = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
   
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "sp_InserirCliente"
         vol_Command.Parameters.Refresh
      
         ' Pega o primeiro (e único) item do ListView
         Set Item = plvwCliente.ListItems(1)
      
         With vol_Command.Parameters
           .Item(1).Value = Item.SubItems(1)  'Descricao
           .Item(2).Value = Item.SubItems(2)  'Capacidade
         End With
   
         vol_Command.Execute
         
      vol_Conexao.cConexao.CommitTrans
   
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       IncluirCliente = False
       MsgBox "Não foi possível incluir informações da Cliente !", vbExclamation
    End If
End Function

Public Function AlterarCliente(ByRef plvwCliente As ListView) As Boolean
Dim Incluir As Boolean
Dim Item As ListItem

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         AlterarCliente = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "sp_AtualizarCliente"
         vol_Command.Parameters.Refresh
      
         ' Pega o primeiro (e único) item do ListView
         Set Item = plvwCliente.ListItems(1)

         With vol_Command.Parameters
           .Item(1).Value = Item.text        'IdCliente
           .Item(2).Value = Item.SubItems(1) 'Descricao
           .Item(3).Value = Item.SubItems(2) 'Capacidade
         End With
      
         vol_Command.Execute
      
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       AlterarCliente = False
       MsgBox "Não foi possível incluir informações da Cliente !", vbExclamation
    End If
End Function

Public Function ExcluirCliente(ByVal pIdCliente As Integer) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         ExcluirCliente = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "sp_ExcluirCliente"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdCliente
         End With
      
         vol_Command.Execute
   
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       ExcluirCliente = False
       MsgBox "Não foi possível excluir informações da Cliente !", vbExclamation
    End If
End Function























