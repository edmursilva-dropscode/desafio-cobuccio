VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransacaoDados"
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
Public Function CarregarGridTransacaoMainRS(ByRef pdtgTransacao As DataGrid, ByRef padoTransacao As Adodc, ByVal ptxtLocalizarTransacao As String, ByVal pcmbLocalizarTransacao As Integer, ByVal pDataInicial As String, ByVal pDataFinal As String) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim sqlConsulta As String

On Error GoTo TrataErros
     
   CarregarGridTransacaoMainRS = True
   vol_Conexao.AbrirConexao

      'Consulta sql
      sqlConsulta = "SELECT * FROM vw_Transacoes (NOLOCK) WHERE ID_Transacao > 0 "
      
      ' Adicionar condições de data à consulta
      If Not IsNull(pDataInicial) And Trim$(pDataInicial) <> Empty Then
          sqlConsulta = sqlConsulta & "AND Data_Transacao >= '" & Format(pDataInicial, "yyyy-MM-dd") & "' "
      End If
      If Not IsNull(pDataFinal) And Trim$(pDataFinal) <> Empty Then
          sqlConsulta = sqlConsulta & "AND Data_Transacao <= '" & Format(pDataFinal, "yyyy-MM-dd") & "' "
      End If
      sqlConsulta = sqlConsulta & "ORDER BY " + IIf(pcmbLocalizarTransacao = 0, "Nome_Cliente", IIf(pcmbLocalizarTransacao = 1, "Numero_Cartao", IIf(pcmbLocalizarTransacao = 2, "Categoria", "Data_Transacao"))) + " "
            
      padoTransacao.ConnectionString = vol_Conexao.cConexao
      padoTransacao.RecordSource = sqlConsulta
      
      padoTransacao.Refresh
      
      ' Conectar o DataGrid ao ADODC
      Set pdtgTransacao.DataSource = padoTransacao
      
   vol_Conexao.FecharConexao
       
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CarregarGridTransacaoMainRS = False
        MsgBox "Não foi possível carregar informações das Transações !", vbExclamation
    End If
End Function

Public Function PesquisarTransacao(ByRef plvwTransacao As ListView, ByVal pTransacao As String, ByVal pBusca As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
On Error GoTo TrataErros
     
   PesquisarTransacao = True
   vol_Conexao.AbrirConexao
      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT ID_Transacao, Numero_Cartao, Data_transacao, Valor_Transacao, Descricao FROM Transacoes (NOLOCK) WHERE " + IIf(pBusca = 0, " ID_Transacao = ", " Numero_Cartao = ") + "'" + pTransacao + "'")
      With vol_Recordset
         If .RecordCount <> 0 Then
            Do Until .EOF
               plvwTransacao.ListItems.Add , , .Fields(0).Value
               plvwTransacao.ListItems(plvwTransacao.ListItems.Count).SubItems(1) = .Fields(1).Value
               plvwTransacao.ListItems(plvwTransacao.ListItems.Count).SubItems(2) = .Fields(2).Value
               plvwTransacao.ListItems(plvwTransacao.ListItems.Count).SubItems(3) = .Fields(3).Value
               plvwTransacao.ListItems(plvwTransacao.ListItems.Count).SubItems(4) = .Fields(4).Value
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
        PesquisarTransacao = False
        MsgBox "Não foi possível encontrar informações da Transacao !", vbExclamation
    End If
End Function

Public Function LocalizarTransacao(ByRef pdtgTransacao As DataGrid, ByRef padoTransacao As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer, ByVal pDataInicial As String, ByVal pDataFinal As String) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim sqlConsulta As String

On Error GoTo TrataErros
     
   LocalizarTransacao = True
   vol_Conexao.AbrirConexao
   
      'Consulta sql
      If Trim$(ptxtLocalizar) = Empty Then
         sqlConsulta = "SELECT * FROM vw_Transacoes (NOLOCK) "
      Else
         sqlConsulta = "SELECT * FROM vw_Transacoes (NOLOCK) " _
                     & "WHERE " + IIf(pcmbLocalizar = 0, "Nome_Cliente", IIf(pcmbLocalizar = 1, "Numero_Cartao", IIf(pcmbLocalizar = 2, "Categoria", "CONVERT(VARCHAR, Data_Transacao, 103)"))) + " LIKE '%" + Trim$(ptxtLocalizar) + "%' "
      End If
   
      ' Adicionar condições de data à consulta
      If Not IsNull(pDataInicial) And Trim$(pDataInicial) <> Empty Then
          sqlConsulta = sqlConsulta & "AND Data_Transacao >= '" & Format(pDataInicial, "yyyy-MM-dd") & "' "
      End If
      If Not IsNull(pDataFinal) And Trim$(pDataFinal) <> Empty Then
          sqlConsulta = sqlConsulta & "AND Data_Transacao <= '" & Format(pDataFinal, "yyyy-MM-dd") & "' "
      End If
      sqlConsulta = sqlConsulta & "ORDER BY " + IIf(pcmbLocalizar = 0, "Nome_Cliente", IIf(pcmbLocalizar = 1, "Numero_Cartao", IIf(pcmbLocalizar = 2, "Categoria", "Data_Transacao"))) + " "
   
      '
      padoTransacao.ConnectionString = vol_Conexao.cConexao
      padoTransacao.RecordSource = sqlConsulta
      
      padoTransacao.Refresh
      
      ' Conectar o DataGrid ao ADODC
      Set pdtgTransacao.DataSource = padoTransacao
      
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarTransacao = False
        MsgBox "Não foi possível encontrar informações da Transacao !", vbExclamation
    End If
End Function

Public Function IncluirTransacao(ByRef plvwTransacao As ListView) As Boolean
Dim Incluir As Boolean
Dim Item As ListItem

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
         
         IncluirTransacao = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
   
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "sp_InserirTransacao"
         vol_Command.Parameters.Refresh
      
         ' Pega o primeiro (e único) item do ListView
         Set Item = plvwTransacao.ListItems(1)
      
         With vol_Command.Parameters
           .Item(1).Value = Item.SubItems(1)  'Numero_cartao
           .Item(2).Value = Item.SubItems(2)  'Data_Transacao
           .Item(3).Value = Item.SubItems(3)  'Valor_Transacao
           .Item(4).Value = Item.SubItems(4)  'Descricao
         End With
   
         vol_Command.Execute
         
      vol_Conexao.cConexao.CommitTrans
   
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       IncluirTransacao = False
       MsgBox "Não foi possível incluir informações da Transacao !", vbExclamation
    End If
End Function

Public Function AlterarTransacao(ByRef plvwTransacao As ListView) As Boolean
Dim Incluir As Boolean
Dim Item As ListItem

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         AlterarTransacao = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "sp_AtualizarTransacao"
         vol_Command.Parameters.Refresh
      
         ' Pega o primeiro (e único) item do ListView
         Set Item = plvwTransacao.ListItems(1)

         With vol_Command.Parameters
           .Item(1).Value = Item.text         'IdTransacao
           .Item(2).Value = Item.SubItems(1)  'Numero_cartao
           .Item(3).Value = Item.SubItems(2)  'Data_Transacao
           .Item(4).Value = Item.SubItems(3)  'Valor_Transacao
           .Item(5).Value = Item.SubItems(4)  'Descricao
         End With
      
      
         vol_Command.Execute
      
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       AlterarTransacao = False
       MsgBox "Não foi possível incluir informações da Transacao !", vbExclamation
    End If
End Function

Public Function ExcluirTransacao(ByVal pIdTransacao As Integer) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         ExcluirTransacao = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "sp_ExcluirTransacao"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdTransacao
         End With
      
         vol_Command.Execute
   
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       ExcluirTransacao = False
       MsgBox "Não foi possível excluir informações da Transacao !", vbExclamation
    End If
End Function











