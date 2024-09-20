VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstatisticaDados"
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
Public Function CarregarGridEstatisticaMainRS(ByRef pdtgEstatisticas As DataGrid, ByRef padoEstatisticas As Adodc, ByVal ptxtLocalizarEstatisticas As String, ByVal pcmbLocalizarEstatisticas As Integer, ByVal pDataInicial As String, ByVal pDataFinal As String) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim sqlConsulta As String

On Error GoTo TrataErros
     
   CarregarGridEstatisticaMainRS = True
   vol_Conexao.AbrirConexao

      'Consulta sql
      sqlConsulta = "SELECT Total.Numero_Cartao, Total.Nome_Cliente, Total.Valor_Total, Total.Quantidade_Transacoes, FORMAT(Total.Ultima_Data, 'dd-MM-yyyy') AS Ultima_Data_Formatada "
      sqlConsulta = sqlConsulta & "FROM "
      sqlConsulta = sqlConsulta & "    ( "
      sqlConsulta = sqlConsulta & "        SELECT c.Numero_Cartao, c.Nome_Cliente, SUM(t.Valor_Transacao) AS Valor_Total, COUNT(t.ID_Transacao) AS Quantidade_Transacoes, MAX(t.Data_Transacao) AS Ultima_Data "
      sqlConsulta = sqlConsulta & "        FROM Transacoes t "
      sqlConsulta = sqlConsulta & "        INNER JOIN Clientes c ON t.Numero_Cartao = c.Numero_Cartao "
      sqlConsulta = sqlConsulta & "        WHERE t.ID_Transacao > 0 "
      ' Adicionar condições de data à consulta
      If Not IsNull(pDataInicial) And Trim$(pDataInicial) <> Empty Then
          sqlConsulta = sqlConsulta & "AND t.Data_Transacao >= '" & Format(pDataInicial, "yyyy-MM-dd") & "' "
      End If
      If Not IsNull(pDataFinal) And Trim$(pDataFinal) <> Empty Then
          sqlConsulta = sqlConsulta & "AND t.Data_Transacao <= '" & Format(pDataFinal, "yyyy-MM-dd") & "' "
      End If
      sqlConsulta = sqlConsulta & "            GROUP BY c.Numero_Cartao, c.Nome_Cliente ) AS Total "
      sqlConsulta = sqlConsulta & "ORDER BY " + IIf(pcmbLocalizarEstatisticas = 0, "Total.Nome_Cliente", "Total.Ultima_Data") + " "
            
      padoEstatisticas.ConnectionString = vol_Conexao.cConexao
      padoEstatisticas.RecordSource = sqlConsulta
      
      padoEstatisticas.Refresh
      
      ' Conectar o DataGrid ao ADODC
      Set pdtgEstatisticas.DataSource = padoEstatisticas
      
   vol_Conexao.FecharConexao
       
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CarregarGridEstatisticaMainRS = False
        MsgBox "Não foi possível carregar informações das Estatísticas !", vbExclamation
    End If
End Function

Public Function LocalizarEstatistica(ByRef pdtgEstatisticas As DataGrid, ByRef padoEstatisticas As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer, ByVal pDataInicial As String, ByVal pDataFinal As String) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim sqlConsulta As String

On Error GoTo TrataErros
     
   LocalizarEstatistica = True
   vol_Conexao.AbrirConexao
      
      'Consulta sql
      sqlConsulta = "SELECT Total.Numero_Cartao, Total.Nome_Cliente, Total.Valor_Total, Total.Quantidade_Transacoes, FORMAT(Total.Ultima_Data, 'dd-MM-yyyy') AS Ultima_Data_Formatada "
      sqlConsulta = sqlConsulta & "FROM "
      sqlConsulta = sqlConsulta & "    ( "
      sqlConsulta = sqlConsulta & "        SELECT c.Numero_Cartao, c.Nome_Cliente, SUM(t.Valor_Transacao) AS Valor_Total, COUNT(t.ID_Transacao) AS Quantidade_Transacoes, MAX(t.Data_Transacao) AS Ultima_Data "
      sqlConsulta = sqlConsulta & "        FROM Transacoes t "
      sqlConsulta = sqlConsulta & "        INNER JOIN Clientes c ON t.Numero_Cartao = c.Numero_Cartao "
      sqlConsulta = sqlConsulta & "        WHERE t.ID_Transacao > 0 "
      ' Adicionar condições de data à consulta
      If Not IsNull(pDataInicial) And Trim$(pDataInicial) <> Empty Then
          sqlConsulta = sqlConsulta & "AND t.Data_Transacao >= '" & Format(pDataInicial, "yyyy-MM-dd") & "' "
      End If
      If Not IsNull(pDataFinal) And Trim$(pDataFinal) <> Empty Then
          sqlConsulta = sqlConsulta & "AND t.Data_Transacao <= '" & Format(pDataFinal, "yyyy-MM-dd") & "' "
      End If
      sqlConsulta = sqlConsulta & "AND " + IIf(pcmbLocalizar = 0, "c.Nome_Cliente", "CONVERT(VARCHAR, t.Data_Transacao, 103)") + " LIKE '%" + Trim$(ptxtLocalizar) + "%' "
      sqlConsulta = sqlConsulta & "            GROUP BY c.Numero_Cartao, c.Nome_Cliente ) AS Total "
      sqlConsulta = sqlConsulta & "ORDER BY " + IIf(pcmbLocalizar = 0, "Total.Nome_Cliente", "Total.Ultima_Data") + " "
      
      padoEstatisticas.ConnectionString = vol_Conexao.cConexao
      padoEstatisticas.RecordSource = sqlConsulta
      
      padoEstatisticas.Refresh
      
      ' Conectar o DataGrid ao ADODC
      Set pdtgEstatisticas.DataSource = padoEstatisticas
      
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarEstatistica = False
        MsgBox "Não foi possível encontrar informações da Estatísticas !", vbExclamation
    End If
End Function

