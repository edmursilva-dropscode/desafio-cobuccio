VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Variáveis do ADO
Private vol_Conexao As New clsConexao

'Declaração para acessar o Registro
Private mvarColuna_01 As String
Private mvarColuna_02 As String
Private mvarColuna_03 As String
Private mvarTamanhoColuna_01 As Long
Private mvarFormataColuna_01 As String





Public Property Let TamanhoColuna_01(ByVal vData As Long)
    mvarTamanhoColuna_01 = vData
End Property

Public Property Get TamanhoColuna_01() As Long
    TamanhoColuna_01 = mvarTamanhoColuna_01
End Property


Public Property Let FormataColuna_01(ByVal vData As String)
    mvarFormataColuna_01 = vData
End Property

Public Property Get FormataColuna_01() As String
    FormataColuna_01 = mvarFormataColuna_01
End Property

Public Property Let Coluna_01(ByVal vData As String)
    mvarColuna_01 = vData
End Property

Public Property Get Coluna_01() As String
    Coluna_01 = mvarColuna_01
End Property

Public Property Let Coluna_02(ByVal vData As String)
    mvarColuna_02 = vData
End Property

Public Property Get Coluna_02() As String
    Coluna_02 = mvarColuna_02
End Property

Public Property Let Coluna_03(ByVal vData As String)
    mvarColuna_03 = vData
End Property

Public Property Get Coluna_03() As String
    Coluna_03 = mvarColuna_03
End Property







Function SysTime() As String

On Error GoTo TrataErros

    Set vrl_Record = New ADODB.Recordset
        vsl_Query = "SELECT TO_CHAR(SYSDATE,'HH24:MI:SS') AS HORA FROM DUAL"
        vrl_Record.Open vsl_Query, vop_Data.Conexao, adOpenStatic
        SysTime = vrl_Record("HORA")
        vrl_Record.Close
    Set vrl_Record = Nothing

TrataErros:
    If Err.Number = 5 Then
        Err.Clear
        Resume Next
    End If

End Function

Function SysDate() As Date

On Error GoTo TrataErros

    Set vrl_Record = New ADODB.Recordset
        vsl_Query = "SELECT TO_CHAR(SYSDATE,'DD/MM/YYYY') AS DATA FROM DUAL"
        vrl_Record.Open vsl_Query, vop_Data.Conexao, adOpenStatic
        SysDate = vrl_Record("DATA")
        vrl_Record.Close
    Set vrl_Record = Nothing

TrataErros:
    If Err.Number = 5 Then
        Err.Clear
        Resume Next
    End If

End Function




Function FindLista(ByVal Table$, ByVal Campo1$, ByVal Campo2$, ByVal pLista As ListView) As Boolean
Dim vol_ListItem As MSComctlLib.ListItem
Dim vol_RecordFindLista As New ADODB.Recordset          'Conexão do Banco de Dados

On Error GoTo TrataErros

    FindLista = False
    pLista.ListItems.Clear
    pLista.ColumnHeaders(1).text = Coluna_01
    pLista.ColumnHeaders(1).Width = TamanhoColuna_01
    pLista.ColumnHeaders(2).text = Coluna_02
    '
    vol_Conexao.AbrirConexao
       pLista.ListItems.Clear
       Set vol_RecordFindLista = vol_Conexao.RetornaRs("SELECT " & Campo1$ & ", " & Campo2$ & " FROM " & Table$ & " (NOLOCK) ORDER BY " & Campo1$)
       Do Until vol_RecordFindLista.EOF
          Set vol_ListItem = pLista.ListItems.Add(, , "     " + CStr(vol_RecordFindLista.Fields(0).Value))
              vol_ListItem.SubItems(1) = CStr(vol_RecordFindLista.Fields(1).Value)
          Set vol_ListItem = Nothing
          vol_RecordFindLista.MoveNext
       Loop
       Set vol_RecordFindLista = Nothing
    vol_Conexao.FecharConexao
    pLista.Sorted = True
    FindLista = True

TrataErros:
    If Err.Number = 5 Then
        Set vol_RecordFindLista = Nothing
        vol_Conexao.FecharConexao
        Err.Clear
        Resume Next
    End If

End Function

'Function CarregaCombo(ByVal Combo As Control, ByVal Table$, ByVal Campo$, Optional ListaDepartamento As ListView)
'Dim vcl_SqlCombo As String
'Dim vil_CountLista As Integer
'Dim vol_RecordCombo As New ADODB.Recordset          'Conexão do Banco de Dados
'
'On Error GoTo TrataErros
'
'    vol_Conexao.AbrirConexao
'       Combo.Clear
'       Set vol_RecordCombo = vol_Conexao.RetornaRs("SELECT " & Campo$ & " FROM " & Table$ & " (NOLOCK) ORDER BY " & Campo$)
'       Do Until vol_RecordCombo.EOF
'          If IsMissing(ListaDepartamento) = True Then
'             For vil_CountLista = 1 To ListaDepartamento.ListItems.Count
'                 If ListaDepartamento.ListItems(vil_CountLista).SubItems(2) = vol_RecordCombo(Campo$) Then
'                    Combo.AddItem vol_RecordCombo(Campo$)
'                    Exit For
'                 End If
'             Next vil_CountLista
'          Else
'             Combo.AddItem vol_RecordCombo(Campo$)
'          End If
'          vol_RecordCombo.MoveNext
'       Loop
'       Set vol_RecordCombo = Nothing
'   vol_Conexao.FecharConexao
'
'TrataErros:
'    If Err.Number = 5 Then
'        Set vol_RecordCombo = Nothing
'        vol_Conexao.FecharConexao
'        Err.Clear
'        Resume Next
'    End If
'
'End Function

Function DataHoraAtual() As String
Dim vsl_Retorno As String
Dim vol_Record As New ADODB.Recordset          'Conexão do Banco de Dados

On Error GoTo TrataErros
    
    vol_Conexao.AbrirConexao
       Set vol_Record = vol_Conexao.RetornaRs("SELECT FORMAT (GETDATE(), 'dd/MM/yyyy HH:mm:ss', 'pt-br') as DataHoraAtual")
       If vol_Record.EOF = False Then
         DataHoraAtual = vol_Record("DataHoraAtual")
       End If
       Set vol_Record = Nothing
   vol_Conexao.FecharConexao

TrataErros:
    If Err.Number = 5 Then
        Set vol_Record = Nothing
        vol_Conexao.FecharConexao
        Err.Clear
        Resume Next
    End If

End Function

Function CarregaCombo(Combo As Object, Tabela As String, ID As String, Descricao As String, Condicao As String)
Dim vsl_Query As String
Dim vcl_SqlCombo As String
Dim vil_CountLista As Integer
Dim vol_RecordCombo As New ADODB.Recordset          'Conexão do Banco de Dados

On Error GoTo TrataErros
    
    vol_Conexao.AbrirConexao
       Combo.Clear
       vsl_Query = "SELECT " & ID & ", " & Descricao & " FROM " & Tabela & " "
       vsl_Query = vsl_Query & Condicao
       Set vol_RecordCombo = vol_Conexao.RetornaRs(vsl_Query)
       Do Until vol_RecordCombo.EOF
          Combo.AddItem Trim$(vol_RecordCombo(Descricao))
          Combo.ItemData(Combo.NewIndex) = vol_RecordCombo(ID)
          vol_RecordCombo.MoveNext
       Loop
       Set vol_RecordCombo = Nothing
   vol_Conexao.FecharConexao

TrataErros:

    If Err.Number <> 0 Then
        Set vol_RecordCombo = Nothing
        vol_Conexao.FecharConexao
        Err.Clear
        Resume Next
    End If

End Function

