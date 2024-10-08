VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstatisticaNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Inicializa vari�veis da classe
Private vop_EstatisticaDados As New clsEstatisticaDados      'Vari�vel de acesso classe clsEstatisticaDados

'Vari�veis da classe
Private vsp_Query As String                'Vari�vel usada nas consultas sql
Private vip_Count As Integer               'Vari�vel utilizadas nos loops
Private vbp_RollBack As Boolean            'Define que uma conex�o est� aberta e necessita de um rollback



Public Function LocalizarEstatisticas(ByRef pdtgEstatisticas As DataGrid, ByRef padoEstatisticas As Adodc, ByVal ptxtLocalizarEstatisticas As String, ByVal pcmbLocalizarEstatisticas As Integer, ByVal pDataInicial As String, ByVal pDataFinal As String) As Boolean

    LocalizarEstatisticas = False
    Screen.MousePointer = vbHourglass
      If vop_EstatisticaDados.LocalizarEstatistica(pdtgEstatisticas, padoEstatisticas, ptxtLocalizarEstatisticas, pcmbLocalizarEstatisticas, pDataInicial, pDataFinal) = True Then
         LocalizarEstatisticas = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function CarregarGridEstatisticaMainRS(ByRef pdtgEstatisticas As DataGrid, ByRef padoEstatisticas As Adodc, ByVal ptxtLocalizarEstatisticas As String, ByVal pcmbLocalizarEstatisticas As Integer, ByVal pDataInicial As String, ByVal pDataFinal As String) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridEstatisticaMainRS = vop_EstatisticaDados.CarregarGridEstatisticaMainRS(pdtgEstatisticas, padoEstatisticas, ptxtLocalizarEstatisticas, pcmbLocalizarEstatisticas, pDataInicial, pDataFinal)
    Screen.MousePointer = vbDefault
    
End Function
