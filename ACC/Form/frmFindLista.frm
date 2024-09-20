VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmFindLista 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   ClipControls    =   0   'False
   Icon            =   "frmFindLista.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5235
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ListView lvwLista 
      Height          =   3050
      Left            =   90
      TabIndex        =   1
      Top             =   1350
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5371
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "."
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   6827
      EndProperty
   End
   Begin VB.TextBox txtFind 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   90
      TabIndex        =   0
      Top             =   1020
      Width           =   5055
   End
   Begin LVbuttons.LaVolpeButton lvbFechar 
      Height          =   360
      Left            =   4150
      TabIndex        =   6
      Top             =   4480
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Fechar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmFindLista.frx":000C
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   0   'False
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton lvbSelecionar 
      Height          =   360
      Left            =   3125
      TabIndex        =   7
      Top             =   4480
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Selecionar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmFindLista.frx":0028
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   0   'False
      BSTYLE          =   0
   End
   Begin VB.OptionButton optSort02 
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   4540
      Width           =   1560
   End
   Begin VB.OptionButton optSort01 
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   4540
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.Label lblTitulo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localizar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3820
      TabIndex        =   5
      Tag             =   "175"
      Top             =   180
      Width           =   1320
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   100
      Picture         =   "frmFindLista.frx":0044
      Top             =   100
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o texto a ser localizado:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   780
      Width           =   2115
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -2640
      Picture         =   "frmFindLista.frx":090E
      Top             =   675
      Width           =   10740
   End
   Begin VB.Image imgBarra 
      Height          =   795
      Left            =   -255
      Picture         =   "frmFindLista.frx":1292
      Top             =   -100
      Width           =   8250
   End
End
Attribute VB_Name = "frmFindLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vip_ItemLista As Integer
Dim vol_ListItem As ListItem               'Usada para manipular itens em um listview

Private Sub Form_Activate()
   
    Me.Refresh
    
End Sub

Private Sub Form_Load()

    vip_ItemLista = 1
    'Código usado para formatar o ListView de pesquisa
    Set vol_ListItem = lvwLista.ListItems.Add(, , "")
    Set vol_ListItem = Nothing
    lvwLista.ListItems.Clear
    '
    lvwLista.SortOrder = lvwAscending
    lvwLista.SortKey = lvwLista.ColumnHeaders(1).Index - 1
    lvwLista.Sorted = True
    txtFind.text = Empty
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    If KeyCode = vbKeyInsert Then Call lvbSelecionar_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmFindLista = Nothing

End Sub


Private Sub lvbFechar_Click()
    
    If frmFindLista.lvwLista.ListItems.Count > 0 Then
       frmFindLista.lvwLista.SelectedItem.text = Empty
       Me.Hide
    Else
       Unload Me
    End If
    
End Sub

Private Sub lvbSelecionar_Click()

On Error GoTo TrataErros

    Me.Hide
    
TrataErros:
    If Err.Number <> 0 Then Exit Sub
    
End Sub

Private Sub lvwLista_DblClick()
    
    Call lvbSelecionar_Click

End Sub

Private Sub lvwLista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    vip_ItemLista = Item.Index

End Sub

Private Sub lvwLista_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call lvbSelecionar_Click
    End If

End Sub

Private Sub optSort01_Click()

On Error GoTo TrataErros

    If optSort01.Value = True Then
       lvwLista.SortOrder = lvwAscending
       lvwLista.SortKey = lvwLista.ColumnHeaders(1).Index - 1
       lvwLista.Sorted = True
       txtFind.text = Empty
       txtFind.SetFocus
    End If

TrataErros:
    If Err.Number = 5 Then Exit Sub

End Sub

Private Sub optSort02_Click()
    
    If optSort02.Value = True Then
       lvwLista.SortOrder = lvwAscending
       lvwLista.SortKey = lvwLista.ColumnHeaders(2).Index - 1
       lvwLista.Sorted = True
       txtFind.text = Empty
       'txtFind.SetFocus
    End If
    
End Sub

Private Sub txtFind_Change()
Dim vil_Searchlen As Integer
Dim vsl_Search As String

    If Trim$(txtFind.text) = Empty Then Exit Sub
    vsl_Search$ = UCase$(txtFind.text)
    vil_Searchlen = Len(vsl_Search$)

On Error GoTo TrataErros

    If vil_Searchlen Then
        If optSort01.Value = False Then
            For vip_ItemLista = 1 To lvwLista.ListItems.Count
                If UCase$(Left$(lvwLista.ListItems.Item(vip_ItemLista).SubItems(1), vil_Searchlen)) = vsl_Search$ Then
                    lvwLista.ListItems(vip_ItemLista).EnsureVisible
                    lvwLista.ListItems(vip_ItemLista).Selected = True
                    Exit For
                End If
            Next
        Else
            For vip_ItemLista = 1 To lvwLista.ListItems.Count
                If UCase$(Left$(lvwLista.ListItems.Item(vip_ItemLista), vil_Searchlen)) = vsl_Search$ Then
                    lvwLista.ListItems(vip_ItemLista).EnsureVisible
                    lvwLista.ListItems(vip_ItemLista).Selected = True
                    Exit For
                End If
            Next
        End If
    End If

TrataErros:
    Exit Sub

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call lvbSelecionar_Click
    End If

End Sub

