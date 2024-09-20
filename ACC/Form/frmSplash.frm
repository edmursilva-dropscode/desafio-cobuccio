VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4410
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6525
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5775
      Top             =   3480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "XYZ      Administradora de Cartões de Crédito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   300
      TabIndex        =   3
      Top             =   435
      Width           =   5220
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1140
      Index           =   0
      Left            =   750
      TabIndex        =   2
      Top             =   1455
      Width           =   4950
   End
   Begin VB.Label lblTitulo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1995
      TabIndex        =   1
      Top             =   4140
      Width           =   4440
   End
   Begin VB.Shape Barra 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   0
      Left            =   0
      Top             =   4125
      Width           =   7095
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo Cobuccio"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   1155
   End
   Begin VB.Shape Barra 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   2
      Left            =   0
      Top             =   -15
      Width           =   2505
   End
   Begin VB.Shape Barra 
      BackColor       =   &H00808080&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   300
      Index           =   1
      Left            =   0
      Top             =   -15
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   4140
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   7635
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

On Error GoTo TrataErro

    lblTitulo(0).Caption = App.FileDescription
    lblTitulo(2).Caption = "Versão " & App.Major & "." & App.Minor
    Call VerTela
    Timer1.Enabled = True
    
TrataErro:
    If Err.Number <> 0 Then TrataErros
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSplash = Nothing
End Sub


Private Sub Timer1_Timer()
On Error GoTo TrataErro
    DoEvents
    Load frmMain
    DoEvents
    frmMain.Show
    Unload frmSplash
TrataErro:
    If Err.Number <> 0 Then TrataErros
End Sub

