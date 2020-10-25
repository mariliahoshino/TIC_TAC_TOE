VERSION 5.00
Begin VB.Form frmJVelha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jogo da Velha 1"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   Icon            =   "frmJVelha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4000
   ScaleMode       =   0  'User
   ScaleWidth      =   4000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBot 
      Height          =   855
      Index           =   8
      Left            =   2400
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdBot 
      Height          =   855
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdBot 
      Height          =   855
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdBot 
      Height          =   855
      Index           =   5
      Left            =   2400
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdBot 
      Height          =   855
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdBot 
      Height          =   855
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdBot 
      Height          =   855
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdBot 
      Height          =   855
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdBot 
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmJVelha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumeroVezes As Integer

Private Sub cmdBot_Click(Index As Integer)  'O parâmetro Index informa o index acionado
    NumeroVezes = NumeroVezes + 1
    cmdBot(Index).Caption = "X"
    
    If cmdBot(0).Caption = "X" And cmdBot(1).Caption = "X" And cmdBot(2).Caption = "X" Then GoTo MensX 'Verifica se umas
    If cmdBot(3).Caption = "X" And cmdBot(4).Caption = "X" And cmdBot(5).Caption = "X" Then GoTo MensX
    If cmdBot(6).Caption = "X" And cmdBot(7).Caption = "X" And cmdBot(8).Caption = "X" Then GoTo MensX
    If cmdBot(0).Caption = "X" And cmdBot(3).Caption = "X" And cmdBot(6).Caption = "X" Then GoTo MensX  'Dessas combinações são válidas
    If cmdBot(1).Caption = "X" And cmdBot(4).Caption = "X" And cmdBot(7).Caption = "X" Then GoTo MensX  'se sim vai à MensX
    If cmdBot(2).Caption = "X" And cmdBot(5).Caption = "X" And cmdBot(8).Caption = "X" Then GoTo MensX
    If cmdBot(0).Caption = "X" And cmdBot(4).Caption = "X" And cmdBot(8).Caption = "X" Then GoTo MensX
    If cmdBot(2).Caption = "X" And cmdBot(4).Caption = "X" And cmdBot(6).Caption = "X" Then GoTo MensX
    
    Do Until cmdBot(M).Caption = ""
        M = Int(Rnd * 9)
    Loop
    cmdBot(M).Caption = "O"
    
    If cmdBot(0).Caption = "O" And cmdBot(1).Caption = "O" And cmdBot(2).Caption = "O" Then GoTo MensO  'Verifica se umas
    If cmdBot(3).Caption = "O" And cmdBot(4).Caption = "O" And cmdBot(5).Caption = "O" Then GoTo MensO
    If cmdBot(6).Caption = "O" And cmdBot(7).Caption = "O" And cmdBot(8).Caption = "O" Then GoTo MensO
    If cmdBot(0).Caption = "O" And cmdBot(3).Caption = "O" And cmdBot(6).Caption = "O" Then GoTo MensO  'Dessas combinações são válidas
    If cmdBot(1).Caption = "O" And cmdBot(4).Caption = "O" And cmdBot(7).Caption = "O" Then GoTo MensO
    If cmdBot(2).Caption = "O" And cmdBot(5).Caption = "O" And cmdBot(8).Caption = "O" Then GoTo MensO
    If cmdBot(0).Caption = "O" And cmdBot(4).Caption = "O" And cmdBot(8).Caption = "O" Then GoTo MensO
    If cmdBot(2).Caption = "O" And cmdBot(4).Caption = "O" And cmdBot(6).Caption = "O" Then GoTo MensO  'se sim vai à MensX
    
    If NumeroVezes = 4 Then GoTo Empate 'Se a variável for igual à 4 vai à Empate
    Exit Sub
    
Empate:
    Resposta$ = MsgBox("Partida Empatada, Deseja" + Chr(13) + "jogar novamente?", 68, "Vencedor")
    If Resposta$ = 6 Then
    JogoNovo
    Else
        End
    End If
    Exit Sub
    
MensX:
    Resposta$ = MsgBox("Você Ganhou, Deseja" + Chr(13) + "jogar novamente?", 68, "Vencedor")
    If Resposta$ = 6 Then
    JogoNovo
    Else
        End
    End If
    Exit Sub
    
MensO:
    Resposta$ = MsgBox("Eu Ganhei, Deseja" + Chr(13) + "jogar novamente?", 68, "Vencedor")
    If Resposta$ = 6 Then
    JogoNovo
    Else
        End
    End If
    Exit Sub
    
End Sub

Private Sub Form_Load()
    Randomize
    JogoNovo
End Sub

Public Sub JogoNovo()
    For i% = 0 To 8
        cmdBot(i%).Caption = ""
    Next i%
    M = Int(Rnd * 9)
    cmdBot(M).Caption = "O" 'jogada inicial
    NumeroVezes = 0         'do micro
End Sub
