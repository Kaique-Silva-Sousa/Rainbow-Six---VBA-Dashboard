VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addDados 
   Caption         =   "Adicionar Dados"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13860
   OleObjectBlob   =   "useForm - R6.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim endereco As String
Dim ws As Worksheet
Dim ultimaLinha As Long
Private Sub btnDesfazer_Click()
    Set ws = ThisWorkbook.Sheets("Dados")
    ultimaLinha = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    If ultimaLinha = 1 Then
        MsgBox ("Não é Possível Apagar")
    Else
        Range("D" & ultimaLinha) = ""
        Range("E" & ultimaLinha) = ""
        Range("F" & ultimaLinha) = ""
        Range("G" & ultimaLinha) = ""
    End If
End Sub
Private Sub btnEnviar_Click()
    Set ws = ThisWorkbook.Sheets("Dados") '

    ultimaLinha = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    If (Me.timeListagem.Text = "") Or (Me.mapaListagem.Text = "") Or (Me.vitoriaCaixa.Text = "") Or (Me.prorrogacaoCaixa.Text = "") Then
        MsgBox ("Todos os Campos Precisam Estar Preenchidos!")
    Else
        Range("D" & ultimaLinha + 1) = Me.timeListagem.Text
        Range("E" & ultimaLinha + 1) = Me.mapaListagem.Text
        Range("F" & ultimaLinha + 1) = Me.vitoriaCaixa.Text
        Range("G" & ultimaLinha + 1) = Me.prorrogacaoCaixa.Text
    
    End If
     
End Sub


Private Sub mapaListagem_Change()
endereco = Application.ThisWorkbook.Path & "\imagens\"
Me.fotoTime.PictureSizeMode = fmPictureSizeModeStretch

Select Case mapaListagem.Text

Case "Banco"
    Me.fotoMapa.Picture = LoadPicture(endereco & "banco" & ".JPG")
    
Case "Chale"
    Me.fotoMapa.Picture = LoadPicture(endereco & "chale" & ".JPG")
    
Case "Laboratorio"
    Me.fotoMapa.Picture = LoadPicture(endereco & "labs" & ".JPG")
    
Case "Oregon"
    Me.fotoMapa.Picture = LoadPicture(endereco & "oregon" & ".JPG")
    
Case "Consulado"
    Me.fotoMapa.Picture = LoadPicture(endereco & "consul" & ".JPG")
    
Case "Kafe"
    Me.fotoMapa.Picture = LoadPicture(endereco & "kafe" & ".JPG")
    
Case "Arranha-Ceu"
    Me.fotoMapa.Picture = LoadPicture(endereco & "arranha" & ".JPG")
    
Case "Fronteira"
    Me.fotoMapa.Picture = LoadPicture(endereco & "fronteira" & ".JPG")
    
Case "Clube"
    Me.fotoMapa.Picture = LoadPicture(endereco & "clube" & ".JPG")
    
End Select


End Sub

Private Sub timeListagem_Change()
endereco = Application.ThisWorkbook.Path & "\imagens\"
Me.fotoTime.PictureSizeMode = fmPictureSizeModeStretch

Select Case timeListagem.Text

Case "Team Liquid"
    Me.fotoTime.Picture = LoadPicture(endereco & "tl" & ".JPG")
    
Case "Faze Clan"
    Me.fotoTime.Picture = LoadPicture(endereco & "faze" & ".JPG")
    
Case "Fluxo"
    Me.fotoTime.Picture = LoadPicture(endereco & "fluxo" & ".JPG")
    
Case "Black Dragons"
    Me.fotoTime.Picture = LoadPicture(endereco & "bd" & ".JPG")
    
Case "E1 Sports"
    Me.fotoTime.Picture = LoadPicture(endereco & "e1" & ".JPG")
    
Case "Furia"
    Me.fotoTime.Picture = LoadPicture(endereco & "furia" & ".JPG")
    
Case "MIBR"
    Me.fotoTime.Picture = LoadPicture(endereco & "mibr" & ".JPG")
    
Case "Ninjas in Pyjamas"
    Me.fotoTime.Picture = LoadPicture(endereco & "nip" & ".JPG")
    
Case "Vivo Keyd Stars"
    Me.fotoTime.Picture = LoadPicture(endereco & "keyd" & ".JPG")
    
Case "W7M Sports"
    Me.fotoTime.Picture = LoadPicture(endereco & "W7M" & ".JPG")
End Select




End Sub

Private Sub UserForm_Initialize()

Me.timeListagem.RowSource = "A2:A11"
Me.timeListagem.Font.Bold = True
Me.timeListagem.Font.Size = 16


Me.mapaListagem.RowSource = "B2:B10"
Me.mapaListagem.Font.Bold = True
Me.mapaListagem.Font.Size = 16

With addDados.prorrogacaoCaixa

    .AddItem "Sim"
    .AddItem "Não"
    .Font.Bold = True
    .Font.Size = 14

End With

With addDados.vitoriaCaixa

    .AddItem "Sim"
    .AddItem "Não"
    .Font.Bold = True
    .Font.Size = 14

End With


End Sub

