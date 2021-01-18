Attribute VB_Name = "Biblioteca"
Global dboDESAFIO As New ADODB.Connection
   
Dim vDataSource  As String
Dim vInitialCatalog As String
Dim vUserid As String
Dim vPassword As String

Public Sub Main()
Dim rstLig As Recordset
Dim IntSize As Long

On Error GoTo Err_Main
    
    
    AbrirArquivoIni
    dboDESAFIO.Open "Provider=SQLOLEDB.1;Password=" & vPassword & ";Persist Security Info=True;User Id=" & vUserid & ";Initial Catalog=" & vInitialCatalog & ";Data Source=" & vDataSource
    WriteEvent ("Sistema inicializado com sucesso.")
    frmMain.Show
    

Ok_Main:

Exit Sub



Err_Main:

Select Case Err.Number
    Case 3024
         MsgBox "Não foi possível iniciar o sistema devido a falta do banco de dados."
         WriteEvent ("Erro ao inicializar")
         
    Case Else
         MsgBox Err.Description
         WriteEvent ("Erro ao inicializar")
End Select

Resume Ok_Main

End Sub

Function RetiraMascara(v As String)

For x = 1 To Len(v)
    If Mid(v, x, 1) <> "-" And Mid(v, x, 1) <> "&" And Mid(v, x, 1) <> "\" And Mid(v, x, 1) <> "*" And Mid(v, x, 1) <> "!" And Mid(v, x, 1) <> "@" And Mid(v, x, 1) <> "#" And Mid(v, x, 1) <> "$" And Mid(v, x, 1) <> "%" And Mid(v, x, 1) <> "(" And Mid(v, x, 1) <> ")" And Mid(v, x, 1) <> "_" And Mid(v, x, 1) <> "-" And Mid(v, x, 1) <> "=" And Mid(v, x, 1) <> "+" And Mid(v, x, 1) <> "}" And Mid(v, x, 1) <> "]" And Mid(v, x, 1) <> "[" And Mid(v, x, 1) <> "{" And Mid(v, x, 1) <> "^" And Mid(v, x, 1) <> "~" And Mid(v, x, 1) <> "?" And Mid(v, x, 1) <> "<" And Mid(v, x, 1) <> ">" And Mid(v, x, 1) <> ";" And Mid(v, x, 1) <> "|" And Mid(v, x, 1) <> "'" Then
       RetiraMascara = RetiraMascara & Mid(v, x, 1)
    End If
Next

End Function


Public Sub WriteEvent(EventName)
Dim FF As Long

FF = FreeFile

Open App.Path & "\LOG" & Format(Date, "yyyymmdd") & ".log" For Append As FF
Print #FF, Format(Date, "dd-mm-yyyy hh:mm") & vtab & vtab & EventName
Close #FF

End Sub

Public Sub AbrirArquivoIni()
Dim F As Long, Linha As String
Dim db As Database, rs As Recordset

F = FreeFile
Open App.Path & "\Desafio.ini" For Input As F   'abre o arquivo texto


Do While Not EOF(F)
   Line Input #F, Linha 'lê uma linha do arquivo texto

  'extrai a informação do arquivo texto usando a função MID
   If Mid(Linha, 1, 11) = "Data Source" Then
      vDataSource = Mid(Linha, 13, Len(Linha) - 12)
   
   ElseIf Mid(Linha, 1, 15) = "Initial Catalog" Then
      vInitialCatalog = Mid(Linha, 17, Len(Linha) - 16)
   
   ElseIf Mid(Linha, 1, 7) = "User Id" Then
      vUserid = Mid(Linha, 9, Len(Linha) - 8)
   
   ElseIf Mid(Linha, 1, 8) = "Password" Then
      vPassword = Mid(Linha, 10, Len(Linha) - 9)
   End If
Loop


On Error Resume Next    'se a tabela não existir escapa da mensagem de erro



End Sub

