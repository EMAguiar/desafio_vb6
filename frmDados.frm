VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dados do País"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   6720
      Width           =   7935
      Begin VB.CommandButton btDados 
         Caption         =   "&Sair"
         Height          =   255
         Index           =   5
         Left            =   6960
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Buscar"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Excluir"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Alterar"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Novo"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pesquisa"
      Height          =   2775
      Left            =   0
      TabIndex        =   17
      Top             =   3960
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid grdDados 
         Height          =   2415
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4260
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Languages"
      Height          =   1575
      Left            =   0
      TabIndex        =   15
      Top             =   2400
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid grdLanguages 
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   2143
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "País"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7935
      Begin VB.TextBox txtCountryFlag 
         Height          =   300
         Left            =   1080
         TabIndex        =   6
         Top             =   1920
         Width           =   6735
      End
      Begin VB.TextBox txtCurrencyISOCode 
         Height          =   300
         Left            =   6960
         TabIndex        =   5
         Top             =   1485
         Width           =   855
      End
      Begin VB.TextBox txtContinentCode 
         Height          =   300
         Left            =   3720
         TabIndex        =   4
         Top             =   1485
         Width           =   855
      End
      Begin VB.TextBox txtPhoneCode 
         Height          =   300
         Left            =   1080
         TabIndex        =   3
         Top             =   1485
         Width           =   855
      End
      Begin VB.TextBox txtCapitalCity 
         Height          =   300
         Left            =   1080
         TabIndex        =   2
         Top             =   1005
         Width           =   4695
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   645
         Width           =   4695
      End
      Begin VB.TextBox txtISOCode 
         Height          =   300
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Country Flag"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1965
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Currency ISOCode"
         Height          =   195
         Left            =   5400
         TabIndex        =   12
         Top             =   1605
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Continent Code"
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   1605
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Phone Code"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1605
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Capital City"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1125
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ISO Code"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs  As ADODB.Recordset
Dim Sql As String

Private Sub Form_Load()
   'Centraliza a tela
    Me.Left = frmMain.Left + ((frmMain.Width - Me.Width) / 2)
    Me.Top = frmMain.Top + ((frmMain.Height - Me.Height) / 2) - 1000
    
   'Monta a grid de Dados e Idiomas
   'e carrega a grid de dados
    Call Montar_Dados
    Call Carregar_dados
    Call Montar_Languages
    
End Sub

Private Sub btDados_Click(Index As Integer)
    Select Case Index
           Case 5: Unload Me 'Sair
    End Select
End Sub

Private Sub Montar_Languages()
    With grdLanguages
         .Row = 0
         .cols = 3
         .ColWidth(0) = 300:  .TextMatrix(0, 0) = "S"
         .ColWidth(1) = 800:  .TextMatrix(0, 1) = "ISOCode"
         .ColWidth(2) = 6300: .TextMatrix(0, 2) = "Name"
    End With
End Sub

Private Sub Montar_Dados()
    With grdDados
         .Row = 0
         .cols = 8
         .ColWidth(0) = 300:  .TextMatrix(0, 0) = "S"
         .ColWidth(1) = 800:  .TextMatrix(0, 1) = "ISOCode"
         .ColWidth(2) = 2000: .TextMatrix(0, 2) = "Name"
         .ColWidth(3) = 2000: .TextMatrix(0, 3) = "Capital City"
         .ColWidth(4) = 1000: .TextMatrix(0, 4) = "Phone Code"
         .ColWidth(5) = 1500: .TextMatrix(0, 5) = "Continent Code"
         .ColWidth(6) = 2000: .TextMatrix(0, 6) = "Current ISO Code"
         .ColWidth(7) = 6300: .TextMatrix(0, 7) = "Country Flag"
    End With
End Sub

Sub Carregar_dados()
Dim Linha  As Integer
Linha = 1

Set rs = New ADODB.Recordset
Set rs.ActiveConnection = dboDESAFIO
rs.CursorLocation = adUseClient
Sql = " Select * from FullCountryInfoAllCountries order by ISOCodeC"

rs.Open Sql

With rs
     If Not .EOF Then
        Do While Not .EOF
           With grdDados
               .rows = Linha + 1
               .Row = Linha
               .TextMatrix(Linha, 1) = rs.fields("ISOCodeC")
               .TextMatrix(Linha, 2) = rs.fields("Name")
               .TextMatrix(Linha, 3) = rs.fields("CapitalCity")
               .TextMatrix(Linha, 4) = rs.fields("PhoneCode")
               .TextMatrix(Linha, 5) = rs.fields("ContinentCode")
               .TextMatrix(Linha, 6) = rs.fields("CurrencyISOCode")
               .TextMatrix(Linha, 7) = rs.fields("CountryFlag")
                Linha = Linha + 1
                rs.MoveNext
           End With
        Loop
     End If
End With

rs.Close

End Sub


Private Sub grdDados_DblClick()
    With grdDados
         If .MouseCol = 0 Then
            .Col = 0
             If .TextMatrix(.RowSel, 0) = "X" Then
                .TextMatrix(.RowSel, 0) = ""
                 LimparTela
             Else
                .TextMatrix(.RowSel, 0) = "X"
                 MostrarCampos (.RowSel)
             End If
         End If
    End With
End Sub

Sub LimparTela()
    txtISOCode = ""
    txtName = ""
    txtCapitalCity = ""
    txtPhoneCode = ""
    txtContinentCode = ""
    txtCurrencyISOCode = ""
    txtCountryFlag = ""
    grdLanguages.rows = 1
End Sub

Function MostrarCampos(Linha As Integer)
Dim Linha2 As Integer
Linha2 = 1

    With grdDados
         txtISOCode = .TextMatrix(Linha, 1)
         txtName = .TextMatrix(Linha, 2)
         txtCapitalCity = .TextMatrix(Linha, 3)
         txtPhoneCode = .TextMatrix(Linha, 4)
         txtContinentCode = .TextMatrix(Linha, 5)
         txtCurrencyISOCode = .TextMatrix(Linha, 6)
         txtCountryFlag = .TextMatrix(Linha, 7)
    End With
        
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = dboDESAFIO
    rs.CursorLocation = adUseClient
    Sql = " Select * from Languages where ISOCODEC = '" & txtISOCode & "' order by ISOCodeC, ISOCode"
    
    rs.Open Sql
    
    If Not rs.EOF Then
        Do While Not rs.EOF
           With grdLanguages
               .rows = Linha2 + 1
               .Row = Linha2
               .TextMatrix(Linha2, 1) = rs.fields("ISOCode")
               .TextMatrix(Linha2, 2) = rs.fields("Name")
                Linha2 = Linha2 + 1
                rs.MoveNext
           End With
        Loop
        rs.Close
    End If
    
End Function
