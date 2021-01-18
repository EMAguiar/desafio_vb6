VERSION 5.00
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00E0E0E0&
   Caption         =   "Modulo para consumir WEB API em VB6"
   ClientHeight    =   8295
   ClientLeft      =   75
   ClientTop       =   585
   ClientWidth     =   11880
   LinkMode        =   1  'Source
   LinkTopic       =   "Main"
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMain 
      Interval        =   60000
      Left            =   840
      Top             =   7320
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Arquivo"
      WindowList      =   -1  'True
      Begin VB.Menu mnuFile1 
         Caption         =   "&Carregar dados"
      End
      Begin VB.Menu mnuFile2 
         Caption         =   "&Banco de dados"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "S&air"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Unload(Cancel As Integer)
   WriteEvent "Saindo do Sistema."
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFile1_Click()
    frmCountries.Show
End Sub

Private Sub mnuFile2_Click()
    frmDados.Show
End Sub
