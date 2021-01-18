VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form RCountry 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informações dos Países"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11085
   ScaleWidth      =   14115
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox rtfXML 
      Height          =   9615
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   20055
      _ExtentX        =   35375
      _ExtentY        =   16960
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"RCountry.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Obter dados do Web  API http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   13815
      Begin VB.CommandButton btn_WEBAPI 
         Caption         =   "Sa&ir"
         Height          =   450
         Index           =   2
         Left            =   3600
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton btn_WEBAPI 
         Caption         =   "&Salvar dados"
         Height          =   450
         Index           =   1
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton btn_WEBAPI 
         Caption         =   "&Baixar dados"
         Height          =   450
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "RCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_objDOMPessoa  As DOMDocument60
Private m_strXmlPath    As String

Dim objPessoaRoot       As IXMLDOMElement
Dim objPessoaElement    As IXMLDOMElement

Dim vISOCode As String
Dim vName As String
Dim vCapitalCity As String
Dim vPhoneCode As String
Dim vsContinentCode As String
Dim vsCurrencyISOCode As String
Dim vsCountryFlag As String

Dim xhr

Private Sub btn_WEBAPI_Click(Index As Integer)
Dim method, url, contents, formatcontent, doc


Select Case Index
       Case 0: 'Baixar Dados

            Set xhr = CreateObject("MSXML2.XMLHTTP")

            method = "GET" 'Escolhe o método HTTP de envio
            url = "http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries" 'url da API
            contents = "length" 'conteudo
            formatcontent = "application/xml" 'Se a API usar outro formato basta alterar aqui

            xhr.Open method, url, False

           'Retorno XML ao invés de JSON
            xhr.setRequestHeader "Accept", "application/xml"

            If method = "POST" Or method = "PUT" Then
               xhr.setRequestHeader "Content-Type", formatcontent
               xhr.setRequestHeader "Content-Length", Len(contents)
               xhr.send contents
            Else
               xhr.send
            End If

            If xhr.Status < 200 Or xhr.Status >= 300 Then
              'Algo falhou, as vezes pode haver uma descrição em `xhr.responseText` ou pode retornar vazio, o `xhr.status` indica o tipo de erro
               MsgBox "Erro HTTP:" & xhr.Status & " - Detalhes: " & xhr.responseText
            Else
              'Faz o parse da String para XML
               Set doc = CreateObject("MSXML2.DOMDocument")
               doc.loadXML (xhr.responseText)


'REMOVER sISOCOADE <> 'A'
'ROTINA
'              'Elementos encontrados para tCountryInfo: " & nodes.length
'               Set Nodes = doc.selectNodes("//tCountryInfo")
'
'               For Each Node In Nodes
'                  'MsgBox "tCountryInfo " & Node.Text
'
'                   If Mid(Node.Text, 1, 1) <> "A" Then
'                     'Get the node you want to delete
'                      Set objnode = doc.selectSingleNode("//tCountryInfo")
'
'                     'Find the parent node of the node you want to delete
'                     'and call the removechild method of that node
'                      objnode.parentNode.removeChild objnode
'                      objnode.removeChild objnode
'                    End If
'               Next
               
              
               rtfXML.Text = xhr.responseText
               doc.save "c:\teste.xml"
            End If
            
       Case 1:
            If Not IsEmpty(xhr) Then
               Call SalvarDB
            Else
               MsgBox "Não é possível salvar os dados, pois não existem dados baixados.", vbInformation
            End If
       Case 2: Unload Me
       
End Select

End Sub


Private Sub SalvarDB()
Dim objPessoaRoot     As IXMLDOMElement
Dim objPessoaElement  As IXMLDOMElement
Dim objNameNode       As IXMLDOMNode
Dim vNode             As Integer 'Total de Nodes tCountryInfo
  
  Set m_objDOMPessoa = New DOMDocument60
  m_objDOMPessoa.resolveExternals = True
  m_objDOMPessoa.validateOnParse = True

  'carrega o XML no documento DOM
  m_objDOMPessoa.async = False
  m_objDOMPessoa.loadXML (xhr.responseText)

  'verifica se a carga do XML foi feita com sucesso
  If m_objDOMPessoa.parseError.reason <> "" Then
    MsgBox m_objDOMPessoa.parseError.reason
    Exit Sub
  End If

  dboDESAFIO.Execute "DELETE FROM FullCountryInfoAllCountries"
  dboDESAFIO.Execute "DELETE FROM Languages"
 
 'Elementos encontrados para tCountryInfo: " & nodes.length
  Set Nodes = m_objDOMPessoa.selectNodes("//tCountryInfo")
  vNode = 0
  
  For Each Node In Nodes
'      MsgBox "tCountryInfo " & Node.Text
'
'      If Mid(Node.Text, 1, 1) = "A" Then
'        'Get the node you want to delete
'         Set objnode = doc.selectSingleNode("//tCountryInfo")
'
'        'Find the parent node of the node you want to delete
'        'and call the removechild method of that node
'         objnode.parentNode.removeChild objnode
'         objnode.removeChild objnode
'      End If
' Next
    
    If Mid(Node.Text, 1, 1) = "A" Then
      'obtem o elemento raiz do XML
       Set objPessoaRoot = m_objDOMPessoa.documentElement.childNodes(vNode)

       For Each objPessoaElement In objPessoaRoot.childNodes
           objPessoaElement.hasChildNodes

           If objPessoaElement.baseName = "sISOCode" Then
              vISOCode = objPessoaElement.nodeTypedValue

           ElseIf objPessoaElement.baseName = "sName" Then
              vName = RetiraMascara(objPessoaElement.nodeTypedValue)
                        
           ElseIf objPessoaElement.baseName = "sCapitalCity" Then
              vCapitalCity = RetiraMascara(objPessoaElement.nodeTypedValue)

           ElseIf objPessoaElement.baseName = "sPhoneCode" Then
              vPhoneCode = objPessoaElement.nodeTypedValue

           ElseIf objPessoaElement.baseName = "sContinentCode" Then
              vContinentCode = objPessoaElement.nodeTypedValue

           ElseIf objPessoaElement.baseName = "sCurrencyISOCode" Then
              vCurrencyISOCode = objPessoaElement.nodeTypedValue

           ElseIf objPessoaElement.baseName = "sCountryFlag" Then
              vCountryFlag = RetiraMascara(objPessoaElement.nodeTypedValue)

           ElseIf objPessoaElement.baseName = "Languages" Then
              dboDESAFIO.Execute "INSERT INTO FullCountryInfoAllCountries (ISOCodeC, Name, CapitalCity, PhoneCode, ContinentCode, CurrencyISOCode, CountryFlag) values ('" & vISOCode & "', '" & vName & "', '" & vCapitalCity & "', '" & vPhoneCode & "','" & vContinentCode & "', '" & vCurrencyISOCode & "','" & vCountryFlag & "')"
              SalvaLanguages objPessoaElement

           End If
       Next
    End If
    vNode = vNode + 1

Next

MsgBox "XML Salvo com sucesso", vbInformation, "Atenção"
  
End Sub

Private Sub SalvaLanguages(objDOMNode As IXMLDOMElement)
  Dim objNameNode      As IXMLDOMNode
  Dim objPessoaElement As IXMLDOMElement
  Dim vISOCodeL As String
  Dim vNameL As String
  
 
 'obtem o nome do elemento selecionado
  Set objNameNode = objDOMNode.selectSingleNode("Language")
 
  For Each objPessoaElement In objDOMNode.childNodes
      vISOCodeL = objPessoaElement.childNodes(0).nodeTypedValue
      vNameL = objPessoaElement.childNodes(1).nodeTypedValue
            
      dboDESAFIO.Execute "INSERT INTO Languages (ISOCodeC, ISOCode, Name) values ('" & vISOCode & "', '" & vISOCodeL & "', '" & vNameL & "')"
  Next
  
End Sub


'Recebe assincronamente o resultado
Sub doReadyStateChange()
    If xhr.ReadyState = 4 Then
        If xhr.Status < 200 Or xhr.Status >= 300 Then
            MsgBox "Erro HTTP:" & xhr.Status & " - Detalhes: " & xhr.responseText
        Else
            doParseXml xhr.responseText
        End If
    End If
End Sub

'xhr.onreadystatechange = GetRef("doReadyStateChange")
'xhr.Open hMethod, hUrl, True
'
'If hAccepts <> "" Then
'    xhr.setRequestHeader "Accept", hAccepts
'End If
'
'If hMethod = "POST" Or hMethod = "PUT" Then
'    'Accpet HTTP request
'    If hFormat = "" Then
'        xhr.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'    Else
'        xhr.setRequestHeader "Content-Type", hFormat
'    End If
'
'    xhr.setRequestHeader "Content-Length", Len(hContents)
'    xhr.send hContents
'Else
'    xhr.send
'End If


