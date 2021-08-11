VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DadosEtiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CodProduto As Long
Dim QtdeItem As Long
Dim PesoProduto As Long
Dim Valido As Boolean
Dim PesoBalanca As Long
Dim CodProdutoFinal As Long
Dim dataProducao As Date

Public Function Consiste_Codigos() As Boolean
    
    Consiste_Codigos = True
    
    Set Produto = New Produtos
    If Not (Produto.busca_produto(CodProduto)) Then
        Consiste_Codigos = False      'Não achou o produto
        MsgErroEtiqueta = "Produto não encontrado"
        
    End If

End Function


Public Function Consiste_Pesos() As Boolean
    Consiste_Pesos = False
    
    If PesoProduto = PesoBalanca Then
        Consiste_Pesos = True
    Else
        If PesoProduto > PesoBalanca Then
            If ((PesoProduto - PesoBalanca) <= 20) Then
                Consiste_Pesos = True
            End If
        Else
            If ((PesoBalanca - PesoProduto) <= 20) Then
                Consiste_Pesos = True
            End If
        End If
    End If
    
    If Not Consiste_Pesos Then
        MsgErroEtiqueta = "Produtos com pesos diferentes"
    Else
        setDataProducao
    End If
End Function

Public Function Consiste_Dados(dados As String) As Boolean
    'Faz a consistência dos dados lidos na pré-Etiqueta
    Carrega_Dados (dados)
    
    Consiste_Dados = True
    
    If CodProduto = "" Then
        Consiste_Dados = False
    Else
        Set Produto = New Produtos
        If Not (Produto.busca_produto(CodProduto)) Then
            Consiste_Dados = False      'Não achou o produto
            MsgErroEtiqueta = "Produto não encontrado"
        End If
    End If
    
    If QtdeItem = "" Then
        Consiste_Dados = False
    End If
    
    If PesoProduto = "" Then
        Consiste_Dados = False
    End If
    
    If Consiste_Dados Then
        Valido = True
    End If
End Function

Private Function Carrega_Dados(dados As String) As Boolean
    'Carrega os valores lidos
    CodProduto = Mid(dados, 1, 2)
    QtdeItem = Mid(dados, 1, 2)
    PesoProduto = Mid(dados, 1, 2)
    
    Carrega_Dados = True
End Function

Public Function getCodProduto() As Long
    getCodProduto = CodProduto
End Function

Public Function getPesoProduto() As Long
    getPesoProduto = PesoProduto
End Function

Public Function getQtdeItem() As Long
    getQtdeItem = QtdeItem
End Function

Public Function getValido() As Long
    getValido = Valido
End Function

Public Sub setValido(valid As Boolean)
    Valido = valid
End Sub

Public Function getPesoBalanca() As Long
    getPesoBalanca = PesoBalanca
End Function

Public Sub setPesoBalanca(peso As Long)
    PesoBalanca = peso
End Sub

Public Function getCodProdutoFinal() As Long
    getCodProdutoFinal = CodProdutoFinal
End Function

Public Sub setCodProdutoFinal(codigo As Long)
    CodProdutoFinal = codigo
End Sub

Public Function getDataProducao() As Date
    getDataProducao = dataProducao
End Function

Public Sub setDataProducao()
    dataProducao = Now
End Sub

Public Function ImprimeEtiqueta() As Boolean
    'Faz a impressão da etiqueta
End Function

Public Function ImprimeEtiquetaERRO() As Boolean
    'Faz a impressão da etiqueta
End Function


