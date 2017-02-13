Public Class clsVariaveis

#Region "Atributos"
    Private _bytConexao As Byte
    Private _strSQL As String
    Private _arrstrCabecalho As New System.Collections.ArrayList
    Private _clsDados As New System.Collections.ArrayList
    Private _NomeTabela As String
#End Region

#Region "Propriedades"

    ''' <summary>
    ''' Define em qual banco irá se conectar: Desenvolvimento ou Produção.
    ''' Defines which bank to connect to: Development or Production.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Retorna 1 para desenvolvimento e 0 para produção. - Returns 1 for development and 0 for production.</returns>
    ''' <remarks></remarks>
    Public Property Conexao() As String

        Get
            Return _bytConexao
        End Get

        Set(ByVal value As String)

            If value.ToUpper = "DESENVOLVIMENTO" Or value.ToUpper = "D" Then
                _bytConexao = 1
            ElseIf value.ToUpper = "PRODUÇÃO" Or value.ToUpper = "P" Then
                _bytConexao = 0
            End If

        End Set

    End Property

    ''' <summary>
    ''' Armazena a query.
    ''' Stores the query.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Retorna a query armazenada. - Returns the stored query.</returns>
    ''' <remarks></remarks>
    Public Property SQL() As String

        Get
            Return _strSQL
        End Get

        Set(ByVal value As String)
            _strSQL = value.ToUpper
        End Set

    End Property

    ''' <summary>
    ''' Guarda os nomes dos campos do banco de dados.
    ''' Saves the names of the database fields.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Retorna o "cabeçalho" (nome dos campos do banco dados). - Returns the "header" (name of the database fields).</returns>
    ''' <remarks></remarks>
    Public Property Cabecalho() As System.Collections.ArrayList

        Get
            Return _arrstrCabecalho
        End Get

        Set(ByVal value As System.Collections.ArrayList)
            _arrstrCabecalho = value
        End Set

    End Property

    ''' <summary>
    ''' Guarda os dados das tabelas.
    ''' Saves the data in the tables.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Retorna os dados das tabelas. - Returns the data in the tables.</returns>
    ''' <remarks></remarks>
    Public Property Dados() As System.Collections.ArrayList

        Get
            Return _clsDados
        End Get

        Set(ByVal value As System.Collections.ArrayList)
            _clsDados = value
        End Set

    End Property

    ''' <summary>
    ''' Armazena nome da Tabela.
    ''' Stores the name of the table.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Retorna o nome da tabela da query armazenada. - Returns the name of the stored query table.</returns>
    ''' <remarks></remarks>
    Public Property NomeTabela() As String

        Get
            Return _NomeTabela
        End Get

        Set(ByVal value As String)
            _NomeTabela = value.ToUpper
        End Set

    End Property

#End Region

End Class