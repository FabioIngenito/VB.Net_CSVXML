Option Strict On
Option Explicit On

Imports System
Imports System.Configuration
Imports System.Data
Imports System.Data.OleDb

''' <summary>
''' Classe para acessar o Access.
''' Class to access Access.
''' </summary>
''' <remarks>23/03/2011 - Fabio I.</remarks>
Public Class SQLAccessHelper

    Public Enum enumConexao
        Producao = 0
        Desenvolvimento = 1
    End Enum

    ' Time out em minutos (in minutes) COMMAND_TIMEOUT/60. 
    Private Const COMMAND_TIMEOUT As Integer = 600

    Private Shared _Connection As New OleDb.OleDbConnection
    Private Shared _ConnectionSchema As New OleDb.OleDbSchemaGuid

    ''' <summary> 
    ''' Construtor declarado <code>private</code> para prever instaciar. 
    ''' Constructor declared <code>private</code> to provide instantiate.
    ''' </summary> 
    Private Sub New()
    End Sub

    Public Shared Sub BeginTransaction()
        Dim oConnection As OleDb.OleDbConnection = _Connection
        If oConnection.State = ConnectionState.Closed Then
            oConnection.Open()
        End If
        '_Transaction = oConnection.BeginTransaction(System.Data.IsolationLevel.ReadUncommitted)
    End Sub

    Public Shared Sub Commit()
        '_Transaction.Commit()
    End Sub

    Public Shared Sub Rollback()
        '_Transaction.Rollback()
    End Sub
    ''''''''''''''''''''''''''

    ''' <summary> 
    ''' Retorna a String de Conexão.
    ''' Returns the Connection String.
    ''' </summary> 
    ''' <returns>A conexo configurada. - The connection is configured.</returns> 
    Public Shared Function GetConnectionStringProd() As String
        'Return ConfigurationManager.ConnectionStrings("Producao").ConnectionString.ToString()
        Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\DBCSVXML.accdb;Jet OLEDB:Database Password=;"
    End Function

    Public Shared Function GetConnectionStringDesev() As String
        'Return ConfigurationManager.ConnectionStrings("Desenvolvimento").ConnectionString.ToString()
        Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\DBCSVXML.accdb;Jet OLEDB:Database Password=;"
    End Function

    ''' <summary> 
    ''' Cria e Retorna uma conexao fechada para o Database configurado. 
    ''' An open connection to the configured database.
    ''' </summary> 
    Public Shared Sub AbrirConexao()

        Try

            If Not _Connection Is Nothing Then

                If _Connection.State = ConnectionState.Closed Or _Connection.State = ConnectionState.Broken Then
                    _Connection.ConnectionString = GetConnectionStringDesev()
                    _Connection.Open()
                End If

            Else
                GetConnection()
            End If

        Catch ex As Exception
            Throw ex
            'Caso dê o seguinte erro: "O provedor 'Microsoft.ACE.OLEDB.12.0' não está registrado na máquina local." Execute os seguintes passos:
            'If you give the following error: "The provider 'Microsoft.ACE.OLEDB.12.0' is not registered on the local machine." Perform the following steps:
            'http://support.microsoft.com/kb/942977/pt-br
            '1. No Solution Explorer, clique o aplicativo com o botão direito do mouse e, em seguida, clique em Propriedades.
            '1. In Solution Explorer, right-click the application, and then click Properties.
            '2. Clique na guia Compile.
            '2. Click the Compile tab.
            '3. Na guia Compile, clique em Advanced Compile Options.
            '3. On the Compile tab, click Advanced Compile Options.
            '4. Na caixa de diálogo Advanced Compiler Settings , clique em x86 na lista Target CPU e, em seguida, clique em OK.
            '4. In the Advanced Compiler Settings dialog box, click x86 in the Target CPU list, and then click.
            '5. No menu arquivo, clique em Salvar itens selecionados.
            '5. On the File menu, click Save Selected Items.
            '- Isto depende: Do SO ser 32byts ou 64byts do Banco de Dados ser 32bytes ou 64bytes
            '- This depends on: OS being 32byts or 64byts of the Database being 32bytes or 64bytes
        End Try

    End Sub

    Public Shared Sub FecharConexao()
        If Not _Connection Is Nothing Then
            If Not _Connection.State = ConnectionState.Closed Then
                _Connection.Close()
                _Connection.Dispose()
                _Connection = Nothing
            End If
        End If
    End Sub

    Public Shared Function GetConnection() As OleDb.OleDbConnection
        If _Connection Is Nothing Then
            _Connection = New OleDb.OleDbConnection(GetConnectionStringDesev())
        End If

        AbrirConexao()

        Return _Connection
    End Function

    ''' <summary>
    ''' Pega os nomes das tabelas do sistema.
    ''' Get the names of the system tables.
    ''' </summary>
    ''' <returns>A própria DataTable. - The DataTable itself.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetOleDbSchemaTable() As DataTable
        GetOleDbSchemaTable = _Connection.GetOleDbSchemaTable( _
                        System.Data.OleDb.OleDbSchemaGuid.Tables, _
                        New Object() {Nothing, Nothing, Nothing, Nothing})

        Return GetOleDbSchemaTable
    End Function

    ''' <summary>
    ''' Pega os campos de uma determinada tabela.
    ''' Gets the fields of a given table.
    ''' </summary>
    ''' <param name="strTabela">Nome da Tabela - Table Name</param>
    ''' <returns>Os campos desta tabela - The fields in this table</returns>
    ''' <remarks></remarks>
    Public Shared Function GetOleDbSchemaTable(ByVal strTabela As String) As System.Data.DataTable
        GetOleDbSchemaTable = _Connection.GetOleDbSchemaTable( _
                        OleDb.OleDbSchemaGuid.Columns, _
                        New Object() {Nothing, Nothing, strTabela})

        Return GetOleDbSchemaTable
    End Function

#Region "ExecuteNonQuery functions"
    ''' <summary> 
    ''' Executa a Stored Procedure. 
    ''' Executes the Stored Procedure.
    ''' </summary> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    Public Shared Sub ExecuteNonQuery(ByVal commandText As String)
        Dim connection As OleDb.OleDbConnection = GetConnection()

        ExecuteNonQuery(connection, commandText)
    End Sub

    ''' <summary> 
    ''' Executes the stored procedure on the specified <see cref="OleDb.OleDbConnection"/> within the specified see cref="SqlTransaction"/>. 
    ''' </summary> 
    ''' <param name="connection">The database connection to be used.</param> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    Public Shared Sub ExecuteNonQuery(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String)
        Dim command As OleDb.OleDbCommand = CreateCommand(connection, commandText)

        If connection.State = ConnectionState.Closed Then connection.Open()

        command.Connection = connection
        command.CommandType = CommandType.Text
        command.CommandText = commandText.ToString

        command.ExecuteNonQuery()
    End Sub

    '''' <summary> 
    '''' Executes the stored procedure with the specified parameters on the specified connection. 
    '''' </summary> 
    '''' <param name="connection">The database connection to be used.</param> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <param name="parameters">The parameters of the stored procedure.</param> 
    'Public Shared Sub ExecuteNonQuery(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String, ByVal parameters As ArrayList)
    '    ExecuteNonQuery(connection, Nothing, commandText, parameters)
    'End Sub

    '''' <summary> 
    '''' Executes the stored procedure with the specified parameters on the specified <see cref="OleDb.OleDbConnection"/> within the specified <see cref="SqlTransaction"/>. 
    '''' </summary> 
    '''' <param name="connection">The database connection to be used.</param> 
    '''' <param name="transaction">The transaction to participate in.</param> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <param name="parameters">The parameters of the stored procedure.</param> 
    'Public Shared Sub ExecuteNonQuery(ByVal connection As OleDb.OleDbConnection, ByVal transaction As SqlTransaction, ByVal commandText As String, ByVal parameters As ArrayList)
    '    If connection.State = ConnectionState.Closed Then
    '        connection.Open()
    '    End If
    '    Dim command As OleDb.OleDbConnection = CreateCommand(connection, transaction, commandText, parameters)

    '    command.ExecuteNonQuery()
    'End Sub

    'Public Shared Function ExecuteNonQuery(ByVal commandText As String, ByVal parameters As ArrayList, ByVal ParametroRetorno As String) As Int64
    '    Dim vRetorno As Int64
    '    Dim connection As OleDb.OleDbConnection = GetConnection()
    '    connection.Open()
    '    Dim transaction As SqlTransaction = Nothing
    '    parameters.Add(New SqlParameter(ParametroRetorno, SqlDbType.BigInt, 8, ParameterDirection.Output, False, 19, _
    '    0, "IDRetorno", DataRowVersion.Proposed, Nothing))
    '    Dim command As OleDb.OleDbConnection = CreateCommand(connection, transaction, commandText, parameters)
    '    command.ExecuteNonQuery()
    '    vRetorno = Convert.ToInt64(command.Parameters(ParametroRetorno).Value)
    '    Return vRetorno
    'End Function

    'Public Shared Function ExecuteNonQuery(ByVal commandText As String, ByVal transaction As SqlTransaction, ByVal parameters As ArrayList, ByVal ParametroRetorno As String) As Int64
    '    Dim vRetorno As Int64
    '    Dim connection As OleDb.OleDbConnection
    '    If transaction Is Nothing Then
    '        connection = GetConnection()
    '        connection.Open()
    '    Else
    '        connection = transaction.Connection
    '    End If

    '    parameters.Add(New SqlParameter(ParametroRetorno, SqlDbType.BigInt, 8, ParameterDirection.Output, False, 19, _
    '    0, "IDRetorno", DataRowVersion.Proposed, Nothing))
    '    Dim command As OleDb.OleDbConnection = CreateCommand(connection, transaction, commandText, parameters)
    '    command.ExecuteNonQuery()
    '    vRetorno = Convert.ToInt64(command.Parameters(ParametroRetorno).Value)
    '    Return vRetorno
    'End Function

#End Region

#Region "ExecuteReader functions"
    ''' <summary> 
    ''' Executes the stored procedure and returns the result as a see cref="SqlDataReader"/>. 
    ''' </summary> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    ''' <returns>A see cref="SqlDataReader"/> containing the results of the stored procedure execution.</returns> 
    Public Shared Function ExecuteReader(ByVal commandText As String) As OleDb.OleDbDataReader
        Dim connection As OleDb.OleDbConnection = GetConnection()
        Return ExecuteReader(connection, commandText)
    End Function

    '''' <summary> 
    '''' Executes the stored procedure on the specified <see cref="OleDb.OleDbConnection"/> and returns the result as a see cref="SqlDataReader"/>. 
    '''' </summary> 
    '''' <param name="connection">The database connection to be used.</param> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <returns>A see cref="SqlDataReader"/> containing the results of the stored procedure execution.</returns> 
    'Public Shared Function ExecuteReader(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String) As OleDb.OleDbDataReader
    '    If connection.State = ConnectionState.Closed Then
    '        connection.Open()
    '    End If
    '    Return ExecuteReader(connection, commandText)
    'End Function

    ''' <summary> 
    ''' Executes the stored procedure on the specified see cref="OleDb.OleDbConnection"/> within the specified see cref="SqlTransaction"/> and returns the result as a see cref="SqlDataReader"/>. 
    ''' </summary> 
    ''' <param name="connection">The database connection to be used.</param> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    ''' <returns>A see cref="SqlDataReader"/> containing the results of the stored procedure execution.</returns> 
    Public Shared Function ExecuteReader(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String) As OleDb.OleDbDataReader
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        Dim command As OleDb.OleDbCommand = CreateCommand(connection, commandText)
        'command.CommandType = CommandType.Text
        Return command.ExecuteReader()
        'connection.Close()
    End Function

    ''' <summary> 
    ''' Executes the stored procedure with the specified parameters and returns the result as a see cref="SqlDataReader"/>. 
    ''' </summary> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    ''' <param name="parameters">The parameters of the stored procedure.</param> 
    ''' <returns>A see cref="SqlDataReader" containing the results of the stored procedure execution.</returns> 
    Public Shared Function ExecuteReader(ByVal commandText As String, ByVal parameters As ArrayList) As OleDb.OleDbDataReader
        Dim connection As OleDb.OleDbConnection = GetConnection()
        connection.Open()
        Return ExecuteReader(commandText)
    End Function

    ''' <summary> 
    ''' Executes the stored procedure on the specified see cref="OleDb.OleDbConnection"/> within the specified see cref="SqlTransaction"/> with the specified parameters and returns the result as a see cref="SqlDataReader"/>. 
    ''' </summary> 
    ''' <param name="connection">The database connection to be used.</param> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    ''' <param name="parameters">The parameters of the stored procedure.</param> 
    ''' <returns>A see cref="SqlDataReader"/> containing the results of the stored procedure execution.</returns> 
    Public Shared Function ExecuteReader(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String, ByVal parameters As ArrayList) As OleDb.OleDbDataReader
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        Dim command As OleDb.OleDbCommand = CreateCommand(connection, commandText, parameters)

        Return command.ExecuteReader()
    End Function

#End Region

#Region "ExecuteScalar functions"
    ''' <summary> 
    ''' Executes the stored procedure, and returns the first column of the first row in the result set returned by the query. Extra columns or rows are ignored. 
    ''' </summary> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    ''' <returns>The first column of the first row in the result set, or a null reference if the result set is empty.</returns> 
    Public Shared Function ExecuteScalar(ByVal commandText As String) As Object

        Dim connection As OleDb.OleDbConnection
        connection = GetConnection()
        Using connection
            Return ExecuteScalar(connection, commandText)
        End Using

    End Function

    ''' <summary> 
    ''' Executes the stored procedure on the specified see cref="OleDb.OleDbConnection"/> within the specified see cref="SqlTransaction"/>, and returns the first column of the first row in the result set returned by the query. Extra columns or rows are ignored. 
    ''' </summary> 
    ''' <param name="connection">The database connection to be used.</param> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    ''' <returns>The first column of the first row in the result set, or a null reference if the result set is empty.</returns> 
    Public Shared Function ExecuteScalar(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String) As Object

        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If

        Using command As OleDb.OleDbCommand = CreateCommand(connection, commandText)
            command.CommandType = CommandType.Text
            command.CommandText = commandText.ToString
            Return command.ExecuteScalar
        End Using

    End Function

    '''' <summary> 
    '''' Executes the stored procedure with the specified parameters, and returns the first column of the first row in the result set returned by the query. Extra columns or rows are ignored. 
    '''' </summary> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <param name="parameters">The parameters of the stored procedure.</param> 
    '''' <returns>The first column of the first row in the result set, or a null reference if the result set is empty.</returns> 
    'Public Shared Function ExecuteScalar(ByVal commandText As String, ByVal parameters As ArrayList) As Object
    '    Using connection As OleDb.OleDbConnection = GetConnection()
    '        Return ExecuteScalar(connection, commandText, parameters)
    '    End Using
    'End Function

    '''' <summary> 
    '''' Executes the stored procedure on the specified see cref="SqlTransaction"/> with the specified parameters, and returns the first column of the first row in the result set returned by the query. Extra columns or rows are ignored. 
    '''' </summary> 
    '''' <param name="connection">The database connection to be used.</param> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <param name="parameters">The parameters of the stored procedure.</param> 
    '''' <returns>The first column of the first row in the result set, or a null reference if the result set is empty.</returns> 
    'Public Shared Function ExecuteScalar(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String, ByVal parameters As ArrayList) As Object
    '    Return ExecuteScalar(connection, Nothing, commandText, parameters)
    'End Function

    '''' <summary> 
    '''' Executes the stored procedure on the specified see cref="SqlTransaction"/> within the specified <see cref="SqlTransaction"/> with the specified parameters, and returns the first column of the first row in the result set returned by the query. Extra columns or rows are ignored. 
    '''' </summary> 
    '''' <param name="connection">The database connection to be used.</param> 
    '''' <param name="transaction">The transaction to participate in.</param> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <param name="parameters">The parameters of the stored procedure.</param> 
    '''' <returns>The first column of the first row in the result set, or a null reference if the result set is empty.</returns> 
    'Public Shared Function ExecuteScalar(ByVal connection As OleDb.OleDbConnection, ByVal transaction As SqlTransaction, ByVal commandText As String, ByVal parameters As ArrayList) As Object
    '    If connection.State = ConnectionState.Closed Then
    '        connection.Open()
    '    End If
    '    Using command As OleDb.OleDbConnection = CreateCommand(connection, transaction, commandText, parameters)
    '        Return command.ExecuteScalar()
    '    End Using
    'End Function
#End Region

#Region "ExecuteDataSet functions"
    '''' <summary> 
    '''' Executes the stored procedure and returns the result as a <see cref="DataSet"/>. 
    '''' </summary> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <returns>A <see cref="DataSet"/> containing the results of the stored procedure execution.</returns> 
    'Public Shared Function ExecuteDataSet(ByVal commandText As String) As DataSet
    '    Using connection As OleDb.OleDbConnection = GetConnection()
    '        Using command As OleDb.OleDbConnection = CreateCommand(connection, commandText)
    '            Return CreateDataSet(command)
    '        End Using
    '    End Using
    'End Function

    '''' <summary> 
    '''' Executes the stored procedure and returns the result as a <see cref="DataSet"/>. 
    '''' </summary> 
    '''' <param name="connection">The database connection to be used.</param> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <returns>A <see cref="DataSet"/> containing the results of the stored procedure execution.</returns> 
    'Public Shared Function ExecuteDataSet(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String) As DataSet
    '    Using command As OleDb.OleDbCommand = CreateCommand(connection, commandText)
    '        Return CreateDataSet(command)
    '    End Using
    'End Function

    ''' <summary> 
    ''' Executes the stored procedure and returns the result as a <see cref="DataSet"/>. 
    ''' </summary> 
    ''' <param name="connection">The database connection to be used.</param> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    ''' <returns>A <see cref="DataSet"/> containing the results of the stored procedure execution.</returns> 
    Public Shared Function ExecuteDataSet(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String) As DataSet
        Using command As OleDb.OleDbCommand = CreateCommand(connection, commandText)
            Return CreateDataSet(command)
        End Using
    End Function

    ''' <summary> 
    ''' Executes the stored procedure and returns the result as a <see cref="DataSet"/>. 
    ''' </summary> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    ''' <param name="parameters">The parameters of the stored procedure.</param> 
    ''' <returns>A <see cref="DataSet"/> containing the results of the stored procedure execution.</returns> 
    Public Shared Function ExecuteDataSet(ByVal commandText As String, ByVal parameters As ArrayList) As DataSet
        Using connection As OleDb.OleDbConnection = GetConnection()
            Using command As OleDb.OleDbCommand = CreateCommand(connection, commandText, parameters)
                Return CreateDataSet(command)
            End Using
        End Using
    End Function

    '''' <summary> 
    '''' Executes the stored procedure and returns the result as a <see cref="DataSet"/>. 
    '''' </summary> 
    '''' <param name="connection">The database connection to be used.</param> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <param name="parameters">The parameters of the stored procedure.</param> 
    '''' <returns>A <see cref="DataSet"/> containing the results of the stored procedure execution.</returns> 
    'Public Shared Function ExecuteDataSet(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String, ByVal parameters As ArrayList) As DataSet
    '    Using command As OleDb.OleDbConnection = CreateCommand(connection, commandText, parameters)
    '        Return CreateDataSet(command)
    '    End Using
    'End Function

    ''' <summary> 
    ''' Executes the stored procedure and returns the result as a <see cref="DataSet"/>. 
    ''' </summary> 
    ''' <param name="connection">The database connection to be used.</param> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    ''' <param name="parameters">The parameters of the stored procedure.</param> 
    ''' <returns>A <see cref="DataSet"/> containing the results of the stored procedure execution.</returns> 
    Public Shared Function ExecuteDataSet(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String, ByVal parameters As ArrayList) As DataSet
        Using command As OleDb.OleDbCommand = CreateCommand(connection, commandText, parameters)
            Return CreateDataSet(command)
        End Using
    End Function
#End Region

#Region "ExecuteDataTable functions"
    ''' <summary> 
    ''' Executes the stored procedure and returns the result as a <see cref="DataTable"/>. 
    ''' </summary> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    ''' <returns>A <see cref="DataTable"/> containing the results of the stored procedure execution.</returns> 
    Public Shared Function ExecuteDataTable(ByVal commandText As String) As DataTable
        Dim connection As OleDb.OleDbConnection = GetConnection()
        Dim command As OleDb.OleDbCommand = CreateCommand(connection, commandText)
        command.CommandType = CommandType.Text
        Return CreateDataTable(command)
    End Function

    '''' <summary> 
    '''' Executes the stored procedure and returns the result as a <see cref="DataTable"/>. 
    '''' </summary> 
    '''' <param name="connection">The database connection to be used.</param> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <returns>A <see cref="DataTable"/> containing the results of the stored procedure execution.</returns> 
    'Public Shared Function ExecuteDataTable(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String) As DataTable
    '    Using command As OleDb.OleDbConnection = CreateCommand(connection, commandText)
    '        Return CreateDataTable(command)
    '    End Using
    'End Function

    ''' <summary> 
    ''' Executes the stored procedure and returns the result as a <see cref="DataTable"/>. 
    ''' </summary> 
    ''' <param name="connection">The database connection to be used.</param> 
    ''' <param name="commandText">The stored procedure to execute.</param> 
    ''' <returns>A <see cref="DataTable"/> containing the results of the stored procedure execution.</returns> 
    Public Shared Function ExecuteDataTable(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String) As DataTable
        Using command As OleDb.OleDbCommand = CreateCommand(connection, commandText)
            Return CreateDataTable(command)
        End Using
    End Function

    '''' <summary> 
    '''' Executes the stored procedure and returns the result as a <see cref="DataTable"/>. 
    '''' </summary> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <param name="parameters">The parameters of the stored procedure.</param> 
    '''' <returns>A <see cref="DataTable"/> containing the results of the stored procedure execution.</returns> 
    'Public Shared Function ExecuteDataTable(ByVal commandText As String, ByVal parameters As ArrayList) As DataTable
    '    Using connection As OleDb.OleDbConnection = GetConnection()
    '        Using command As OleDb.OleDbConnection = CreateCommand(connection, commandText, parameters)
    '            Return CreateDataTable(command)
    '        End Using
    '    End Using
    'End Function

    '''' <summary> 
    '''' Executes the stored procedure and returns the result as a <see cref="DataTable"/>. 
    '''' </summary> 
    '''' <param name="connection">The database connection to be used.</param> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <param name="parameters">The parameters of the stored procedure.</param> 
    '''' <returns>A <see cref="DataTable"/> containing the results of the stored procedure execution.</returns> 
    'Public Shared Function ExecuteDataTable(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String, ByVal parameters As ArrayList) As DataTable
    '    Using command As OleDb.OleDbConnection = CreateCommand(connection, commandText, parameters)
    '        Return CreateDataTable(command)
    '    End Using
    'End Function

    '''' <summary> 
    '''' Executes the stored procedure and returns the result as a <see cref="DataTable"/>. 
    '''' </summary> 
    '''' <param name="connection">The database connection to be used.</param> 
    '''' <param name="commandText">The stored procedure to execute.</param> 
    '''' <param name="parameters">The parameters of the stored procedure.</param> 
    '''' <returns>A see cref="DataTable"/> containing the results of the stored procedure execution.</returns> 
    'Public Shared Function ExecuteDataTable(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String, ByVal parameters As ArrayList) As DataTable
    '    Using command As OleDb.OleDbCommand = CreateCommand(connection, commandText, parameters)
    '        Return CreateDataTable(OleDb.OleDbCommand)
    '    End Using
    'End Function
#End Region

#Region "compact access data base"
    'Dim connection As OleDb.OleDbConnection = GetConnection()
#End Region

#Region "Utility functions"
    ''' <summary> 
    ''' Sets the specified see cref="SqlParameter"/>'s <code>Value</code> property to <code>DBNull.Value</code> if it is <code>null</code>. 
    ''' </summary> 
    ''' <param name="vParametro">The see cref="SqlParameter"/> that should be checked for nulls.</param> 
    ''' <returns>The see cref="SqlParameter"/> with a potentially updated <code>Value</code> property.</returns> 

    Public Shared Function AddParametro(ByVal vParametro As String, ByVal oValor As Object) As ArrayList
        vParametro = AddArroba(vParametro)
        Dim parameters As New ArrayList()
        parameters.Add(New OleDb.OleDbParameter(vParametro, oValor))
        Return parameters
    End Function

    Public Shared Sub AddParametro(ByRef parameters As ArrayList, ByVal vParametro As String, ByVal oValor As Object)
        vParametro = AddArroba(vParametro)
        parameters.Add(New OleDb.OleDbParameter(vParametro, oValor))
    End Sub

    Private Shared Function AddArroba(ByVal vParametro As String) As String
        If vParametro.Substring(0, 1) <> "@" Then
            vParametro.Insert(0, "@")
        End If
        Return vParametro
    End Function

    Private Shared Function CheckParameter(ByVal parameter As OleDb.OleDbParameter) As OleDb.OleDbParameter
        If parameter.Value Is Nothing Then
            parameter.Value = DBNull.Value
        End If
        If (parameter.DbType = DbType.[String] OrElse parameter.DbType = DbType.AnsiString OrElse parameter.DbType = DbType.AnsiStringFixedLength OrElse parameter.DbType = DbType.StringFixedLength) AndAlso parameter.Value.ToString() = "" Then
            parameter.Value = DBNull.Value
        End If
        If (parameter.DbType = DbType.DateTime OrElse parameter.DbType = DbType.[Date] OrElse parameter.DbType = DbType.Time) AndAlso Convert.ToDateTime(parameter.Value.ToString()).ToShortDateString() = Convert.ToDateTime("01/01/1900 00:00:00").ToShortDateString() Then
            parameter.Value = DBNull.Value
        End If
        If (parameter.DbType = DbType.Int16 OrElse parameter.DbType = DbType.Int32 OrElse parameter.DbType = DbType.Int64 OrElse parameter.DbType = DbType.Currency OrElse parameter.DbType = DbType.[Decimal] OrElse parameter.DbType = DbType.[Double] OrElse parameter.DbType = DbType.[Single] OrElse parameter.DbType = DbType.UInt16 OrElse parameter.DbType = DbType.UInt32 OrElse parameter.DbType = DbType.UInt64 OrElse parameter.DbType = DbType.VarNumeric OrElse parameter.DbType = DbType.[Byte]) AndAlso parameter.Value.ToString() = "-1" Then
            parameter.Value = DBNull.Value
        End If

        Return parameter
    End Function

#Region "CreateCommand"
    '''' <summary> 
    '''' Creates, initializes, and returns a <see cref="OleDb.OleDbConnection"/> instance. 
    '''' </summary> 
    '''' <param name="connection">The see cref="OleDb.OleDbConnection"/> the <see cref="OleDb.OleDbConnection"/> should be executed on.</param> 
    '''' <param name="commandText">The name of the stored procedure to execute.</param> 
    '''' <returns>An initialized <see cref="OleDb.OleDbConnection"/> instance.</returns> 
    'Private Shared Function CreateCommand(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String) As OleDb.OleDbCommand
    '    Return CreateCommand(connection, Nothing, commandText)
    'End Function

    ''' <summary> 
    ''' Creates, initializes, and returns a <see cref="OleDb.OleDbConnection"/> instance. 
    ''' </summary> 
    ''' <param name="connection">The see cref="OleDb.OleDbConnection"/> the <see cref="OleDb.OleDbConnection"/> should be executed on.</param> 
    ''' <param name="commandText">The name of the stored procedure to execute.</param> 
    ''' <returns>An initialized <see cref="OleDb.OleDbConnection"/> instance.</returns> 
    Private Shared Function CreateCommand(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String) As OleDb.OleDbCommand
        Dim command As New OleDb.OleDbCommand()
        command.Connection = connection
        command.CommandText = commandText
        command.CommandTimeout = COMMAND_TIMEOUT

        '25/03/2011 - Fabio I. - Stored Procedured somente para SQL Server
        'If commandText.Trim.ToUpper().StartsWith("SELECT") Then
        command.CommandType = CommandType.Text
        'Else
        '   command.CommandType = CommandType.StoredProcedure
        ' End If

        Return command
    End Function

    '''' <summary> 
    '''' Creates, initializes, and returns a <see cref="OleDb.OleDbConnection"/> instance. 
    '''' </summary> 
    '''' <param name="connection">The <see cref="OleDb.OleDbConnection"/> the <see cref="OleDb.OleDbConnection"/> should be executed on.</param> 
    '''' <param name="commandText">The name of the stored procedure to execute.</param> 
    '''' <param name="parameters">The parameters of the stored procedure.</param> 
    '''' <returns>An initialized <see cref="OleDb.OleDbConnection"/> instance.</returns> 
    'Private Shared Function CreateCommand(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String, ByVal parameters As ArrayList) As OleDb.OleDbCommand
    '    Return CreateCommand(connection, Nothing, commandText, parameters)
    'End Function

    ''' <summary> 
    ''' Creates, initializes, and returns a <see cref="OleDb.OleDbConnection"/> instance. 
    ''' </summary> 
    ''' <param name="connection">The see cref="OleDb.OleDbConnection"/> the <see cref="OleDb.OleDbConnection"/> should be executed on.</param> 
    ''' <param name="commandText">The name of the stored procedure to execute.</param> 
    ''' <param name="parameters">The parameters of the stored procedure.</param> 
    ''' <returns>An initialized <see cref="OleDb.OleDbConnection"/> instance.</returns> 
    Private Shared Function CreateCommand(ByVal connection As OleDb.OleDbConnection, ByVal commandText As String, ByVal parameters As ArrayList) As OleDb.OleDbCommand
        Dim command As New OleDb.OleDbCommand()
        command.Connection = connection
        command.CommandText = commandText
        command.CommandTimeout = COMMAND_TIMEOUT
        command.CommandType = CommandType.StoredProcedure
        'command.Transaction = transaction

        ' Append each parameter to the command 
        For Each parameter As OleDb.OleDbParameter In parameters
            command.Parameters.Add(CheckParameter(parameter))
        Next

        Return command
    End Function
#End Region

    Private Shared Function CreateDataSet(ByVal command As OleDb.OleDbCommand) As DataSet
        Using dataAdapter As New OleDbDataAdapter(command)
            Dim dataSet As New DataSet()
            dataAdapter.Fill(dataSet)
            Return dataSet
        End Using
    End Function

    Private Shared Function CreateDataTable(ByVal command As OleDb.OleDbCommand) As DataTable
        Using dataAdapter As New OleDbDataAdapter(command)
            Dim dataTable As New DataTable()
            dataAdapter.Fill(dataTable)
            Return dataTable
        End Using
    End Function
#End Region

End Class