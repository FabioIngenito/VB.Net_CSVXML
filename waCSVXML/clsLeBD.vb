Imports System.Data
Imports System.Data.oledb

Public Class clsLeBD

    ''' <summary>
    ''' Monta uma query e dispara, gravando na classe "clsVariaveis".
    ''' It assembles a query and shoots, recording in the "class variables".
    ''' </summary>
    ''' <param name="strSQL">Recebe a query passada. Receives the last query.</param>
    ''' <returns>Retorna a classe inteira montada. - Returns the entire class mounted.</returns>
    ''' <remarks></remarks>
    Public Function DisparaSQL(ByVal strSQL As String) As clsVariaveis
        Const strQtdRegistros As String = " TOP 100 "
        Dim clsVAR As New clsVariaveis
        Dim intSelectPosicao As Integer
        Dim SchemaTable As DataTable
        Dim rsBDAccess As System.Data.IDataReader
        Dim intInicio As Integer
        Dim intFim As Integer

        'Limpar variáveis. - Clear variables.
        intSelectPosicao = 0
        clsVAR.SQL = strSQL

        'Se tiver TOP na Text da Query, não considerar!
        'If you have TOP in the Text of the Query, do not consider it!
        If Strings.InStr(Strings.UCase(clsVAR.SQL), " TOP ") = 0 Then
            'Colocar o TOP
            intSelectPosicao = Strings.InStr(Strings.UCase(clsVAR.SQL), "SELECT") + 5
            clsVAR.SQL = Strings.Left(clsVAR.SQL, intSelectPosicao) & strQtdRegistros & _
                     Strings.Right(clsVAR.SQL, Strings.Len(clsVAR.SQL) - intSelectPosicao)
        End If

        'Captura a posição do nome da Tabela.
        'Captures the position of the Table name
        intInicio = clsVAR.SQL.IndexOf("FROM ") + 6
        intFim = Strings.InStr(intInicio, clsVAR.SQL, " ")
        clsVAR.NomeTabela = Strings.Mid(clsVAR.SQL, intInicio, intFim - intInicio)

        Try
            'Abre conexão com Access. - Opens connection with Access.
            SQLAccessHelper.AbrirConexao()

            'Pega o nome dos campos da tabela - Gets the name of the table fields.
            SchemaTable = SQLAccessHelper.GetOleDbSchemaTable(clsVAR.NomeTabela)

            For Each dr_field As DataRow In SchemaTable.Rows
                clsVAR.Cabecalho.Add(dr_field("COLUMN_NAME").ToString)
            Next

            'Executa a Query para pegar os dados da tabela.
            'Executes the Query to get the data from the table.
            rsBDAccess = SQLAccessHelper.ExecuteReader(clsVAR.SQL)

            'Chama a leitura para acessar os dados.
            'Calls the read to access the data.
            While rsBDAccess.Read()

                For i As Integer = 0 To rsBDAccess.FieldCount - 1
                    clsVAR.Dados.Add(rsBDAccess.Item(i).ToString)
                Next

            End While

            'Fecha o "IDataReader" depois de ler tudo.
            'Close the "IDataReader" after reading everything.
            rsBDAccess.Close()

        Catch ex As Exception
            Console.WriteLine("Erro:" & ex.Message & vbCrLf & ex.ToString)
        Finally
            SQLAccessHelper.FecharConexao()
        End Try

        Return clsVAR
    End Function

End Class
