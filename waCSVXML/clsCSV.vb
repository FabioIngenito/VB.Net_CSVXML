Option Strict On
Option Explicit On

Imports System.IO

Public Class clsCSV

    ''' <summary>
    ''' Grava um arquivo na pasta de trabalho com a extensão ".csv"
    ''' Writes a file to the workbook with the extension ".csv"
    ''' </summary>
    ''' <param name="clsDados">Classe de variáveis (clsVariaveis).</param>
    ''' <remarks>Pode ser aberto pelo excel diretamente. - It can be opened by Excel directly.</remarks>
    Public Sub GravaSaidaCSV(ByVal clsDados As clsVariaveis)
        Dim strNomeArqS As String
        Dim intConta As Integer

        'Pega o caminho e nome do arquivo de saída
        'Take the path and file name of the output
        strNomeArqS = My.Computer.FileSystem.CurrentDirectory & "\Arquivo.CSV"

        'Cria uma instância do 'StreamWriter' para escrever texto no arquivo.
        'Creates an instance of 'StreamWriter' to write text to file.
        Using sw As StreamWriter = New StreamWriter(strNomeArqS)

            'Preenche os dados do cabeçalho
            'Fills header data
            For intConta = 0 To clsDados.Cabecalho.Count - 1
                sw.Write(clsDados.Cabecalho.Item(intConta).ToString)

                If Not intConta = clsDados.Cabecalho.Count - 1 Then sw.Write(";") Else sw.Write(Constants.vbCrLf)
            Next

            'Preenche os dados dos campos
            'Fill in field data
            For intConta = 0 To clsDados.Dados.Count - 1
                sw.Write(clsDados.Dados.Item(intConta).ToString())

                If Not intConta Mod clsDados.Cabecalho.Count - 1 = 0 Then sw.Write(";") Else sw.Write(Constants.vbCrLf)

            Next

            sw.Close()
        End Using

    End Sub

End Class
