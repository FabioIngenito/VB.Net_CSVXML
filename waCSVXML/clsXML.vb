Option Strict On
Option Explicit On

Imports System.Xml

Public Class clsXML

    ''' <summary>
    ''' Grava um arquivo na pasta de trabalho com a extensão ".xml"
    ''' Writes a file to the workbook with the extension ".xml".
    ''' </summary>
    ''' <param name="clsDados">Classe de variáveis (clsVariaveis). - Variable class (variable variables).</param>
    ''' <remarks>Pode ser aberto pelo excel quase que direto. - It can be opened by Excel almost straightforward.</remarks>
    Public Sub GravaSaidaXML(ByVal clsDados As clsVariaveis)
        Dim strNomeArqE As String
        Dim writer As XmlTextWriter
        Dim intContaCabecalho As Integer = 0

        'Pega o caminho e nome do arquivo de entrada.
        ''Get the path and filename of the input file.
        strNomeArqE = My.Computer.FileSystem.CurrentDirectory & "\Arquivo.XML"

        writer = New XmlTextWriter(strNomeArqE, System.Text.UTF8Encoding.UTF8) 'encoding="UTF-8"

        'Inicia o documento xml - Start the xml document
        writer.WriteStartDocument(True) 'standalone="yes"
        'Define a indentação do arquivo - Sets the file indentation
        writer.Formatting = Formatting.Indented
        'Escreve um comentario - Write a comment
        writer.WriteComment(strNomeArqE)
        'Escreve o elemento raiz - Writes the root element
        writer.WriteStartElement("ARQUIVO")

        'Preenche os dados dos campos - Fill in field data
        For intContaDados As Integer = 0 To clsDados.Dados.Count - 1

            If intContaCabecalho < clsDados.Cabecalho.Count - 1 Then writer.WriteStartElement("REGISTRO")

            'Preenche os dados do cabeçalho - Fills header data
            writer.WriteStartElement(clsDados.Cabecalho.Item(intContaCabecalho).ToString)
            writer.WriteRaw(clsDados.Dados.Item(intContaDados).ToString)
            writer.WriteEndElement()

            If intContaCabecalho >= clsDados.Cabecalho.Count - 1 Then
                intContaCabecalho = 0
                'Encerra os elementos itens. - Closes items items.
                writer.WriteEndElement()
            Else
                intContaCabecalho += 1
            End If

        Next

        'Encerra o elemento raiz. - Closes the root element.
        writer.WriteFullEndElement()
        'Escreve o XML para o arquivo e fecha o escritor. - Writes the XML to the file and closes the writer.
        writer.Close()
    End Sub

End Class
