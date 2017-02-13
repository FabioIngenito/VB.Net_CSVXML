Option Strict On
Option Explicit On

Public Class EntryPoint
    Shared clsVAR As New clsVariaveis
    Shared strTexto(1) As String

    Shared Sub Main()
        Dim bytConta As Byte = 0
        Dim clsLBD As New clsLeBD
        Dim clsCS As New clsCSV
        Dim clsXM As New clsXML
        Dim clsEX As New clsExcel
        Dim clsDadosFilho As New clsVariaveis

        'Exemplo: O parâmetro é passado por linha de comando.
        'Example: The parameter is passed by command line.
        Console.WriteLine("Exemplo: O parâmetro é passado por linha de comando.")
        For Each argument As String In My.Application.CommandLineArgs
            strTexto(bytConta) = argument
            bytConta += CType(1, Byte)
        Next

        'Atenção: Apagar TODOS os Console.WriteLine e Console.ReadLine
        'Warning: Delete ALL Console.WriteLine and Console.ReadLine

        'preencher variáveis - Fill variables
        Console.WriteLine("preencher variáveis")
        clsVAR.Conexao = strTexto(0)
        Console.WriteLine("strTexto(0): " & strTexto(0))
        clsVAR.SQL = strTexto(1)
        Console.WriteLine("strTexto(1): " & strTexto(1))
        Console.WriteLine("preencher variáveis")
        'Ler banco e dados e preencher variáveis.
        'Read bank and data and populate variables.
        Console.WriteLine("Ler banco e dados e preencher variáveis")
        clsDadosFilho = clsLBD.DisparaSQL(clsVAR.SQL)

        'Gravar arquivo de saída CSV. - Write CSV output file.
        Console.WriteLine("Gravar arquivo de saída CSV")
        clsCS.GravaSaidaCSV(clsDadosFilho)
        'Gravar arquivo de saída XML - Write XML output file.
        Console.WriteLine("Gravar arquivo de saída XML")
        clsXM.GravaSaidaXML(clsDadosFilho)
        'Cria arquivo abrindo o Excel (mostrando ou não o excel aberto)
        'Create file by opening Excel (showing whether or not Excel is open).
        Console.WriteLine("Cria arquivo abrindo o Excel (mostrando ou não o excel aberto)")
        clsEX.DisparaExcel(clsDadosFilho)
        Console.WriteLine("FIM - Abra a pasta de '''..\bin\Debug\''' para ver os resultados.")
        Console.ReadLine()
    End Sub

End Class