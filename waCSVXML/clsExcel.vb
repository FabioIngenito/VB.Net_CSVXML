Option Strict Off
Option Explicit On

Public Class clsExcel

    ''' <summary>
    ''' Exporta os dados para uma nova planilha do Excel.
    ''' Exports the data to a new Excel worksheet.
    ''' </summary>
    ''' <remarks>Esta forma de fazer usando "CreateObject" NÃO é boa, pois usa um "late binding", 
    '''   sendo impossível usar "Option Strict On", pois o "Option Strict On" exige "early binding".
    ''' This way of doing using "CreateObject" is NOT good, because it uses a "late binding", 
    '''   being impossible to use "Option Strict On", since "Option Strict On" requires "early binding".
    ''' O "Option Strict On" também exige que todas as conversões de dados seja explícitas no código e
    '''   isto é muito bom para evitar erros em "run-time" (em tempo de execução ou "na cara do usuário") 
    ''' Option Strict On also requires all data conversions to be explicit in code and 
    '''   this is very good for avoiding run-time errors (at runtime or "on the user's face").
    ''' Também não se tem direito ao "Intelisense"...
    ''' You also do not have the right to "Intelisense" ...
    ''' Se alguem souber fazer de outra forma, por favor, me avise para podermos transformar este "Sub"
    '''   decadente em 'Amon-Rá', o "sub" eterno. Vide bibliografia: 'He-Man'.
    ''' If anyone can do otherwise, please let me know so we can transform this decaying Sub 
    '''    in 'Amon-Ra', the eternal 'sub'. See bibliography: 'He-Man'.
    ''' </remarks>
    Public Sub DisparaExcel(ByVal clsDados As clsVariaveis)
        Dim oExcel As Object
        Dim oPlanilha As Object
        Dim oFolha As Object
        Dim oWS As Object
        Dim intColuna As Integer
        Dim intLinha As Integer = 2

        Try
            'Cria uma nova planilha.
            'Creates a new spreadsheet.
            oExcel = CreateObject("excel.application")
            oPlanilha = oExcel.Workbooks.Add
            oFolha = oPlanilha.Sheets(1)

            'Definir o nome da Folha (Sheet) 
            'Set Sheet Name
            oPlanilha.Sheets(1).Name = clsDados.NomeTabela

            'Define a folha ativa.
            'Sets the active sheet.
            oWS = oExcel.ActiveWorkbook.ActiveSheet

            'Exibe o aplicativo Excel.
            'Displays the Excel application.
            oExcel.Application.Visible = True

            'Colocar o cabeçalho - Nome das colunas
            'Placing the Header - Column Name
            For intColuna = 1 To clsDados.Cabecalho.Count
                oWS.Cells(1, intColuna) = clsDados.Cabecalho.Item(intColuna - 1)
                oWS.Cells(1, intColuna).Borders.LineStyle = 1
                oWS.Cells(1, intColuna).Interior.ThemeColor = 1                  'xlThemeColorDark1
                oWS.Cells(1, intColuna).Interior.TintAndShade = -0.25
            Next

            intColuna = 1

            'Colocar as linhas - Valores das colunas.
            'Place Rows - Column Values.
            For i As Integer = 1 To clsDados.Dados.Count
                oWS.Cells(intLinha, intColuna).NumberFormat = "@"
                oWS.Cells(intLinha, intColuna).Value = clsDados.Dados.Item(i - 1).ToString()
                oWS.Cells(intLinha, intColuna).Borders.LineStyle = 1             'xlContinuous

                If Not intLinha Mod clsDados.Cabecalho.Count - 1 = 0 Then
                    oWS.Cells(intLinha, intColuna).Interior.ThemeColor = 1       'xlThemeColorDark1
                    oWS.Cells(intLinha, intColuna).Interior.TintAndShade = -0.05
                End If

                If intColuna >= clsDados.Cabecalho.Count Then
                    intLinha += 1           'Passa para próxima linha. - Move to the next line.
                    intColuna = 1           'Coloca na primeira coluna. - Places in the first column.
                Else
                    intColuna += 1          'Adiciona uma coluna. - Adds a column.
                End If

            Next

            'Ajusta a largura de todas as colunas. - Adjusts the width of all columns.
            oWS.Columns("A:AZ").AutoFit()

            'Retire este trecho para ver a planilha aberta:
            'Remove this passage to see the worksheet open:
            oPlanilha.SaveAs(My.Computer.FileSystem.CurrentDirectory & "\Arquivo.XLSX")
            oPlanilha.Saved = True
            oExcel.UserControl = False
            oExcel.Quit()

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

End Class