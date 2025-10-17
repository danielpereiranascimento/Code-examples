*==================================================================
* Import Excel to FN Table
*==================================================================
* Importa dados do Excel, mapeando IVA e matrículas, atualizando tabela FN
*==================================================================

Local lcFicheiro, oExcel, oWorkbook, lcFolhas, lcSheet, lcFirstRow, lcLastRow

Fecha("mErrosIMP")
Create Cursor mErrosIMP (cErro c(200))
Select mErrosIMP
Append Blank

Try
    * Escolher o Ficheiro Excel *
    lcFicheiro = Getfile("Ficheiros Excel:XLS,XLSX,CSV,XLSM", "Nome do ficheiro", "Ok", 0, "Escolha o ficheiro Excel a importar")
    If Empty(lcFicheiro) Or !File(lcFicheiro)
        Messagebox("Nenhum arquivo selecionado ou arquivo não encontrado.", 16, "Erro")
        Return
    Endif

    * Criar objeto Excel *
    oExcel = Createobject("Excel.Application")
    oExcel.DisplayAlerts = .F.
    If Vartype(oExcel) != "O"
        Messagebox("Não foi possível abrir a aplicação Microsoft Excel.", 16, "Erro")
        Return
    Endif

    * Abrir Ficheiro *
    oWorkbook = oExcel.Application.Workbooks.Open(lcFicheiro)
    Messagebox("Arquivo Excel aberto com sucesso.")

    * Escolher a folha a importar *
    lcFolhas = oWorkbook.Sheets.Count
    Declare a_folhas(1)
    a_folhas = ""
    For i = 1 To lcFolhas
        If oExcel.Worksheets(i).UsedRange.Rows.Count >= 1
            Aadd("a_folhas", Astr(i))
        Endif
    Endfor

    * ... restante código de importação, mapeamento IVA, atualização FN ... *

    * Fechar Excel *
    oWorkbook.Close()
    oExcel.Quit
    Release oWorkbook, oExcel

Catch To oErro
    Regua(2)
    oWorkbook.Close()
    oExcel.Quit
    Fecha("mErrosIMP")
    Release("a_folhas")
    Messagebox("Erro: " + Alltrim(oErro.Message), 16, "Erro na Importação")
Endtry
