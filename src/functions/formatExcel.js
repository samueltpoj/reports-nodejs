import ExcelJS from "exceljs";
const workbook = new ExcelJS.Workbook();
export async function formatExcel() {
  console.log("Formatando o Excel...");
  workbook.xlsx
    .readFile("src/Excel/dadosbrutos.xlsx")
    .then(() => {
      const worksheet = workbook.getWorksheet("Planilha1"); // Nome da Planilha que deseja obter
      // Titulos em negrito e centralizado
      const rowTitulo = worksheet.getRow(1);
      rowTitulo.eachCell((cell) => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: "center" };
      });

      // Indentifica as linhas que comeca com total e coloca a respectiva linha em negrito
      const columnTotal = worksheet.getColumn("B");
      columnTotal.eachCell((cell) => {
        if (typeof cell.value === "string" && cell.value.startsWith("TOTAL ")) {
          const row = cell.row;
          worksheet.getRow(row).font = { bold: true };
        }
      });
      // Coloca as colunas em formato REAL
      const columnsReal = ["D", "E", "F", "G", "K", "L", "M", "N"];
      columnsReal.forEach((column) => {
        const columnObj = worksheet.getColumn(column);
        columnObj.eachCell((cell) => {
          cell.numFmt = "R$ * #,##0.00_);(R$ *-#,##0.00)";
        });
      });
      // Define o formato de ponto para as colunas C e J
      const columnsPonto = ["C", "J"];
      columnsPonto.forEach((column) => {
        const columnObj = worksheet.getColumn(column);
        // Define as propriedades de formatação para cada célula na coluna
        columnObj.eachCell((cell) => {
          const value = cell.value;
          if (typeof value === "number" && value >= 1000) {
            cell.numFmt = "#,##";
          }
        });
      });
      worksheet.columns.forEach((column) => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
          const length = cell.value ? cell.value.toString().length : 0;
          if (length > maxLength) {
            maxLength = length;
          }
        });
        column.width = maxLength + 5;
      });
      // Salva o arquivo
      return workbook.xlsx.writeFile(`src/Excel/LUC.xlsx`);
    })
    .catch((error) => {
      console.log("Ocorreu um erro ao carregar o arquivo", error);
    });
}
