using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using ReportDealerGoogleSheets;
using System.Globalization;

LogsHandler.LogWriter(0, $"Запуск обработки файлов отчётов Excel...");
try
{
    IConfigurationRoot config = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
    SettingConfig settings = new SettingConfig();
    config.Bind(settings);

    int countIteration = 0;
    string[] files = GetExcelFiles();
    GoogleSheets googleSheets = new GoogleSheets(settings.SpreadsheetId);
    foreach (var file in files)
    {
        FileInfo fileInfo = new FileInfo(file);

        string nameSheet = fileInfo.Name.Split('.')[0];

        ExcelFileReadingAndWriteInList(file);
        List<DataReport> dataReport = DataReport.ListDataReports[countIteration];

        googleSheets.CreateSheet(nameSheet);
        googleSheets.WriteCells(dataReport, nameSheet);

        RenameFile(file);
        countIteration++;
    }
}
catch (Exception ex)
{
    LogsHandler.LogWriter((LogsHandler.Status)1, ex.Message + "\n" + ex.StackTrace);
}
LogsHandler.LogWriter(0, $"Работа программы завершена");

static string[] GetExcelFiles()
{
    string[] allExcelFiles = Directory.GetFiles(Directory.GetCurrentDirectory(),"*.xlsx");
    List<string> necessaryExcelFiles = new List<string>();
    foreach (string filename in allExcelFiles)
    {
        if (!filename.Contains("processed"))
            necessaryExcelFiles.Add(filename);
    }
    LogsHandler.LogWriter(0, $"Количество найденных excel файлов для загруки: {necessaryExcelFiles.Count}");
    return necessaryExcelFiles.ToArray();
}
static void ExcelFileReadingAndWriteInList(string pathFile)
{
    using (var package = new ExcelPackage(new FileInfo(pathFile)))
    {
        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
        int rowCount = worksheet.Dimension.Rows;
        List<DataReport> listDataReports = new List<DataReport>();

        for (int row = 1; row <= rowCount; row++)
        {
            string id = worksheet.Cells[row, 3].Text;
            string name = worksheet.Cells[row, 4].Text;
            string amount = worksheet.Cells[row, 13].Text;

            if (!String.IsNullOrEmpty(id))
            {
                listDataReports.Add(new DataReport(id, name, amount));
            }
        }

        DeleteDuplicates(listDataReports);

        if (listDataReports.Count > 0)
            DataReport.ListDataReports.Add(listDataReports);
        else
            throw new Exception("Не найдены данные для загрузки!");
    }    
}
static void RenameFile(string pathFile)
{
    string newDirectory = Directory.GetCurrentDirectory() + @"\archive\";
    FileInfo fileInfo = new FileInfo(pathFile);
    string newName = "processed_" + fileInfo.Name;

    if (!Directory.Exists("archive"))
        Directory.CreateDirectory("archive");

    LogsHandler.LogWriter(0, $"Переименовываем и перемещаем обработанный файл excel {fileInfo.Name} в {newName}");
    File.Move(pathFile, newDirectory + newName);
}
static void DeleteDuplicates(List<DataReport> dataReports)
{
    LogsHandler.LogWriter(0, $"Удаляем дубли строк по ID клиента, до удаление количество строк: {dataReports.Count}");
    Dictionary<(string id, string name), DataReport> uniqueDataReports = new Dictionary<(string, string), DataReport>();

    foreach (var report in dataReports)
    {
        var key = (report.clientID, report.clientName);

        if (uniqueDataReports.ContainsKey(key))
        {
            decimal amountDuplicate = StringAsDecimal(report.clientAmount);
            decimal amountOriginal = StringAsDecimal(uniqueDataReports[key].clientAmount);
            decimal result = amountDuplicate + amountOriginal;
            uniqueDataReports[key].clientAmount = result.ToString();
        }
        else
        {
            uniqueDataReports[key] = report;
        }
    }

    dataReports.Clear();
    dataReports.AddRange(uniqueDataReports.Values);
    LogsHandler.LogWriter(0, $"Дубли удалены, оставшиеся количество строк: {dataReports.Count}");
}
static decimal StringAsDecimal(string amountString)
{
    amountString = amountString.Replace("₽", "");
    amountString = amountString.Replace(" ", "");
    amountString = amountString.Replace(",", ".");
    return Decimal.Parse(amountString, NumberStyles.Any, CultureInfo.InvariantCulture);
}
