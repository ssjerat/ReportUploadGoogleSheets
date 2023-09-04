using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

public class GoogleSheets
{
    static string[] Scopes = { SheetsService.Scope.Spreadsheets };
    static string ApplicationName = "Table";
    static string SpreadsheetId;
    public GoogleSheets(string id) { SpreadsheetId = id; }

    public static UserCredential GetCredential()
    {
        try
        {
            UserCredential credential;
            using (var stream =
                   new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }
            return credential;
        }
        catch (FileNotFoundException ex)
        {
            LogsHandler.LogWriter((LogsHandler.Status)1, ex.Message + "\n" + ex.StackTrace);
            throw new FileNotFoundException(ex.Message);
        }
    }
    public void WriteCells(List<DataReport> dataReports, string sheetName)
    {
        var service = new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = GetCredential(),
            ApplicationName = ApplicationName
        });

        SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum insertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;

        var countRow = dataReports.Count;
        string range = $"{sheetName}!A1:C{countRow}";

        var rangeData = dataReports.Select(data => new List<object>
        {
            data.clientID,
            data.clientName,
            data.clientAmount
        }).Cast<IList<object>>().ToList();

        var valueRange = new ValueRange
        {
            Values = rangeData
        };

        SpreadsheetsResource.ValuesResource.AppendRequest request = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
        request.ValueInputOption = valueInputOption;
        request.InsertDataOption = insertDataOption;

        LogsHandler.LogWriter(0, $"Попытка отправить запрос на запись в гугл таблицу в диапазон - {range}...");
        AppendValuesResponse response = request.Execute();
        LogsHandler.LogWriter(0, $"Запись в гугл таблицу успешно завершена - {range}");
    }
    public void CreateSheet(string nameSheet)
    {
        var service = new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = GetCredential(),
            ApplicationName = ApplicationName
        });

        string newSheetName = nameSheet;

        var addSheetRequest = new AddSheetRequest
        {
            Properties = new SheetProperties
            {
                Title = newSheetName
            }
        };

        var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
        {
            Requests = new List<Request>
        {
            new Request
            {
                AddSheet = addSheetRequest
            }
        }
        };
        LogsHandler.LogWriter(0, $"Попытка в гугл таблице создать лист с именем: {nameSheet} ...");
        service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, SpreadsheetId).Execute();
        LogsHandler.LogWriter(0, $"Лист с именем {nameSheet} успешно создан");
    }
}

