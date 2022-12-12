using System.Text;
using Microsoft.Extensions.CommandLineUtils;

using CancellationTokenSource cts = new();

Console.OutputEncoding = Encoding.UTF8;
Console.CancelKeyPress += (s, e) => {
    e.Cancel = true;
    cts.Cancel();
};

var app = new CommandLineApplication(throwOnUnexpectedArg: false);

var csvPathArg = app.Argument("csv", "input csv file");
var xlsxPathArg = app.Argument("xlsx", "output xlsx file");
var decimalColumnIndexesOpt = app.Option("-d|--decimal-columns", "decimal columns as excel indexes (A to ZZZ)", CommandOptionType.MultipleValue);
var decimalSeparatorOpt = app.Option("-ds|--decimal-separator", "decimal separator", CommandOptionType.SingleValue);
var dateColumnIndexesOpt = app.Option("-dt|--date-columns", "date columns as excel indexes (A to ZZZ)", CommandOptionType.MultipleValue);
var dateFormatOpt = app.Option("-dtf|--date-format", "date format", CommandOptionType.SingleValue);

app.OnExecute(async () =>
{
    string csvPath = csvPathArg.Value;
    string xlsxPath = xlsxPathArg.Value;

    if (string.IsNullOrEmpty(csvPath))
    {
        app.ShowHelp();
        return 2;
    }

    if (string.IsNullOrEmpty(xlsxPath))
    {
        xlsxPath = Path.ChangeExtension(csvPath, ".xlsx");
    }

    var decimalColumnIndexes = decimalColumnIndexesOpt.Values;
    var decimalSeparator = decimalSeparatorOpt.Values.FirstOrDefault();
    var dateColumnIndexes = dateColumnIndexesOpt.Values;
    var dateFormat = dateFormatOpt.Values.FirstOrDefault();

    try
    {
        await CsvToXlsxConverter.ConvertAsync(
            csvPath,
            xlsxPath,
            decimalColumnIndexes,
            decimalSeparator,
            dateColumnIndexes,
            dateFormat,
            cts.Token);
    }
    catch (OperationCanceledException)
    {
        return 1;
    }

    return 0;
});

return app.Execute(args);
