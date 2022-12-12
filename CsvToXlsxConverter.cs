using System.Globalization;
using CsvHelper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

static class CsvToXlsxConverter
{
    public static async Task ConvertAsync(
        string csvPath,
        string xlsxPath,
        List<string> decimalColumnIndexes,
        string? decimalSeparator,
        List<string> dateColumnIndexes,
        string? dateFormat,
        CancellationToken ct)
    {
        var rows = await ReadCsvAsync(
            csvPath,
            decimalColumnIndexes,
            decimalSeparator,
            dateColumnIndexes,
            dateFormat,
            ct);
        await WriteXlsxAsync(xlsxPath, rows, ct);
    }

    private static async Task<List<List<object?>>> ReadCsvAsync(
        string csvPath,
        List<string> decimalColumnIndexes,
        string? decimalSeparator,
        List<string> dateColumnIndexes,
        string? dateFormat,
        CancellationToken ct)
    {
        List<int> decimalColumnIds = decimalColumnIndexes.Select(ind => OpenXmlExtensions.ColumnIndexToColumnId(ind)).ToList();
        List<int> dateColumnIds = dateColumnIndexes.Select(ind => OpenXmlExtensions.ColumnIndexToColumnId(ind)).ToList();

        using (var reader = new StreamReader(csvPath))
        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            List<List<object?>> rows = new();

            if (!await csv.ReadAsync())
            {
                return rows;
            }
            
            if (!csv.ReadHeader() || (csv.HeaderRecord?.Length ?? 0) == 0)
            {
                return rows;
            }

            rows.Add(csv.HeaderRecord!.Cast<object?>().ToList());

            while (await csv.ReadAsync())
            {
                ct.ThrowIfCancellationRequested();
                
                var row = new List<object?>();
                rows.Add(row);

                int i = 0;
                foreach (var header in csv.HeaderRecord!)
                {
                    var stringValue = csv.GetField(header);

                    if (decimalColumnIds.Contains(i))
                    {
                        decimal? decimalValue = ParseDecimal(stringValue, decimalSeparator);
                        row.Add(decimalValue);
                    }
                    else if (dateColumnIds.Contains(i))
                    {
                        DateTime? dateValue = ParseDate(stringValue, dateFormat);
                        row.Add(dateValue);
                    }
                    else
                    {
                        row.Add(stringValue);
                    }
                    
                    i++;
                }
            }

            return rows;
        }
    }

    private static async Task WriteXlsxAsync(string xlsxPath, List<List<object?>> rows, CancellationToken ct)
    {
        using var outputStream = new MemoryStream();
        {
            using var document = SpreadsheetDocument.Create(outputStream, SpreadsheetDocumentType.Workbook, autoSave: true);

            document.InitWorkbook();

            var worksheet = document.AddWorksheet("Sheet1");

            foreach (var row in rows)
            {
                var xlsxRow = worksheet.AppendRelativeRow();

                foreach (var cellValue in row)
                {
                    if (cellValue is decimal decimalValue)
                    {
                        xlsxRow.AppendRelativeNumberCell(number: decimalValue);
                        continue;
                    }
                    
                    if (cellValue is DateTime dateTimeValue)
                    {
                        xlsxRow.AppendRelativeDateCell(date: dateTimeValue, styleIndex: 1);
                        continue;
                    }
                    
                    if (cellValue is string strValue)
                    {
                        xlsxRow.AppendRelativeInlineStringCell(text: strValue);
                        continue;
                    }

                    xlsxRow.AppendRelativeInlineStringCell(text: string.Empty);
                }
            }

            worksheet.Finalize();
        }
        
        outputStream.Position = 0;
        using FileStream fs = new(xlsxPath, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 4096, useAsync: true);
        await outputStream.CopyToAsync(fs, ct);
    }

    private static decimal? ParseDecimal(string? value, string? decimalSeparator)
    {
        if (string.IsNullOrEmpty(value))
        {
            return null;
        }

        NumberFormatInfo nfi;
        if (!string.IsNullOrEmpty(decimalSeparator))
        {
            nfi = new NumberFormatInfo { CurrencyDecimalSeparator = decimalSeparator };
        }
        else
        {
            nfi = CultureInfo.InvariantCulture.NumberFormat;
        }

        if (!decimal.TryParse(value, NumberStyles.Currency, nfi, out var result))
        {
            return null;
        }

        return result;
    }

    private static DateTime? ParseDate(string? value, string? dateFormat)
    {
        if (string.IsNullOrEmpty(value))
        {
            return null;
        }

        if (!string.IsNullOrEmpty(dateFormat))
        {
            if (!DateTime.TryParseExact(value, dateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var result))
            {
                return null;
            }

            return result;
        }
        else
        {
            if (!DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var result))
            {
                return null;
            }

            return result;
        }
    }
}