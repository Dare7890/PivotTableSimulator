namespace PivotTableSimulator
{
    public static class FileCombiner
    {
        private static readonly string _filesExtension = "*.xlsx";
        private static readonly IReadOnlyCollection<string> _ignoreFileNames =
        [
            Constants.MERGED_FILE_PATH,
            Constants.PIVOT_TABLE_FILE_PATH
        ];

        public static void Combine(int headerRawsAmount)
        {
            var books = Directory.GetFiles(Constants.FILES_FOLDER_PATH, _filesExtension, SearchOption.TopDirectoryOnly)
                .Where(f => !_ignoreFileNames.Contains(f))
                .Select(f => new Aspose.Cells.Workbook(f))
                .ToArray();

            if (books.Length == 0)
            {
                throw new InvalidDataException("Отсутствуют файлы для объединения");
            }

            var resultBook = new Aspose.Cells.Workbook();
            foreach (var sheet in books.SelectMany(b => b.Worksheets))
            {
                if (resultBook.Worksheets[sheet.Name] == null)
                {
                    resultBook.Worksheets.Add(sheet.Name);
                    resultBook.Worksheets[sheet.Name].Copy(sheet);
                }
                else
                {
                    var sheetRowsAmount = resultBook.Worksheets[sheet.Name].Cells.Rows.Count;

                    var sourceRange = sheet.Cells.CreateRange(
                        firstRow: headerRawsAmount,
                        firstColumn: 0,
                        totalRows: sheet.Cells.MaxDataRow - headerRawsAmount + 1,
                        totalColumns: sheet.Cells.MaxDataColumn + 1);

                    var destRange = resultBook.Worksheets[sheet.Name].Cells.CreateRange(
                        firstRow: sheetRowsAmount,
                        firstColumn: 0,
                        totalRows: sourceRange.RowCount,
                        totalColumns: sourceRange.ColumnCount);

                    destRange.Copy(sourceRange);
                }
            }

            resultBook.Save(Constants.MERGED_FILE_PATH);
        }
    }
}
