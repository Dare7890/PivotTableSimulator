namespace PivotTableSimulator
{
    public static class FileCombiner
    {
        private static readonly string _filesExtension = "*.xlsx";
        private static readonly IReadOnlyCollection<string> _ignoreFileNames =
        [
            Constants.mergedFilePath,
            Constants.pivotTableFilePath
        ];

        public static void Combine(int headerRawsAmount)
        {
            var books = Directory.GetFiles(Constants.filesFolderPath, _filesExtension, SearchOption.TopDirectoryOnly)
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
                        headerRawsAmount,
                        firstColumn: 0, sheet.Cells.MaxDataRow - headerRawsAmount + 1, sheet.Cells.MaxDataColumn + 1);

                    var destRange = resultBook.Worksheets[sheet.Name].Cells.CreateRange(
                        sheetRowsAmount,
                        firstColumn: 0,
                        sourceRange.RowCount,
                        sourceRange.ColumnCount);

                    destRange.Copy(sourceRange);
                }
            }

            resultBook.Save(Constants.mergedFilePath);
        }
    }
}
