using Aspose.Cells;

namespace PivotTableSimulator
{
    public static class HeaderService
    {
        public static void ReplaceOnInit(int metaColumnsRawsAmount, int headerRawsAmount)
        {
            var destWb = new Workbook(Constants.PIVOT_TABLE_FILE_PATH);
            var destSheet = destWb.Worksheets[Constants.RESULT_SHEET_NAME];

            CopyTable(destSheet);
            CopyHeader(destSheet, headerRawsAmount);

            for (int i = 0; i < metaColumnsRawsAmount; i++)
            {
                var initRaw = headerRawsAmount - 1;
                for (int j = headerRawsAmount - 1; j < destSheet.Cells.MaxRow + 1; j++)
                {
                    if (string.IsNullOrEmpty(destSheet.Cells[j, i].StringValue) ||
                        j == initRaw)
                    {
                        continue;
                    }

                    destSheet.Cells.Merge(initRaw, i, j - initRaw, 1);
                    initRaw = j;
                }
            }

            destWb.Save(Constants.PIVOT_TABLE_FILE_PATH);
        }

        private static void CopyHeader(Worksheet destSheet, int headerRawsAmount)
        {
            var mergedWb = new Workbook(Constants.MERGED_FILE_PATH);
            var mergedSheet = mergedWb.Worksheets[1];

            var headerSourceRange = mergedSheet.Cells.CreateRange(0, 0, headerRawsAmount, mergedSheet.Cells.MaxColumn + 1);
            var headerDestRange = destSheet.Cells.CreateRange(0, 0, headerRawsAmount, headerSourceRange.ColumnCount);

            headerDestRange.Copy(headerSourceRange);
        }

        private static void CopyTable(Worksheet destSheet)
        {
            var mergedWb = new Workbook(Constants.PIVOT_TABLE_FILE_PATH);
            var mergedSheet = mergedWb.Worksheets[Constants.RESULT_SHEET_NAME];
            var sourceRange = mergedSheet.Cells.CreateRange(0, 0, mergedSheet.Cells.MaxRow, mergedSheet.Cells.MaxColumn + 1);

            var destRange = destSheet.Cells.CreateRange(0, 0, mergedSheet.Cells.MaxRow, mergedSheet.Cells.MaxColumn + 1);

            if (destSheet.PivotTables.Any())
            {
                destSheet.PivotTables.RemoveAt(0);
            }

            destRange.Copy(sourceRange);
        }
    }
}
