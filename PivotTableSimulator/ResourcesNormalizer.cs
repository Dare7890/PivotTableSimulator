using Aspose.Cells;

namespace PivotTableSimulator
{
    public static class ResourcesNormalizer
    {
        public static void Normalize()
        {
            NormalizeFile();
            NormalizeFolder();
        }

        private static void NormalizeFile()
        {
            var wb = new Workbook(Constants.PIVOT_TABLE_FILE_PATH);

            var deletedSheetIndexes = wb.Worksheets
                .Where(w => w.Name != Constants.RESULT_SHEET_NAME)
                .Select(w => w.Index)
                .ToList();

            for (int i = deletedSheetIndexes.Count - 1; i >= 0; i--)
            {
                wb.Worksheets.RemoveAt(deletedSheetIndexes[i]);
            }

            wb.Save(Constants.PIVOT_TABLE_FILE_PATH);
        }

        private static void NormalizeFolder()
        {
            File.Delete(Constants.MERGED_FILE_PATH);
        }
    }
}
