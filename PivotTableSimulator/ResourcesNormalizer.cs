using Aspose.Cells;

namespace PivotTableSimulator
{
    public static class ResourcesNormalizer
    {
        private const string _resultSheetName = "Sheet1";

        public static void Normalize()
        {
            NormalizeFile();
            NormalizeFolder();
        }

        private static void NormalizeFile()
        {
            var wb = new Workbook(Constants.pivotTableFilePath);

            var deletedSheetIndexes = wb.Worksheets
                .Where(w => w.Name != _resultSheetName)
                .Select(w => w.Index)
                .ToList();

            for (int i = deletedSheetIndexes.Count - 1; i >= 0; i--)
            {
                wb.Worksheets.RemoveAt(deletedSheetIndexes[i]);
            }

            wb.Save(Constants.pivotTableFilePath);
        }

        private static void NormalizeFolder()
        {
            File.Delete(Constants.mergedFilePath);
        }
    }
}
