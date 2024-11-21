namespace PivotTableSimulator
{
    public static class Constants
    {
        public static readonly string FILES_FOLDER_PATH = GetResourcesFullPath(string.Empty);
        public static readonly string MERGED_FILE_PATH = GetResourcesFullPath("Merged_Table.xlsx");
        public static readonly string PIVOT_TABLE_FILE_PATH = GetResourcesFullPath("PivotTable_Table.xlsx");

        public static readonly string RESULT_SHEET_NAME = "resultSheet";

        private static string GetResourcesFullPath(string fileName)
        {
            return $"{Environment.CurrentDirectory}\\..\\..\\..\\..\\Resources" +
                (string.IsNullOrEmpty(fileName) ? string.Empty : string.Format("\\{0}", fileName));
        }
    }
}
