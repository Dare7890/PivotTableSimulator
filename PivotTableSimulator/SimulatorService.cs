using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableSimulator
{
    public static class SimulatorService
    {
        private const string _pivotTableName = "PivotTable";

        public static void CreateTable(int metaColumnsRawsAmount, int headerRawsAmount)
        {
            var wb = new Workbook(Constants.MERGED_FILE_PATH);
            var pivotTableSheet = wb.Worksheets[0];
            pivotTableSheet.Name = Constants.RESULT_SHEET_NAME;

            var sheet = wb.Worksheets[1];
            var mergedAreasList = sheet.Cells.GetMergedAreas();
            var range = sheet.Cells.CreateRange(
                firstRow: headerRawsAmount - 1,
                firstColumn: 0,
                totalRows: sheet.Cells.MaxRow + (headerRawsAmount - 1),
                totalColumns: sheet.Cells.MaxColumn + 1);

            range.UnMerge();
            UnionSeveralCells(sheet, mergedAreasList);

            var pivotTables = pivotTableSheet.PivotTables;
            var index = pivotTables.Add(range.RefersTo, row: headerRawsAmount - 1, column: 0, _pivotTableName);

            var pivotTable = pivotTables[index];
            InitDefaultPivotTableSettings(pivotTable);

            for (int i = 0; i < metaColumnsRawsAmount; i++)
            {
                BuildMetaColumns(pivotTable, i);
            }

            for (int i = metaColumnsRawsAmount; i < range.ColumnCount; i++)
            {
                BuildCalculateColumns(pivotTable, i);
            }

            pivotTable.AddFieldToArea(PivotFieldType.Column, pivotTable.DataField);

            pivotTable.RefreshData();
            pivotTable.CalculateData();

            wb.Save(Constants.PIVOT_TABLE_FILE_PATH);
        }

        private static void InitDefaultPivotTableSettings(PivotTable pivotTable)
        {
            pivotTable.RowGrand = false;
            pivotTable.ColumnGrand = false;
            pivotTable.DataFieldHeaderName = string.Empty;
            pivotTable.ShowDrill = false;
            pivotTable.ShowRowHeaderCaption = false;
        }

        private static void UnionSeveralCells(Worksheet sheet, CellArea[] mergedAreas)
        {
            foreach (var area in mergedAreas)
            {
                var value = sheet.Cells[area.StartRow, area.StartColumn];
                var areaRange = sheet.Cells.CreateRange(
                    firstRow: area.StartRow,
                    firstColumn: area.StartColumn,
                    totalRows: area.EndRow + 1 - area.StartRow,
                    totalColumns: area.EndColumn + 1 - area.StartColumn);

                areaRange.PutValue(value.StringValue, false, false);
            }
        }

        private static void BuildMetaColumns(PivotTable pivotTable, int index)
        {
            pivotTable.AddFieldToArea(PivotFieldType.Row, index);

            var rowField = pivotTable.RowFields[index];
            rowField.IsAutoSort = true;
            rowField.IsAscendSort = true;

            rowField.SetSubtotals(PivotFieldSubtotalType.None, true);
        }

        private static void BuildCalculateColumns(PivotTable pivotTable, int index)
        {
            pivotTable.AddFieldToArea(PivotFieldType.Data, index);
        }
    }
}