using PivotTableSimulator;

namespace PivotTableSimulatorUI
{
    internal class Program
    {
        static void Main()
        {
            var headerRawsAmount = GetHeaderRawsAmount();
            var metaColumnsAmount = GetMetaColumnsAmount();

            FileCombiner.Combine(headerRawsAmount);
            SimulatorService.CreateTable(metaColumnsAmount, headerRawsAmount);
            HeaderService.ReplaceOnInit(metaColumnsAmount, headerRawsAmount);
        }

        private static int GetHeaderRawsAmount()
        {
            return GetAmount("Введите количество строк в заголовке: ");
        }

        private static int GetMetaColumnsAmount()
        {
            return GetAmount("Введите количество столбцов для сортировки: ");
        }

        private static int GetAmount(string message)
        {
            Console.Write(message);
            return int.TryParse(Console.ReadLine(), out int amount)
                ? amount
                : throw new InvalidCastException("Необходимо ввести целочисленное значение");
        }
    }
}
