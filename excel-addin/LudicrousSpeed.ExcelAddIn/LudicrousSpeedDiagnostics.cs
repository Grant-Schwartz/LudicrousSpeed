using ExcelDna.Integration;

namespace LudicrousSpeed.ExcelAddIn
{
    public static class LudicrousSpeedDiagnostics
    {
        [ExcelFunction(
            Name = "LUDICROUS_PING",
            Description = "Returns a message when the LudicrousSpeed Excel add-in is loaded.")]
        public static string Ping()
        {
            return "LudicrousSpeed add-in loaded";
        }
    }
}
