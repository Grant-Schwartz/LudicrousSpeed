using ExcelDna.Integration;

namespace WarpSpeed.ExcelAddIn
{
    public static class WarpSpeedDiagnostics
    {
        [ExcelFunction(
            Name = "WARPSPEED_PING",
            Description = "Returns a message when the WarpSpeed Excel add-in is loaded.")]
        public static string Ping()
        {
            return "WarpSpeed add-in loaded";
        }
    }
}
