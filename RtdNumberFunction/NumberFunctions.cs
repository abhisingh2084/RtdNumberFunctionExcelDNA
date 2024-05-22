using ExcelDna.Integration;

public static class NumberFunctions
{
    [ExcelFunction(Name = "func.Number", Description = "Returns an increasing number every second")]
    public static object GetIncrementingNumber()
    {
        return XlCall.RTD("NumberServer", null, "NOW");
    }
}
