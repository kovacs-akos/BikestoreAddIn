using System;
using System.Collections.Generic;
using System.Text;
using ExcelDna.Integration;

namespace BikestoreAddIn
{
    public static class BikestoreFunctions
    {
        [ExcelFunction(Description = "Nagyot csinál a kicsiből")]
        public static string MyUpper(string lower) => lower.ToUpper();

        [ExcelFunction(Description = "Példa spill array")]
        public static object[,] MySpillArray() => new object[,]
        {
        { 1, "hello" },
        { 3.14, true }
        };
    }
}
