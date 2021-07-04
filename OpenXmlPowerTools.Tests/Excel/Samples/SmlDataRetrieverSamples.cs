using System.IO;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Excel.Samples
{
    public class SmlDataRetrieverSamples : TestsBase
    {
        public SmlDataRetrieverSamples(ITestOutputHelper log) : base(log)
        {
        }
        
        private static string GetFilePath(string path) =>
            Path.Combine("../../../Excel/Samples/SmlDataRetriever/", path);

        [Fact]
        public void Sample1()
        {
            var fi = new FileInfo(GetFilePath("SampleSpreadsheet.xlsx"));

            // Retrieve range from Sheet1
            var data = SmlDataRetriever.RetrieveRange(fi.FullName, "Sheet1", "A1:C3");
            Log.WriteLine(data.ToString());

            // Retrieve entire sheet
            data = SmlDataRetriever.RetrieveSheet(fi.FullName, "Sheet1");
            Log.WriteLine(data.ToString());

            // Retrieve table
            data = SmlDataRetriever.RetrieveTable(fi.FullName, "VehicleTable");
            Log.WriteLine(data.ToString());
        }
    }
}
