using Clippit.Excel;

namespace Clippit.Tests.Excel.Samples
{
    public class SmlDataRetrieverSamples() : Clippit.Tests.TestsBase
    {
        private static string GetFilePath(string path) =>
            Path.Combine("../../../Excel/Samples/SmlDataRetriever/", path);

        [Test]
        public void Sample1()
        {
            var fi = new FileInfo(GetFilePath("SampleSpreadsheet.xlsx"));
            // Retrieve range from Sheet1
            var data = SmlDataRetriever.RetrieveRange(fi.FullName, "Sheet1", "A1:C3");
            Console.WriteLine(data.ToString());
            // Retrieve entire sheet
            data = SmlDataRetriever.RetrieveSheet(fi.FullName, "Sheet1");
            Console.WriteLine(data.ToString());
            // Retrieve table
            data = SmlDataRetriever.RetrieveTable(fi.FullName, "VehicleTable");
            Console.WriteLine(data.ToString());
        }
    }
}
