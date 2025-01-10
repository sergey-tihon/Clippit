using System.Globalization;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Clippit.Tests.Common.Samples
{
    public class ChartUpdaterSamples(ITestOutputHelper log) : TestsBase(log)
    {
        private static string GetFilePath(string path) => Path.Combine("../../../Common/Samples/ChartUpdater/", path);

        [Theory]
        [InlineData("Chart-Cached-Data-01.docx")]
        [InlineData("Chart-Cached-Data-02.docx")]
        [InlineData("Chart-Cached-Data-03.docx")]
        [InlineData("Chart-Cached-Data-04.docx")]
        [InlineData("Chart-Cached-Data-05.docx")]
        [InlineData("Chart-Cached-Data-06.docx")]
        [InlineData("Chart-Cached-Data-07.docx")]
        [InlineData("Chart-Embedded-Xlsx-01.docx")]
        [InlineData("Chart-Embedded-Xlsx-02.docx")]
        [InlineData("Chart-Embedded-Xlsx-03.docx")]
        [InlineData("Chart-Embedded-Xlsx-04.docx")]
        [InlineData("Chart-Embedded-Xlsx-05.docx")]
        [InlineData("Chart-Embedded-Xlsx-06.docx")]
        [InlineData("Chart-Embedded-Xlsx-07.docx")]
        [InlineData("Chart-Embedded-Xlsx-08.docx")]
        [InlineData("Chart-Embedded-Xlsx-10.docx")]
        public void UpdateWords(string fileName)
        {
            var srcFile = new FileInfo(GetFilePath(fileName));
            var fName = Path.Combine(TempDir, srcFile.Name);
            File.Copy(srcFile.FullName, fName, true);

            var fi = new FileInfo(fName);
            var newFileName = "Updated-" + fi.Name;
            var fi2 = new FileInfo(Path.Combine(TempDir, newFileName));
            File.Copy(fi.FullName, fi2.FullName, true);

            using var wDoc = WordprocessingDocument.Open(fi2.FullName, true);
            var chart1Data = new ChartData
            {
                SeriesNames = new[] { "Car", "Truck", "Van", "Bike", "Boat" },
                CategoryDataType = ChartDataType.String,
                CategoryNames = new[] { "Q1", "Q2", "Q3", "Q4" },
                Values = new[]
                {
                    new double[] { 100, 310, 220, 450 },
                    new double[] { 200, 300, 350, 411 },
                    new double[] { 80, 120, 140, 600 },
                    new double[] { 120, 100, 140, 400 },
                    new double[] { 200, 210, 210, 480 },
                },
            };
            ChartUpdater.UpdateChart(wDoc, "Chart1", chart1Data);

            var chart2Data = new ChartData
            {
                SeriesNames = new[] { "Series" },
                CategoryDataType = ChartDataType.String,
                CategoryNames = new[] { "Cars", "Trucks", "Vans", "Boats" },
                Values = new[] { new double[] { 320, 112, 64, 80 } },
            };
            ChartUpdater.UpdateChart(wDoc, "Chart2", chart2Data);

            var chart3Data = new ChartData
            {
                SeriesNames = new[] { "X1", "X2", "X3", "X4", "X5", "X6" },
                CategoryDataType = ChartDataType.String,
                CategoryNames = new[] { "Y1", "Y2", "Y3", "Y4", "Y5", "Y6" },
                Values = new[]
                {
                    new[] { 3.0, 2.1, .7, .7, 2.1, 3.0 },
                    new[] { 3.0, 2.1, .8, .8, 2.1, 3.0 },
                    new[] { 3.0, 2.4, 1.2, 1.2, 2.4, 3.0 },
                    new[] { 3.0, 2.7, 1.7, 1.7, 2.7, 3.0 },
                    new[] { 3.0, 2.9, 2.5, 2.5, 2.9, 3.0 },
                    new[] { 3.0, 3.0, 3.0, 3.0, 3.0, 3.0 },
                },
            };
            ChartUpdater.UpdateChart(wDoc, "Chart3", chart3Data);

            var chart4Data = new ChartData
            {
                SeriesNames = new[] { "Car", "Truck", "Van" },
                CategoryDataType = ChartDataType.DateTime,
                CategoryFormatCode = 14,
                CategoryNames = new[]
                {
                    ToExcelInteger(new DateTime(2013, 9, 1)),
                    ToExcelInteger(new DateTime(2013, 9, 2)),
                    ToExcelInteger(new DateTime(2013, 9, 3)),
                    ToExcelInteger(new DateTime(2013, 9, 4)),
                    ToExcelInteger(new DateTime(2013, 9, 5)),
                    ToExcelInteger(new DateTime(2013, 9, 6)),
                    ToExcelInteger(new DateTime(2013, 9, 7)),
                    ToExcelInteger(new DateTime(2013, 9, 8)),
                    ToExcelInteger(new DateTime(2013, 9, 9)),
                    ToExcelInteger(new DateTime(2013, 9, 10)),
                    ToExcelInteger(new DateTime(2013, 9, 11)),
                    ToExcelInteger(new DateTime(2013, 9, 12)),
                    ToExcelInteger(new DateTime(2013, 9, 13)),
                    ToExcelInteger(new DateTime(2013, 9, 14)),
                    ToExcelInteger(new DateTime(2013, 9, 15)),
                    ToExcelInteger(new DateTime(2013, 9, 16)),
                    ToExcelInteger(new DateTime(2013, 9, 17)),
                    ToExcelInteger(new DateTime(2013, 9, 18)),
                    ToExcelInteger(new DateTime(2013, 9, 19)),
                    ToExcelInteger(new DateTime(2013, 9, 20)),
                },
                Values = new[]
                {
                    new double[] { 1, 2, 3, 2, 3, 4, 5, 4, 5, 6, 5, 4, 5, 6, 7, 8, 7, 8, 8, 9 },
                    new double[] { 2, 3, 3, 4, 4, 5, 6, 7, 8, 7, 8, 9, 9, 9, 7, 8, 9, 9, 10, 11 },
                    new double[] { 2, 3, 3, 3, 3, 2, 2, 2, 3, 2, 3, 3, 4, 4, 4, 3, 4, 5, 5, 4 },
                },
            };
            ChartUpdater.UpdateChart(wDoc, "Chart4", chart4Data);
        }

        [Theory]
        [InlineData("Chart-Cached-Data-41.pptx")]
        [InlineData("Chart-Embedded-Xlsx-41.pptx")]
        public void UpdatePowerPoints(string fileName)
        {
            var srcFile = new FileInfo(GetFilePath(fileName));
            var fName = Path.Combine(TempDir, srcFile.Name);
            File.Copy(srcFile.FullName, fName, true);

            var fi = new FileInfo(fName);
            var newFileName = "Updated-" + srcFile.Name;
            var fi2 = new FileInfo(Path.Combine(TempDir, newFileName));
            File.Copy(fi.FullName, fi2.FullName, true);

            using var pDoc = PresentationDocument.Open(fi2.FullName, true);
            var chart1Data = new ChartData
            {
                SeriesNames = new[] { "Car", "Truck", "Van" },
                CategoryDataType = ChartDataType.String,
                CategoryNames = new[] { "Q1", "Q2", "Q3", "Q4" },
                Values = new[]
                {
                    new double[] { 320, 310, 320, 330 },
                    new double[] { 201, 224, 230, 221 },
                    new double[] { 180, 200, 220, 230 },
                },
            };
            ChartUpdater.UpdateChart(pDoc, 1, chart1Data);
        }

        private static string ToExcelInteger(DateTime dateTime)
        {
            return dateTime.ToOADate().ToString(CultureInfo.CurrentCulture);
        }
    }
}
