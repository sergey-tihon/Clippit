using Xunit;

namespace Clippit.Tests.Common.Samples
{
    public class MetricsGetterSamples(ITestOutputHelper log) : TestsBase(log)
    {
        private static string GetFilePath(string path) => Path.Combine("../../../Common/Samples/MetricsGetter/", path);

        [Theory]
        [InlineData("ContentControls.docx", false)] // No text from content controls
        [InlineData("ContentControls.docx", true)] // With text from content controls
        [InlineData("TrackedRevisions.docx", true)] // Tracked Revisions
        [InlineData("Styles.docx", true)] // Style Hierarchy
        public void Word(string fileName, bool includeTextInControls)
        {
            var fi = new FileInfo(GetFilePath(fileName));
            var settings = new MetricsGetterSettings { IncludeTextInContentControls = includeTextInControls };
            var metrics = MetricsGetter.GetMetrics(fi.FullName, settings);
            Log.WriteLine(metrics.ToString());
        }

        [Fact]
        public void Excel()
        {
            var fi = new FileInfo(GetFilePath("Tables.xlsx"));
            var settings = new MetricsGetterSettings
            {
                IncludeTextInContentControls = false,
                IncludeXlsxTableCellData = true,
            };
            var metrics = MetricsGetter.GetMetrics(fi.FullName, settings);
            Log.WriteLine(metrics.ToString());
        }
    }
}
