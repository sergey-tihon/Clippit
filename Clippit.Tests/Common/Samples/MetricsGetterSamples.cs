namespace Clippit.Tests.Common.Samples;

public class MetricsGetterSamples : TestsBase
{
    private static string GetFilePath(string path) => Path.Combine("../../../Common/Samples/MetricsGetter/", path);

    [Test]
    [Arguments("ContentControls.docx", false)] // No text from content controls
    [Arguments("ContentControls.docx", true)] // With text from content controls
    [Arguments("TrackedRevisions.docx", true)] // Tracked Revisions
    [Arguments("Styles.docx", true)] // Style Hierarchy
    public void Word(string fileName, bool includeTextInControls)
    {
        var fi = new FileInfo(GetFilePath(fileName));
        var settings = new MetricsGetterSettings { IncludeTextInContentControls = includeTextInControls };
        var metrics = MetricsGetter.GetMetrics(fi.FullName, settings);
        Console.WriteLine(metrics.ToString());
    }

    [Test]
    public void Excel()
    {
        var fi = new FileInfo(GetFilePath("Tables.xlsx"));
        var settings = new MetricsGetterSettings
        {
            IncludeTextInContentControls = false,
            IncludeXlsxTableCellData = true,
        };
        var metrics = MetricsGetter.GetMetrics(fi.FullName, settings);
        Console.WriteLine(metrics.ToString());
    }
}
