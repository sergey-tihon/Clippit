using System.IO;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Excel.Samples
{
    public class WorksheetAccessorSamples : TestsBase
    {
        public WorksheetAccessorSamples(ITestOutputHelper log) : base(log)
        {
        }
        
        private static string GetFilePath(string path) =>
            Path.Combine("../../../Excel/Samples/WorksheetAccessor/", path);

        [Fact]
        public void Sample1()
        {
            var sourceFile = GetFilePath("Formulas.xlsx");
            // Change sheet name in formulas
            using (var streamDoc = new OpenXmlMemoryStreamDocument( SmlDocument.FromFileName(sourceFile)))
            {
                using (var doc = streamDoc.GetSpreadsheetDocument())
                {
                    WorksheetAccessor.FormulaReplaceSheetName(doc, "Source", "'Source 2'");
                }
                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(TempDir, "FormulasUpdated.xlsx"));
            }

            // Change sheet name in formulas
            using (var streamDoc = new OpenXmlMemoryStreamDocument(SmlDocument.FromFileName(sourceFile)))
            {
                using (var doc = streamDoc.GetSpreadsheetDocument())
                {
                    var sheet = WorksheetAccessor.GetWorksheet(doc, "References");
                    WorksheetAccessor.CopyCellRange(doc, sheet, 1, 1, 7, 5, 4, 8);
                }
                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(TempDir, "FormulasCopied.xlsx"));
            }
        }
    }
}
