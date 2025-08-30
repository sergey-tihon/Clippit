using System.Xml.Linq;
using Clippit.Word;

namespace Clippit.Tests.Word.Samples;

public class DocumentAssemblerSamples : TestsBase
{
    private const string FolderPath = "../../../Word/Samples/DocumentAssembler/";
    private const string TemplateDocumentFilePath = FolderPath + "TemplateDocument.docx";
    private const string TemplateDataFilePath = FolderPath + "Data.xml";

    [Test]
    public async Task Sample1()
    {
        var wmlDoc = new WmlDocument(TemplateDocumentFilePath);
        var data = XElement.Load(TemplateDataFilePath);
        var wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, data, out var templateError);
        await Assert.That(templateError).IsFalse().Because("Errors in template");

        var assembledDoc = new FileInfo(Path.Combine(TempDir, "AssembledDoc.docx"));
        if (assembledDoc.Exists)
            assembledDoc.Delete();
        wmlAssembledDoc.SaveAs(assembledDoc.FullName);
    }

    [Test]
    public void Sample2()
    {
        var dataFile = new FileInfo(Path.Combine(TempDir, "Data.xml"));
        // The following method generates a large data file with random data.
        // In a real world scenario, this is where you would query your data source and produce XML that will drive your document generation process.
        var data = GenerateDataFromDataSource(dataFile);
        var wmlDoc = new WmlDocument(TemplateDocumentFilePath);
        var count = 1;
        foreach (var customer in data.Elements("Customer"))
        {
            var assembledDoc = new FileInfo(Path.Combine(TempDir, $"Letter-{count++:0000}.docx"));
            Console.WriteLine(assembledDoc.Name);
            var wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, customer, out var templateError);
            if (templateError)
            {
                Console.WriteLine("Errors in template.");
                Console.WriteLine("See {assembledDoc.Name} to determine the errors in the template.");
            }

            wmlAssembledDoc.SaveAs(assembledDoc.FullName);
        }
    }

    private static readonly string[] s_productNames =
    [
        "Unicycle",
        "Bicycle",
        "Tricycle",
        "Skateboard",
        "Roller Blades",
        "Hang Glider",
    ];

    private static XElement GenerateDataFromDataSource(FileInfo dataFi)
    {
        var numberOfDocumentsToGenerate = 500;
        var customers = new XElement("Customers");
        var r = new Random();
        for (var i = 0; i < numberOfDocumentsToGenerate; ++i)
        {
            var customer = new XElement(
                "Customer",
                new XElement("CustomerID", i + 1),
                new XElement("Name", "Eric White"),
                new XElement("HighValueCustomer", r.Next(2) == 0 ? "True" : "False"),
                new XElement("Orders")
            );
            var orders = customer.Element("Orders");
            var numberOfOrders = r.Next(10) + 1;
            for (var j = 0; j < numberOfOrders; j++)
            {
                var order = new XElement(
                    "Order",
                    new XAttribute("Number", j + 1),
                    new XElement("ProductDescription", s_productNames[r.Next(s_productNames.Length)]),
                    new XElement("Quantity", r.Next(10)),
                    new XElement("OrderDate", "September 26, 2015")
                );
                orders.Add(order);
            }

            customers.Add(customer);
        }

        customers.Save(dataFi.FullName);
        return customers;
    }
}
