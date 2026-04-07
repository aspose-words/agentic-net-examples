using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string xmlPath = Path.Combine(outputDir, "orders.xml");
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPath = Path.Combine(outputDir, "report.docx");

        // 1. Create a sample XML file with a list of orders.
        CreateSampleXml(xmlPath);

        // 2. Build a Word template containing LINQ Reporting tags.
        CreateTemplateDocument(templatePath);

        // 3. Load the template document.
        Document doc = new Document(templatePath);

        // 4. Create an XmlDataSource from the XML file.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // 5. Configure the ReportingEngine.
        ReportingEngine.UseReflectionOptimization = true; // Enable reflection optimization.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(System.Math)); // Register a known external type (example).

        // 6. Build the report. The data source name is "orders".
        engine.BuildReport(doc, dataSource, "orders");

        // 7. Save the generated report.
        doc.Save(reportPath);
    }

    // Generates a simple XML file with many Order elements.
    private static void CreateSampleXml(string path)
    {
        using (StreamWriter writer = new StreamWriter(path))
        {
            writer.WriteLine("<Orders>");
            for (int i = 1; i <= 100; i++)
            {
                writer.WriteLine($"  <Order>");
                writer.WriteLine($"    <Id>{i}</Id>");
                writer.WriteLine($"    <Customer>Customer {i}</Customer>");
                writer.WriteLine($"    <Amount>{i * 10.5}</Amount>");
                writer.WriteLine($"  </Order>");
            }
            writer.WriteLine("</Orders>");
        }
    }

    // Creates a Word document with LINQ Reporting tags.
    private static void CreateTemplateDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("==============");
        builder.Writeln();
        // Begin foreach loop over the collection named "orders".
        builder.Writeln("<<foreach [order in orders]>>");
        // Output fields from each Order element.
        builder.Writeln("Id: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.Customer]>>");
        builder.Writeln("Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}
