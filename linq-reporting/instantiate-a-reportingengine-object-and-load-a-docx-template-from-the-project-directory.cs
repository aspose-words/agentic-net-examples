using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code pages for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var order = new Order { CustomerName = "John Doe" };

        // Create a DOCX template with a LINQ Reporting tag.
        const string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template document.
        var doc = new Document(templatePath);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        const string outputPath = "report.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[order.CustomerName]>>");
        doc.Save(path);
    }

    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
    }
}
