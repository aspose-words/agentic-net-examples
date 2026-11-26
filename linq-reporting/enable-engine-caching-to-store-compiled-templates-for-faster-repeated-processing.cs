using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // Create a simple template document with a LINQ Reporting tag.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document loadedTemplate = new Document(templatePath);

        // Sample data model.
        Order order = new Order { CustomerName = "John Doe" };

        // Create a ReportingEngine.
        // Reusing the same engine instance provides caching of compiled templates.
        ReportingEngine engine = new ReportingEngine();

        // First report generation.
        engine.BuildReport(loadedTemplate, order, "order");
        string outputPath1 = Path.Combine(workDir, "Report1.docx");
        loadedTemplate.Save(outputPath1);

        // Modify data and generate a second report using the same engine.
        order.CustomerName = "Jane Smith";

        // Reload the template to simulate a fresh document instance.
        Document secondTemplate = new Document(templatePath);
        engine.BuildReport(secondTemplate, order, "order");
        string outputPath2 = Path.Combine(workDir, "Report2.docx");
        secondTemplate.Save(outputPath2);
    }
}

// Public data model class.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
}
