using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create an output folder for the generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the final report.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // Sample data model.
        Order order = new Order
        {
            CustomerName = "John Doe"
        };

        // -----------------------------------------------------------------
        // 1. Build the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Use the default LINQ Reporting delimiters << and >>.
        // The tag <<[order.CustomerName]>> will be replaced with the value from the model.
        builder.Writeln("Customer Name: <<[order.CustomerName]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and generate the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // No custom delimiters are needed because we used the default ones.
        // Build the report using the root object name "order".
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save(reportPath);
    }
}

// Simple public data model class.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
}
