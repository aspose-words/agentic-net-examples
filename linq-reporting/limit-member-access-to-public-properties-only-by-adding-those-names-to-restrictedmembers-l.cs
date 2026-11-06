using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    // Public properties that will be accessed from the template.
    public string CustomerName { get; set; } = "John Doe";
    public int OrderId { get; set; } = 12345;

    // Private member – not exposed to the reporting engine.
    private string SecretCode { get; set; } = "XYZ";
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document containing LINQ Reporting tags.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";

        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert tags that reference the public properties of the Order class.
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln("Order ID: <<[model.OrderId]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data source.
        // -----------------------------------------------------------------
        var order = new Order();

        // -----------------------------------------------------------------
        // 4. Configure the reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();

        // No need to add any restricted members – the engine already
        // limits access to public members only. Removing the invalid
        // RestrictedMembers calls fixes the compilation error.

        // -----------------------------------------------------------------
        // 5. Build the report.
        // -----------------------------------------------------------------
        engine.BuildReport(reportDoc, order, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
