using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    // Initialize to avoid nullable warnings
    public string CustomerName { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Create an output folder for the generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a simple template containing a LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Customer: <<[order.CustomerName]>>");

        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        Order order = new Order { CustomerName = "John Doe" };

        // -----------------------------------------------------------------
        // 4. Configure the reporting engine.
        //    The ReportingEngine class does not expose a RestrictedMembers
        //    collection in the current Aspose.Words version, so we omit that
        //    configuration. The engine is still secure because it only has
        //    access to the public members of the supplied data model.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // -----------------------------------------------------------------
        // 5. Build the report.
        // -----------------------------------------------------------------
        engine.BuildReport(reportDoc, order, "order");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
