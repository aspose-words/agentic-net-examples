using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for potential encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the template document programmatically.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>>");
        // This property does not exist in the data model and will cause an evaluation error.
        builder.Writeln("Description: <<[item.Description]>>");
        // The <<error>> tag will display the evaluation error inline.
        builder.Writeln("<<error>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data with missing Description property.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Item 1" },
                new Item { Name = "Item 2" }
            }
        };

        // Configure the reporting engine to show inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);

        // Output the result status.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item class without a Description property to trigger an error.
public class Item
{
    public string Name { get; set; } = "";
}
