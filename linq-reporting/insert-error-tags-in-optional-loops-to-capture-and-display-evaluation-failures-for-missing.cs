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
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel
        {
            Title = "Sample LINQ Reporting",
            Items = new List<Item>
            {
                new Item { Name = "Alice" },
                new Item { Name = "Bob" }
                // Note: Item does NOT have an Age property – this will cause an evaluation error.
            }
        };

        // Create a template document programmatically
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report Title: <<[model.Title]>>");
        builder.Writeln();
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("- Name: <<[item.Name]>>");
        // Attempt to access a missing property 'Age' and capture the error with <<error>>
        builder.Writeln("- Age: <<[item.Age]>> <<error>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection)
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // Load the template for reporting
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Output simple status (no interactive input)
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Report saved to: {Path.GetFullPath(outputPath)}");
    }
}

// Data model classes
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    // No Age property – used to demonstrate missing data handling.
}
