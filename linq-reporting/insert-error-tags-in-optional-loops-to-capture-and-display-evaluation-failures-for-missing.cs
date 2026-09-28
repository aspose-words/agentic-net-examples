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
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a temporary folder for the example files
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        // Paths for template and result documents
        string templatePath = Path.Combine(workDir, "Template.docx");
        string resultPath = Path.Combine(workDir, "Report.docx");

        // -------------------------------------------------
        // Step 1: Build the LINQ Reporting template document
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("LINQ Reporting Example with Inline Error Messages");
        builder.Writeln();
        // Optional loop over Items collection
        builder.Writeln("<<foreach [item in Items]>>");
        // Normal field
        builder.Writeln("Name: <<[item.Name]>> <<error>>");
        // Intentional missing field to trigger an error
        builder.Writeln("Missing: <<[item.NonExisting]>> <<error>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Step 2: Load the template document for reporting
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // Step 3: Prepare sample data model
        // -------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alice" },
                new Item { Name = "Bob" },
                // This item will have a null Name to demonstrate handling of null values
                new Item { Name = null }
            }
        };

        // -------------------------------------------------
        // Step 4: Build the report with InlineErrorMessages option
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        bool success = engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // Step 5: Save the generated report
        // -------------------------------------------------
        doc.Save(resultPath);

        // Output simple status (no interactive prompts)
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report saved to: {resultPath}");
    }
}

// -------------------------------------------------
// Data model classes
// -------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    // Name may be null to illustrate missing data handling
    public string? Name { get; set; }
}
