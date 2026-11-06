using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output documents
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a template document with an intentional runtime error
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a simple heading
        builder.Writeln("Report with Inline Error Messages");

        // Correct foreach loop
        builder.Writeln("<<foreach [item in Items]>>");
        // The expression <<[item.Unknown]>> does not exist in the data model and will trigger an inline error message
        builder.Writeln("Item: <<[item.Name]>> <<[item.Unknown]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Prepare sample data model
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Alpha" },
                new Item { Name = "Beta" },
                new Item { Name = "Gamma" }
            }
        };

        // -----------------------------------------------------------------
        // Step 3: Load the template and build the report with inline errors
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Enable inline error messages (default template will be used)
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report; success will be false due to the runtime error
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the resulting document
        reportDoc.Save(outputPath);

        // Output the success flag (optional, not required for the report)
        Console.WriteLine($"Report build success: {success}");
    }
}

// ---------------------------------------------------------------------
// Data model classes
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
}
