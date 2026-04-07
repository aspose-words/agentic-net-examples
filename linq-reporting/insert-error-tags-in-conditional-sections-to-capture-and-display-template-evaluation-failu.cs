using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Write a normal field that exists.
        builder.Writeln("Name: <<[item.Name]>>");

        // Conditional block that references a missing property (item.MissingProp).
        // This will cause a template evaluation error.
        builder.Writeln("<<if [item.MissingProp]>>");
        builder.Writeln("This text will never appear because the condition fails.");
        builder.Writeln("<</if>>");

        // Insert the <<error>> tag to capture any errors that occurred in the block above.
        builder.Writeln("<<error>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple" },
                new Item { Name = "Banana" },
                new Item { Name = "Cherry" }
            }
        };

        // Configure the reporting engine to inline error messages.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report. The root object name must match the tag prefix used in the template.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // The success flag is meaningful only when InlineErrorMessages is enabled.
        Console.WriteLine($"Report build success: {success}");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Initialize the collection to avoid nullable warnings.
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    // Note: No MissingProp property is defined intentionally to trigger an error.
}
