using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("Report with Inline Error Messages");
        builder.Writeln();

        // Optional foreach loop over Items. Inside the loop we reference a missing member (Name)
        // and place an <<error>> tag to capture any evaluation failures.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item Id: <<[item.Id]>>");
        builder.Writeln("Item Name: <<[item.Name]>>"); // 'Name' does not exist on Item.
        builder.Writeln("<<error>>"); // Will display the error message for the missing member.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document for reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model with missing members.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Id = 1 }, // Missing Name.
                new Item { Id = 2 }  // Missing Name.
            }
        };

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine to use InlineErrorMessages.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The bool indicates success when InlineErrorMessages is enabled.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);

        // Output the result to the console.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{reportPath}'.");
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    // The collection referenced by the foreach loop in the template.
    public List<Item> Items { get; set; } = new();
}

// Item intentionally lacks a 'Name' property to trigger missing member errors.
public class Item
{
    public int Id { get; set; }
}
