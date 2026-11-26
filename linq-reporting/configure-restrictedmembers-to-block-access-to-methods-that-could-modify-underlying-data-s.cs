using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Collection of items to be displayed in the report.
    public List<string> Items { get; set; } = new();

    // Method that adds an item to the collection and returns the new count.
    public int AddItem(string item)
    {
        Items.Add(item);
        return Items.Count;
    }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Display the number of items before attempting to add a new one.
        builder.Writeln("Items before: <<[model.Items.Count]>>");

        // Attempt to add a new item via a method call in the template.
        // This will invoke ReportModel.AddItem during report generation.
        builder.Writeln("Attempt to add item: <<[model.AddItem(\"NewItem\")]>>");

        // Display the number of items after the method call.
        builder.Writeln("Items after: <<[model.Items.Count]>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for reporting.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel();
        model.Items.Add("Item1");
        model.Items.Add("Item2");

        // -----------------------------------------------------------------
        // Configure the reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Enable inline error messages (useful for debugging).
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // NOTE: In newer versions of Aspose.Words the RestrictedMembers
        // collection may not be available. If it exists, you could block
        // methods like AddItem here. For compatibility we omit this step.

        // -----------------------------------------------------------------
        // Build the report.
        // -----------------------------------------------------------------
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "output.docx");
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Output saved to: {outputPath}");
    }
}
