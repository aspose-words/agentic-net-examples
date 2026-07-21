using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Template header.
        builder.Writeln("Report generated with Aspose.Words LINQ Reporting");
        builder.Writeln();

        // Optional loop over the collection "Items".
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        // This expression refers to a non‑existent member and will cause an evaluation error.
        builder.Writeln("MissingProperty: <<[item.MissingProperty]>>");
        // The <<error>> tag will display the inline error message for the above failure.
        builder.Writeln("<<error>>");
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine to inline error messages.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report using the model as the root data source named "model".
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alice", Age = 30 },
                new Item { Name = "Bob" } // No Age, but Age is not used in the template.
            }
        };

        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated document.
        const string outputPath = "ReportWithErrors.docx";
        doc.Save(outputPath);

        // Output the success flag (will be true because InlineErrorMessages is enabled).
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model used inside the foreach loop.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int? Age { get; set; }
}
