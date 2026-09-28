using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the template with LINQ Reporting tags.
        builder.Writeln("Report with Inline Error Messages");
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[model.Age]>>");
        // This tag references a non‑existent property and will trigger an inline error.
        builder.Writeln("Invalid field: <<[model.Missing]>>");

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Name = "John Doe",
            Age = 30
        };

        // Configure the reporting engine to insert inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);

        // Output the result status.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
