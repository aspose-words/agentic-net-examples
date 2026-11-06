using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with a tag that references a nonexistent member.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Start");
        builder.Writeln("<<[nonexistent]>>"); // This member does not exist in the model.
        builder.Writeln("End");

        // Simple model with no members.
        var model = new Model();

        // Configure the reporting engine to handle missing members gracefully.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages; // Missing members will be treated as null.

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the document (optional).
        doc.Save("output.docx");

        // Verify that the missing member produced no output.
        string resultText = doc.GetText()
                               .Replace("\r", string.Empty)
                               .Replace("\n", string.Empty)
                               .Trim();

        Console.WriteLine($"Report built successfully: {success}");
        Console.WriteLine($"Resulting document text (without line breaks): '{resultText}'");
        // Expected output: "StartEnd"
    }

    // Empty model class – no members defined.
    public class Model
    {
    }
}
