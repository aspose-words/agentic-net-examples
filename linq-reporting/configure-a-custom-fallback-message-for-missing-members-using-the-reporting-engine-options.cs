using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple template document with a missing member tag.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Existing property: <<[model.ExistingProperty]>>");
        builder.Writeln("Missing property: <<[model.MissingProperty]>>");

        // Prepare the data model (does not contain MissingProperty).
        ReportModel model = new()
        {
            ExistingProperty = "Sample value"
        };

        // Configure the reporting engine to show inline error messages.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.InlineErrorMessages;
        // Note: In older versions of Aspose.Words the custom fallback message property
        // may not be available. The inline error messages will indicate missing members.

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated document.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Output result information.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    public class ReportModel
    {
        public string ExistingProperty { get; set; } = string.Empty;
        // MissingProperty is intentionally omitted to trigger the fallback message.
    }
}
