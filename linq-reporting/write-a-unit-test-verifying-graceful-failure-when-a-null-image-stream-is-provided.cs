using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ImageStreamNullTest
{
    // Model class used by the LINQ Reporting engine.
    public class ReportModel
    {
        // This property intentionally returns null to simulate a missing image stream.
        public Stream? ImageStream { get; set; } = null;
    }

    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox to host the image tag (required by the engine).
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting tag that expects a Stream (or other supported types) for the image.
        builder.Write("<<image [model.ImageStream]>>");

        // Prepare the data model with a null Stream.
        ReportModel model = new ReportModel();

        // Configure the reporting engine to emit inline error messages instead of throwing.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report. The method returns false when parsing fails and InlineErrorMessages is set.
        bool success = engine.BuildReport(template, model, "model");

        // Verify that the engine reported a failure (graceful handling of the null stream).
        if (!success)
        {
            Console.WriteLine("Test passed: BuildReport returned false as expected for null image stream.");
        }
        else
        {
            Console.WriteLine("Test failed: BuildReport succeeded unexpectedly.");
        }

        // Optionally, save the resulting document to inspect the inline error message.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportWithError.docx");
        template.Save(outputPath);
        Console.WriteLine($"Report saved to: {outputPath}");
    }
}
