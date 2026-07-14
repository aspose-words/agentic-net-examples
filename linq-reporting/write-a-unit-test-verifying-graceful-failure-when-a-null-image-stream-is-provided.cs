using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Model
{
    // The image stream can be null to simulate a missing image.
    public Stream ImageStream { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox to host the image tag (required by LINQ Reporting).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting tag that expects a Stream, byte[], Image or a path.
        builder.Write("<<image [model.ImageStream]>>");

        // Prepare the data model with a null image stream.
        var model = new Model { ImageStream = null };

        // Configure the reporting engine to inline error messages instead of throwing.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The method returns false when an error occurs and InlineErrorMessages is enabled.
        bool success = engine.BuildReport(doc, model, "model");

        // Output the result. Expecting graceful failure (success == false).
        Console.WriteLine($"Report build success: {success}");

        // Save the resulting document to inspect the inline error message (optional).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportWithError.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
