using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

#nullable enable

public class ReportModel
{
    // The image stream is intentionally left null to test graceful failure.
    public Stream? ImageStream { get; set; } = null;
}

public class Program
{
    public static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting tag that tries to render an image from a null stream.
        builder.Write("<<image [model.ImageStream]>>");

        // Prepare the data model with a null image stream.
        ReportModel model = new ReportModel();

        // Configure the reporting engine to inline error messages instead of throwing.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The method should return false indicating a parsing error,
        // but it must not throw an exception.
        bool success = engine.BuildReport(doc, model, "model");

        // Output the result of the build operation.
        Console.WriteLine($"BuildReport succeeded: {success}");

        // Save the resulting document for manual inspection if needed.
        doc.Save("Output.docx");
    }
}
