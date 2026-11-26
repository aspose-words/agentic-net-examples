using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple PNG image as a byte array (1x1 pixel, transparent).
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X0ZcAAAAASUVORK5CYII=");

        // Create a MemoryStream that will be used as the image source.
        MemoryStream imageStream = new MemoryStream(pngBytes);
        // Ensure the stream is positioned at the beginning.
        imageStream.Position = 0;

        // Build the data model.
        ReportModel model = new()
        {
            ImageStream = imageStream
        };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Insert a textbox to host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // Write the image tag using a Stream source and the -fitSizeLim switch.
        builder.Write("<<image [model.ImageStream] -fitSizeLim>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the report.
        // -----------------------------------------------------------------
        Document reportDoc = new(templatePath);

        // Reset the stream before the engine reads it.
        model.ImageStream.Position = 0;

        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // The image source stream. Must be reset before each use.
    public Stream ImageStream { get; set; } = Stream.Null;
}
