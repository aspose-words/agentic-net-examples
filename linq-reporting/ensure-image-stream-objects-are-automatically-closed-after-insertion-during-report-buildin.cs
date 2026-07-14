using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // The stream will be closed automatically by the reporting engine after insertion.
    public Stream ImageStream { get; set; }

    public ReportModel()
    {
        // A 1x1 pixel PNG image (base64 encoded).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AApEB/6V6nVQAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        ImageStream = new MemoryStream(pngBytes);
    }
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "template.docx";
        const string reportPath = "report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document with an image tag inside a textbox.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting image tag referencing the stream property.
        builder.Write("<<image [model.ImageStream] -fitSize>>");

        // Save the template (required before building the report).
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 2. Prepare the data model containing the image stream.
        // ---------------------------------------------------------------
        ReportModel model = new ReportModel();

        // ---------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // ---------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name must match the tag prefix ("model").
        engine.BuildReport(templateDoc, model, "model");

        // Save the populated report.
        templateDoc.Save(reportPath);

        // ---------------------------------------------------------------
        // 4. Verify that the image stream has been closed automatically.
        // ---------------------------------------------------------------
        bool streamClosed = false;
        try
        {
            // Accessing a disposed stream throws ObjectDisposedException.
            long _ = model.ImageStream.Position;
        }
        catch (ObjectDisposedException)
        {
            streamClosed = true;
        }

        // Output the verification result.
        Console.WriteLine($"Image stream closed after report build: {streamClosed}");
    }
}
