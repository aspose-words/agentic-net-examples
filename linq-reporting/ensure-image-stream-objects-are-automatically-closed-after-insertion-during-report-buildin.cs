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
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6V8AAAAASUVORK5CYII=");

        // Create the data model that supplies a Stream for the image tag.
        ReportModel model = new ReportModel(pngBytes);

        // -----------------------------------------------------------------
        // Step 1: Build the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // LINQ Reporting image tag that uses the Stream property.
        builder.Write("<<image [model.ImageStream] -fitSize>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // BuildReport will consume the Stream and close it automatically.
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        report.Save(reportPath);

        // -----------------------------------------------------------------
        // Step 3: Verify that the Stream has been closed by the engine.
        // -----------------------------------------------------------------
        bool isClosed;
        try
        {
            // Attempt to read from the stream; an ObjectDisposedException means it was closed.
            int _ = model.ImageStream.ReadByte();
            isClosed = false;
        }
        catch (ObjectDisposedException)
        {
            isClosed = true;
        }
        catch (Exception)
        {
            // Any other exception also indicates the stream is not usable.
            isClosed = true;
        }

        Console.WriteLine(isClosed
            ? "The image stream was automatically closed after insertion."
            : "The image stream is still open (unexpected).");
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // The Stream that provides image data to the <<image>> tag.
    public Stream ImageStream { get; set; }

    // Initialise the Stream with the supplied image bytes.
    public ReportModel(byte[] imageBytes)
    {
        ImageStream = new MemoryStream(imageBytes);
    }
}
