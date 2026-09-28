using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // A 1x1 pixel PNG image encoded in Base64.
    private static readonly byte[] SampleImageData = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=");

    // Returns a fresh stream containing the sample image.
    public Stream Image
    {
        get
        {
            // MemoryStream is positioned at the beginning by default.
            return new MemoryStream(SampleImageData);
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox that will hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag that expects a Stream from the model.
        builder.Write("<<image [model.Image] -fitSize>>");

        // Save the template (optional, just for inspection).
        doc.Save("Template.docx");

        // Prepare the data model.
        ReportModel model = new ReportModel();

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
