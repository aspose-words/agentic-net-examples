using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample PNG image (1x1 pixel, transparent) encoded as Base64.
    private static readonly string Base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";

    // Byte array containing the image data.
    public byte[] ImageBytes { get; set; } = Convert.FromBase64String(Base64Png);

    // Returns a fresh stream for each access, as required by the image tag.
    public Stream ImageStream => new MemoryStream(ImageBytes);
}

public class Program
{
    public static void Main()
    {
        // Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox to host the image.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // Write the image tag with the -fitSizeLim switch.
        builder.Write("<<image [model.ImageStream] -fitSizeLim>>");

        // Save the template to a memory stream and reload it (required by the workflow).
        using (var templateStream = new MemoryStream())
        {
            template.Save(templateStream, SaveFormat.Docx);
            templateStream.Position = 0;
            Document loadedTemplate = new Document(templateStream);

            // Prepare the data model.
            var model = new ReportModel();

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            loadedTemplate.Save("output.docx");
        }
    }
}
