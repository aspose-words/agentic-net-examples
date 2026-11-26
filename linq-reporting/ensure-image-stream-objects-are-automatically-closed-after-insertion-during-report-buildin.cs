using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // The image stream will be provided at runtime.
    public Stream ImageStream { get; set; } = Stream.Null;
}

public class Program
{
    public static void Main()
    {
        // Prepare a simple 1x1 PNG image as a base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Create the template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox that will hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag referencing the stream property.
        builder.Write("<<image [model.ImageStream] -fitSize>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare the model with the image stream.
        using (MemoryStream imageStream = new MemoryStream(imageBytes))
        {
            ReportModel model = new ReportModel
            {
                ImageStream = imageStream
            };

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");
        } // The using block ensures the stream is closed automatically.

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
