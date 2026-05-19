using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingHeaderImage
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for legacy encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths.
            string workDir = Directory.GetCurrentDirectory();
            string templatePath = Path.Combine(workDir, "HeaderImageTemplate.docx");
            string outputPath = Path.Combine(workDir, "HeaderImageReport.docx");
            string imagePath = Path.Combine(workDir, "SampleImage.png");

            // Create a minimal PNG image (1x1 pixel) and write it to disk.
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=");
            File.WriteAllBytes(imagePath, pngBytes);

            // -----------------------------------------------------------------
            // 1. Build the template document with a header containing an image tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Move cursor to the primary header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert a textbox that will host the image tag.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 100);
            builder.MoveTo(textBox.FirstParagraph);

            // Insert the LINQ Reporting image tag with the fitWidth switch.
            builder.Write("<<image [model.ImagePath] -fitWidth>>");

            // Add a simple paragraph in the body.
            builder.MoveToDocumentEnd();
            builder.Writeln("Report body content goes here.");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report using LINQ Reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Prepare the data model.
            ReportModel model = new ReportModel
            {
                ImagePath = imagePath
            };

            // Create and run the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report.
            reportDoc.Save(outputPath);
        }
    }

    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Path to the image that will be inserted into the header.
        public string ImagePath { get; set; } = string.Empty;
    }
}
