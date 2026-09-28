using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using System.Text;

namespace ImageInHeaderExample
{
    public class ReportModel
    {
        public string ImagePath { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare working directories
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a simple PNG image (1x1 pixel, red) and save it locally
            string imageFile = Path.Combine(outputDir, "sample.png");
            byte[] pngData = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=");
            File.WriteAllBytes(imageFile, pngData);

            // Create the template document with a header containing an image tag
            string templateFile = Path.Combine(outputDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Move to the primary header
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert a textbox shape to host the image tag
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
            builder.MoveTo(textBox.FirstParagraph);

            // Write the LINQ Reporting image tag with -fitWidth switch
            builder.Write("<<image [model.ImagePath] -fitWidth>>");

            // Return to the main body and add a placeholder paragraph
            builder.MoveToDocumentEnd();
            builder.Writeln("Report body content goes here.");

            // Save the template
            templateDoc.Save(templateFile);

            // Load the template for report generation
            Document reportDoc = new Document(templateFile);

            // Prepare the data model
            ReportModel model = new ReportModel
            {
                ImagePath = imageFile
            };

            // Build the report using LINQ Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final document
            string outputFile = Path.Combine(outputDir, "output.docx");
            reportDoc.Save(outputFile);
        }
    }
}
