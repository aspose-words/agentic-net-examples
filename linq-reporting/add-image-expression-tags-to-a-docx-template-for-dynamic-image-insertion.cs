using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingImageExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for Aspose.Words).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare a folder for generated files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // Create a sample PNG image (1x1 pixel, red) and save it.
            string imagePath = Path.Combine(outputDir, "sample.png");
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=");
            File.WriteAllBytes(imagePath, pngBytes);

            // Define the data model.
            ReportModel model = new()
            {
                ImagePath = imagePath
            };

            // Create the template document with an image tag inside a textbox.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            // Insert a textbox shape to host the image tag.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("<<image [model.ImagePath] -fitSize>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            Document reportDoc = new(templatePath);

            // Build the report using LINQ Reporting Engine.
            ReportingEngine engine = new();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report.
            string reportPath = Path.Combine(outputDir, "Report.docx");
            reportDoc.Save(reportPath);
        }
    }

    // Public data model class.
    public class ReportModel
    {
        // Path to the image file to be inserted.
        public string ImagePath { get; set; } = string.Empty;
    }
}
