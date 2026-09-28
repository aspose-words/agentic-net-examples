using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingImageInFooter
{
    public class ReportModel
    {
        // Path to the image that will be inserted into the footer.
        public string ImagePath { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required in .NET Core).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare working directories.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(workDir);

            // Create a simple 1x1 pixel PNG image from a Base64 string.
            string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YV6ZV8AAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            string imagePath = Path.Combine(workDir, "sample.png");
            File.WriteAllBytes(imagePath, imageBytes);

            // Create the template document programmatically.
            string templatePath = Path.Combine(workDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a primary footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // Insert a textbox shape that will host the image tag.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 50);
            // Move the cursor inside the textbox.
            builder.MoveTo(textBox.FirstParagraph);
            // Write the image tag with -fitSize switch.
            builder.Write("<<image [model.ImagePath] -fitSize>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            Document reportDoc = new Document(templatePath);

            // Prepare the model with the image path.
            ReportModel model = new ReportModel
            {
                ImagePath = imagePath
            };

            // Build the report using LINQ Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            bool success = engine.BuildReport(reportDoc, model, "model");

            // Save the generated document.
            string outputPath = Path.Combine(workDir, "ReportWithFooterImage.docx");
            reportDoc.Save(outputPath);

            // Optionally, indicate success (no console interaction required).
            Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
        }
    }
}
