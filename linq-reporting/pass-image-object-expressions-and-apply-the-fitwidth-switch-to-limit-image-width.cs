using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Image data (byte array) that will be inserted into the report.
        public byte[] Image { get; set; } = null!;

        // Optional title to demonstrate additional fields.
        public string Title { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Working directory for temporary files.
            string workDir = Directory.GetCurrentDirectory();

            // -----------------------------------------------------------------
            // 1. Prepare a sample PNG image (a small red pixel) as a byte array.
            // -----------------------------------------------------------------
            // Base64-encoded PNG of a 1x1 red pixel.
            const string base64RedPixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z8BQDwAF/AL+XK5ZVQAAAABJRU5ErkJggg==";
            byte[] imageBytes = Convert.FromBase64String(base64RedPixel);

            // -----------------------------------------------------------------
            // 2. Prepare the data model instance with the image bytes.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Image = imageBytes,
                Title = "Sample Image Report"
            };

            // -----------------------------------------------------------------
            // 3. Build the template document programmatically.
            //    The image tag must be placed inside a textbox container.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(workDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Optional title paragraph.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln(); // Empty line.

            // Insert a textbox that will hold the image.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
            // Move the cursor into the textbox's first paragraph.
            builder.MoveTo(textBox.FirstParagraph);
            // Image tag with -fitWidth switch to limit the width.
            builder.Write("<<image [model.Image] -fitWidth>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 4. Load the template and generate the report using ReportingEngine.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report; the root object name is "model".
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(workDir, "ReportWithImage.docx");
            reportDoc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine("Report generated successfully at:");
            Console.WriteLine(outputPath);
        }
    }
}
