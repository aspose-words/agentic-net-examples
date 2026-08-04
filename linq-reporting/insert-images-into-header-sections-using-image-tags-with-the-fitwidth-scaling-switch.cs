using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

namespace AsposeWordsLinqReportingHeaderImage
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Path to the image that will be inserted into the header.
        public string ImagePath { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Folder for all generated files.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(workDir);

            // -----------------------------------------------------------------
            // 1. Create a sample image file (a tiny red square PNG).
            // -----------------------------------------------------------------
            string imageFile = Path.Combine(workDir, "SampleImage.png");
            CreateSamplePng(imageFile);

            // -----------------------------------------------------------------
            // 2. Build the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            string templateFile = Path.Combine(workDir, "Template.docx");
            CreateTemplate(templateFile);

            // -----------------------------------------------------------------
            // 3. Prepare the data model.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                ImagePath = imageFile
            };

            // -----------------------------------------------------------------
            // 4. Load the template and run the reporting engine.
            // -----------------------------------------------------------------
            Document templateDoc = new Document(templateFile);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(templateDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            string reportFile = Path.Combine(workDir, "Report.docx");
            templateDoc.Save(reportFile);
        }

        // Creates a minimal PNG image (red 100x100) and writes it to the specified path.
        private static void CreateSamplePng(string filePath)
        {
            // PNG binary for a 1x1 red pixel, then repeat to make a larger image.
            // This avoids using System.Drawing.
            byte[] pngData = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAIAAAD/gAIDAAAABnRSTlMAAAAAAABupkeAAAA" +
                "JcEhZcwAADsMAAA7DAcdvqGQAAABWSURBVHja7cEBDQAAAMKg909tDjegAAAAAAAAAAAAAAAA" +
                "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPwG" +
                "AAGXAAEJAAAAAElFTkSuQmCC");
            // Write the same data to create a visible image.
            File.WriteAllBytes(filePath, pngData);
        }

        // Generates a Word document that contains a header with an image tag.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move cursor to the primary header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert a textbox that will host the image tag.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
            // Position the cursor inside the textbox.
            builder.MoveTo(textBox.FirstParagraph);
            // Write the LINQ Reporting image tag with the fitWidth switch.
            builder.Write("<<image [model.ImagePath] -fitWidth>>");

            // Return to the main body and add a simple paragraph.
            builder.MoveToDocumentEnd();
            builder.Writeln("Report generated with an image in the header.");

            // Save the template.
            doc.Save(filePath);
        }
    }
}
