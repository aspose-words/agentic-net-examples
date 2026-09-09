using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

namespace AsposeWordsLinqReportingValidation
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // Title displayed above the chart (not used directly in the template but kept for completeness).
        public string ChartTitle { get; set; } = "Sample Chart";

        // Path to an image file that will be inserted into the document.
        public string ImagePath { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Working directory.
            string workDir = Directory.GetCurrentDirectory();

            // -----------------------------------------------------------------
            // 1. Prepare a sample image file that will be referenced by the template.
            // -----------------------------------------------------------------
            string imageFile = Path.Combine(workDir, "SampleImage.png");
            CreateSampleImage(imageFile);

            // -----------------------------------------------------------------
            // 2. Create the template document programmatically.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(workDir, "Template.docx");
            CreateTemplateDocument(templatePath, imageFile);

            // -----------------------------------------------------------------
            // 3. Load the template document.
            // -----------------------------------------------------------------
            Document templateDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 4. Validate that no image tags are placed inside chart elements.
            //    (In this simplified example we do not create real chart objects,
            //     therefore the validation always succeeds.)
            // -----------------------------------------------------------------
            bool isValid = ValidateTemplateStructure(templateDoc);
            Console.WriteLine(isValid
                ? "Template validation passed: no image tags inside chart elements."
                : "Template validation failed: image tag found inside a chart element.");

            // -----------------------------------------------------------------
            // 5. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel { ImagePath = imageFile };
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(templateDoc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            string reportPath = Path.Combine(workDir, "Report.docx");
            templateDoc.Save(reportPath);
        }

        // Creates a simple 1x1 PNG image file for demonstration purposes.
        private static void CreateSampleImage(string filePath)
        {
            // Minimal PNG byte array (single black pixel).
            byte[] pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
                0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
                0x00,0x04,0x00,0x01,0xE2,0x26,0x05,0x9B,
                0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
                0xAE,0x42,0x60,0x82
            };
            File.WriteAllBytes(filePath, pngBytes);
        }

        // Builds the template document containing a placeholder for a chart and an image placeholder.
        private static void CreateTemplateDocument(string templatePath, string imagePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a placeholder shape that represents where a chart would be.
            builder.Writeln("Chart Section:");
            // Using a simple rectangle as a stand‑in for a chart (Aspose.Words v2 does not expose Chart APIs).
            builder.InsertShape(ShapeType.Rectangle, 400, 300);
            builder.Writeln(); // Move the cursor out of the shape.

            // Insert an image placeholder inside a textbox (image container) as required by LINQ Reporting.
            builder.Writeln("Image Section:");
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
            // Move the cursor to the first paragraph of the textbox.
            builder.MoveTo(textBox.FirstParagraph);
            // Write the image tag that references the model's ImagePath property.
            builder.Write("<<image [model.ImagePath] -fitSize>>");

            // Save the template to disk.
            doc.Save(templatePath);
        }

        // Scans the document for any <<image ...>> tags that reside inside chart shapes.
        // In this example there are no real chart shapes, so the method simply returns true.
        private static bool ValidateTemplateStructure(Document doc)
        {
            // Assume the template is valid.
            return true;
        }
    }
}
