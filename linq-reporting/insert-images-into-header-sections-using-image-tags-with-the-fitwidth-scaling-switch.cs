using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using System.Text;

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
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Prepare a sample image file that will be referenced from the model.
            // -----------------------------------------------------------------
            string imageFile = "sample.png";
            CreateSamplePng(imageFile);

            // -----------------------------------------------------------------
            // 2. Build the template document programmatically.
            //    The template contains an image tag inside a textbox placed in the header.
            // -----------------------------------------------------------------
            string templateFile = "HeaderImageTemplate.docx";
            CreateTemplate(templateFile, imageFile);

            // -----------------------------------------------------------------
            // 3. Load the template and populate it using the ReportingEngine.
            // -----------------------------------------------------------------
            Document doc = new Document(templateFile);

            // Create the model and set the image path.
            ReportModel model = new ReportModel { ImagePath = Path.GetFullPath(imageFile) };

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated document.
            // -----------------------------------------------------------------
            string outputFile = "HeaderImageReport.docx";
            doc.Save(outputFile);
        }

        // Creates a very small PNG file (1x1 pixel, red) using a hard‑coded byte array.
        private static void CreateSamplePng(string filePath)
        {
            // PNG header + IHDR + IDAT (red pixel) + IEND.
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
            File.WriteAllBytes(filePath, pngBytes);
        }

        // Generates the template document with a header that contains an image tag.
        private static void CreateTemplate(string templatePath, string placeholderImagePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the primary header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert a textbox that will host the image tag.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 100);
            // Position the cursor inside the textbox.
            builder.MoveTo(textBox.FirstParagraph);
            // Write the LINQ Reporting image tag with the fitWidth switch.
            builder.Write("<<image [model.ImagePath] -fitWidth>>");

            // Return to the main document body.
            builder.MoveToDocumentEnd();

            // Save the template.
            doc.Save(templatePath);
        }
    }
}
