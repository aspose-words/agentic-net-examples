using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingImageFitSize
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Path to the image that will be inserted into the document.
        public string ImagePath { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a working directory for all generated files.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(workDir);

            // -----------------------------------------------------------------
            // 1. Create a sample image file (a small red square) from a Base64 string.
            // -----------------------------------------------------------------
            string imagePath = Path.Combine(workDir, "sample.png");
            const string base64Png =
                "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFklEQVQoU2NkYGD4z0AEYBxVSFIAABcAAf+XKc8AAAAASUVORK5CYII=";
            File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

            // -----------------------------------------------------------------
            // 2. Build the template document programmatically.
            //    The image tag is placed inside a textbox and uses the -fitSize switch
            //    to stretch the image and fill the containing paragraph.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(workDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a textbox that will act as the image container.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
            // Move the cursor into the textbox's first paragraph.
            builder.MoveTo(textBox.FirstParagraph);
            // LINQ Reporting image tag with -fitSize scaling.
            builder.Write("<<image [model.ImagePath] -fitSize>>");

            // Save the template for later loading.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and populate it using the ReportingEngine.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Create the data model instance and point it to the sample image.
            ReportModel model = new ReportModel { ImagePath = imagePath };

            // Build the report. The root object name in the template is "model".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the final document.
            // -----------------------------------------------------------------
            string resultPath = Path.Combine(workDir, "result.docx");
            reportDoc.Save(resultPath);
        }
    }
}
