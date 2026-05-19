using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

namespace AsposeWordsLinqReportingExample
{
    // Data model containing the image data as a byte array.
    public class ReportModel
    {
        // Holds the image bytes.
        public byte[] ImageData { get; set; }

        public ReportModel()
        {
            // A tiny 1x1 pixel PNG image encoded in Base64.
            const string base64Png =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9Yc2cVIAAAAASUVORK5CYII=";
            ImageData = Convert.FromBase64String(base64Png);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -------------------------------------------------
            // Step 1: Create the template document programmatically.
            // -------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a textbox that will host the image tag.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
            // Move the cursor inside the textbox.
            builder.MoveTo(textBox.FirstParagraph);
            // Write the LINQ Reporting image tag referencing the model's ImageData property.
            builder.Write("<<image [model.ImageData] -fitSize>>");

            // Save the template.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            template.Save(templatePath);

            // -------------------------------------------------
            // Step 2: Load the template for report generation.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // -------------------------------------------------
            // Step 3: Build the report using the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // Step 4: Save the generated report.
            // -------------------------------------------------
            string reportPath = Path.Combine(outputDir, "Report.docx");
            reportDoc.Save(reportPath);
        }
    }
}
