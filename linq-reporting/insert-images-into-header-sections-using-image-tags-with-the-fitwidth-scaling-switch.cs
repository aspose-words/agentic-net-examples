using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingHeaderImage
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Path to the image that will be inserted into the header.
        public string ImagePath { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Folder for temporary files.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);

            // -----------------------------------------------------------------
            // 1. Create a sample PNG image (a single red pixel) and save it.
            // -----------------------------------------------------------------
            string imagePath = Path.Combine(workDir, "sample.png");
            // Base64 for a 1x1 red PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/6VYVAAAAAElFTkSuQmCC";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, imageBytes);

            // -----------------------------------------------------------------
            // 2. Build the template document programmatically.
            //    The header contains a textbox with an image tag that uses -fitWidth.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Move cursor to the primary header of the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert a textbox that will host the image tag.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
            // Move the builder inside the textbox's first paragraph.
            builder.MoveTo(textBox.FirstParagraph);
            // Write the LINQ Reporting image tag with the fitWidth switch.
            builder.Write("<<image [model.ImagePath] -fitWidth>>");

            // Return to the main body (optional, not strictly required).
            builder.MoveToDocumentEnd();

            // Save the template to disk (required before building the report).
            string templatePath = Path.Combine(workDir, "Template.docx");
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data model instance.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                ImagePath = imagePath
            };

            // -----------------------------------------------------------------
            // 4. Load the template and run the LINQ Reporting engine.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(workDir, "ReportWithHeaderImage.docx");
            doc.Save(outputPath);

            // The example finishes without waiting for user input.
        }
    }
}
