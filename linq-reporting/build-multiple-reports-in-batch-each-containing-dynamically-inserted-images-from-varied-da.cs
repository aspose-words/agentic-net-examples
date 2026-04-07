using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

namespace LinqReportingBatchImages
{
    // Root data model for a single report.
    public class ReportModel
    {
        public string Title { get; set; } = string.Empty;
        public string ImagePath { get; set; } = string.Empty;
    }

    // Simple item used to generate several reports.
    public class Item
    {
        public string Title { get; set; } = string.Empty;
        public string ImagePath { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare working folders.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            string imagesDir = Path.Combine(workDir, "Images");
            string outputDir = Path.Combine(workDir, "Output");
            Directory.CreateDirectory(imagesDir);
            Directory.CreateDirectory(outputDir);

            // Create two sample PNG images (red and green squares) from valid Base64 strings.
            // These Base64 strings represent 1x1 pixel PNG images of the respective colors.
            const string redPixelBase64 =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/lK5ZAAAAAElFTkSuQmCC";
            const string greenPixelBase64 =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8DAAQAD/6V9WQAAAABJRU5ErkJggg==";

            CreateSampleImage(Path.Combine(imagesDir, "RedSquare.png"), redPixelBase64);
            CreateSampleImage(Path.Combine(imagesDir, "GreenSquare.png"), greenPixelBase64);

            // Build a list of items, each pointing to a different image.
            List<Item> items = new()
            {
                new Item { Title = "Red Square Report", ImagePath = Path.Combine(imagesDir, "RedSquare.png") },
                new Item { Title = "Green Square Report", ImagePath = Path.Combine(imagesDir, "GreenSquare.png") }
            };

            // Create the template document programmatically.
            string templatePath = Path.Combine(workDir, "Template.docx");
            CreateTemplate(templatePath);

            // Process each item and generate a separate report.
            int index = 1;
            foreach (Item item in items)
            {
                // Load the template.
                Document reportDoc = new Document(templatePath);

                // Prepare the model for this iteration.
                ReportModel model = new()
                {
                    Title = item.Title,
                    ImagePath = item.ImagePath
                };

                // Build the report using LINQ Reporting Engine.
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(reportDoc, model, "model");

                // Save the generated report.
                string outPath = Path.Combine(outputDir, $"Report_{index}.docx");
                reportDoc.Save(outPath);
                index++;
            }
        }

        // Creates a PNG file from a Base64 string.
        private static void CreateSampleImage(string filePath, string base64)
        {
            byte[] bytes = Convert.FromBase64String(base64);
            File.WriteAllBytes(filePath, bytes);
        }

        // Generates a Word template containing a title placeholder and an image placeholder inside a textbox.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Title placeholder.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln();

            // Insert a textbox that will host the image.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
            builder.MoveTo(textBox.FirstParagraph);
            // Image tag with fitSize switch.
            builder.Write("<<image [model.ImagePath] -fitSize>>");

            // Save the template.
            doc.Save(filePath);
        }
    }
}
