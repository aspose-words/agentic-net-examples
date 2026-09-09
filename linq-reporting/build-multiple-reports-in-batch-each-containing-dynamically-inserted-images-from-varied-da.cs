using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingBatch
{
    // Root data model for a single report.
    public class ReportModel
    {
        // Title displayed in the report.
        public string Title { get; set; } = "Untitled Report";

        // Collection of products to list.
        public List<Product> Products { get; set; } = new();
    }

    // Simple product class containing a name and a path to an image file.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public string ImagePath { get; set; } = string.Empty;
    }

    public class Program
    {
        // Entry point.
        public static void Main()
        {
            // Ensure the working directories exist.
            string baseDir = Directory.GetCurrentDirectory();
            string imagesDir = Path.Combine(baseDir, "Images");
            string outputDir = Path.Combine(baseDir, "Output");
            Directory.CreateDirectory(imagesDir);
            Directory.CreateDirectory(outputDir);

            // Create two sample image files (tiny PNGs) that will be used in the reports.
            CreateSampleImage(Path.Combine(imagesDir, "apple.png"),   ApplePngBase64);
            CreateSampleImage(Path.Combine(imagesDir, "banana.png"),  BananaPngBase64);

            // Build two distinct report models.
            var reports = new List<ReportModel>
            {
                new()
                {
                    Title = "Fruit Report – Set A",
                    Products = new()
                    {
                        new() { Name = "Apple",  ImagePath = Path.Combine(imagesDir, "apple.png") },
                        new() { Name = "Banana", ImagePath = Path.Combine(imagesDir, "banana.png") }
                    }
                },
                new()
                {
                    Title = "Fruit Report – Set B",
                    Products = new()
                    {
                        new() { Name = "Banana", ImagePath = Path.Combine(imagesDir, "banana.png") },
                        new() { Name = "Apple",  ImagePath = Path.Combine(imagesDir, "apple.png") }
                    }
                }
            };

            // Create the LINQ Reporting template once.
            string templatePath = Path.Combine(baseDir, "ReportTemplate.docx");
            CreateTemplate(templatePath);

            // Process each report model in batch.
            int index = 1;
            foreach (var model in reports)
            {
                // Load the template for each iteration.
                var doc = new Document(templatePath);

                // Build the report using the LINQ Reporting engine.
                var engine = new ReportingEngine();
                engine.Options = ReportBuildOptions.None; // default options
                bool success = engine.BuildReport(doc, model, "model");

                // Save the generated report.
                string outputPath = Path.Combine(outputDir, $"Report_{index}.docx");
                doc.Save(outputPath);
                Console.WriteLine($"Report {index} generated: {outputPath} (Success = {success})");
                index++;
            }
        }

        // Creates a simple Word template containing LINQ Reporting tags.
        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Title placeholder.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln();

            // Begin foreach over Products collection.
            builder.Writeln("<<foreach [p in Products]>>");

            // Table with two columns: product name and image.
            var table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Product");
            builder.InsertCell();
            builder.Writeln("Image");
            builder.EndRow();

            // Data row (repeated for each product).
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");
            builder.InsertCell();

            // Insert a textbox to host the image tag (required by the engine).
            var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 150, 100);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("<<image [p.ImagePath] -fitSize>>");

            // End of the data row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // End foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(filePath);
        }

        // Writes a PNG file from a Base64 string.
        private static void CreateSampleImage(string filePath, string base64Data)
        {
            byte[] bytes = Convert.FromBase64String(base64Data);
            File.WriteAllBytes(filePath, bytes);
        }

        // Very small 1x1 red PNG (Apple placeholder).
        private const string ApplePngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AApEB" +
            "gZ6XK6cAAAAASUVORK5CYII=";

        // Very small 1x1 yellow PNG (Banana placeholder).
        private const string BananaPngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8AFAgEB" +
            "A6cK6VQAAAAASUVORK5CYII=";
    }
}
