using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for the Table class

namespace AsposeWordsLinqReportingImageDemo
{
    // Simple product model with a name and an image file path.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public string ImagePath { get; set; } = string.Empty;
    }

    // Root model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Folder for temporary files.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);

            // Create a tiny PNG image (red dot) from a Base64 string.
            string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9WvVhVQAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            string imagePath = Path.Combine(workDir, "RedDot.png");
            File.WriteAllBytes(imagePath, imageBytes);

            // Prepare sample data.
            var model = new ReportModel
            {
                Products = new List<Product>
                {
                    new Product { Name = "Product A", ImagePath = imagePath },
                    new Product { Name = "Product B", ImagePath = imagePath },
                    new Product { Name = "Product C", ImagePath = imagePath }
                }
            };

            // -----------------------------------------------------------------
            // Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(workDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin a foreach block over the Products collection.
            builder.Writeln("<<foreach [p in Products]>>");

            // Build a simple two‑column table: name | image.
            Table table = builder.StartTable();

            // First column – product name.
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");

            // Second column – image inside a textbox (required for image tags).
            builder.InsertCell();
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("<<image [p.ImagePath] -fitSize>>");

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // End the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // The root object name in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the final document.
            string outputPath = Path.Combine(workDir, "Report.docx");
            reportDoc.Save(outputPath);
        }
    }
}
