using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingImageExample
{
    // Root data model for the report.
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    // Individual product with image information.
    public class Product
    {
        public string Name { get; set; } = "";
        public string ImageUri { get; set; } = "";
        // The image data (byte array) that will be used by the LINQ Reporting engine.
        public byte[] ImageData { get; set; } = Array.Empty<byte>();
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the working directories exist.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a sample image file.
            string sampleImagePath = Path.Combine(outputDir, "sample.png");
            CreateSampleImage(sampleImagePath);

            // 2. Prepare sample data.
            ReportModel model = new()
            {
                Products = new List<Product>
                {
                    new Product { Name = "Product A", ImageUri = sampleImagePath },
                    new Product { Name = "Product B", ImageUri = sampleImagePath }
                }
            };

            // 3. Convert image URIs to byte arrays.
            foreach (var product in model.Products)
            {
                if (File.Exists(product.ImageUri))
                {
                    product.ImageData = File.ReadAllBytes(product.ImageUri);
                }
                else
                {
                    // Fallback to an empty image (1x1 pixel PNG) if the file is missing.
                    product.ImageData = CreateBlankPng();
                }
            }

            // 4. Create the LINQ Reporting template programmatically.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            CreateTemplate(templatePath);

            // 5. Load the template and build the report.
            Document doc = new(templatePath);
            ReportingEngine engine = new();
            engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            string reportPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(reportPath);
        }

        // Generates a simple PNG image for demonstration purposes.
        private static void CreateSampleImage(string path)
        {
            // Minimal PNG byte array (1x1 pixel).
            byte[] pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
                0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
                0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                0x42,0x60,0x82
            };
            File.WriteAllBytes(path, pngBytes);
        }

        // Creates a 1x1 pixel transparent PNG as a fallback image.
        private static byte[] CreateBlankPng()
        {
            return new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
                0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
                0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                0x42,0x60,0x82
            };
        }

        // Builds a Word document containing LINQ Reporting tags.
        private static void CreateTemplate(string path)
        {
            Document doc = new();
            DocumentBuilder builder = new(doc);

            builder.Writeln("Products Report");
            builder.Writeln("<<foreach [p in Products]>>");

            // Table header.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Image");
            builder.EndRow();

            // Data row.
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");
            builder.InsertCell();

            // Insert a textbox to host the image tag.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 120, 120);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("<<image [p.ImageData] -fitSize>>");

            builder.EndRow();
            builder.EndTable();

            builder.Writeln("<</foreach>>");

            doc.Save(path);
        }
    }
}
