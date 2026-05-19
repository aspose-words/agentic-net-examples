using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingImageInTable
{
    // Data model for a product.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public string ImagePath { get; set; } = string.Empty;
    }

    // Wrapper class required by the ReportingEngine (anonymous types are not allowed).
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare output folder.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(workDir);

            // Create a tiny PNG image from a Base64 string.
            string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9W6c2V8AAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            string imageFile = Path.Combine(workDir, "sample.png");
            File.WriteAllBytes(imageFile, pngBytes);

            // Sample data source.
            List<Product> products = new()
            {
                new Product { Name = "Product A", ImagePath = imageFile },
                new Product { Name = "Product B", ImagePath = imageFile }
            };

            // -----------------------------------------------------------------
            // Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin a foreach block over the Products collection.
            builder.Writeln("<<foreach [p in Products]>>");

            // Create a table with a header row.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Image");
            builder.EndRow();

            // Data row: product name.
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");

            // Data row: image inside a textbox shape.
            builder.InsertCell();
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("<<image [p.ImagePath] -fitSize>>");

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // End the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            string templatePath = Path.Combine(workDir, "template.docx");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;

            // Wrap the data source in a public class.
            var model = new ReportModel { Products = products };
            engine.BuildReport(reportDoc, model);

            // Save the generated report.
            string reportPath = Path.Combine(workDir, "report.docx");
            reportDoc.Save(reportPath);
        }
    }
}
