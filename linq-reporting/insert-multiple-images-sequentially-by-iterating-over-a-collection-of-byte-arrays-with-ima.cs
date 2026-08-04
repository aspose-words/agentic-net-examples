using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Added to resolve the Table type

namespace AsposeWordsLinqReportingExample
{
    // Model class that holds a collection of images as byte arrays.
    public class ReportModel
    {
        // Initialize the list to avoid nullable warnings.
        public List<ImageItem> Images { get; set; } = new();

        // Helper method to populate the collection with sample images.
        public void LoadSampleImages()
        {
            // Two tiny PNG images encoded in Base64.
            const string base64PngRed = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9YcV4V8AAAAASUVORK5CYII=";
            const string base64PngGreen = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8AABgAD/6V6XwAAAABJRU5ErkJggg==";

            Images.Add(new ImageItem { Data = Convert.FromBase64String(base64PngRed) });
            Images.Add(new ImageItem { Data = Convert.FromBase64String(base64PngGreen) });
        }
    }

    // Simple wrapper for a single image byte array.
    public class ImageItem
    {
        public byte[] Data { get; set; } = Array.Empty<byte>();
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create the LINQ Reporting template programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Begin a foreach loop over the Images collection.
            builder.Writeln("<<foreach [img in Images]>>");

            // Create a table that will hold one image per row.
            Table table = builder.StartTable();
            builder.InsertCell();

            // Insert a textbox shape; the image tag must be inside the textbox.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Writeln("<<image [img.Data]>>");

            // Close the row and the table for each iteration.
            builder.EndRow();
            builder.EndTable();

            // End the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Step 2: Prepare the data model with sample images.
            ReportModel model = new();
            model.LoadSampleImages();

            // Step 3: Load the template and build the report.
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };
            engine.BuildReport(report, model, "model");

            // Step 4: Save the generated report.
            const string reportPath = "Report.docx";
            report.Save(reportPath);
        }
    }
}
