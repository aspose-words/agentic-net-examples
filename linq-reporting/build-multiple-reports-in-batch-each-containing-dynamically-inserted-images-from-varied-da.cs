using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingBatchImages
{
    // Data model for each report.
    public class ReportData
    {
        public string Title { get; set; } = string.Empty;
        public string ImagePath { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words operations).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare folders.
            string workDir = Directory.GetCurrentDirectory();
            string assetsDir = Path.Combine(workDir, "Assets");
            string outputDir = Path.Combine(workDir, "Output");
            Directory.CreateDirectory(assetsDir);
            Directory.CreateDirectory(outputDir);

            // Create sample images.
            CreateSampleImage(Path.Combine(assetsDir, "Red.png"),   "#FF0000");
            CreateSampleImage(Path.Combine(assetsDir, "Green.png"), "#00FF00");

            // Build the template document.
            string templatePath = Path.Combine(workDir, "Template.docx");
            BuildTemplate(templatePath);

            // Prepare data for batch reports.
            var reports = new List<ReportData>
            {
                new() { Title = "Red Square Report",   ImagePath = Path.Combine(assetsDir, "Red.png")   },
                new() { Title = "Green Square Report", ImagePath = Path.Combine(assetsDir, "Green.png") }
            };

            // Generate each report.
            int index = 1;
            foreach (var data in reports)
            {
                // Load fresh template for each iteration.
                var doc = new Document(templatePath);

                // Build the report using LINQ Reporting Engine.
                var engine = new ReportingEngine();
                engine.Options = ReportBuildOptions.None;
                engine.BuildReport(doc, data, "model");

                // Save the generated report.
                string outPath = Path.Combine(outputDir, $"Report_{index}.docx");
                doc.Save(outPath);
                index++;
            }
        }

        // Creates a simple PNG image of a solid color using a Base64 string.
        private static void CreateSampleImage(string filePath, string hexColor)
        {
            // 1x1 pixel PNG base64 for the requested color.
            // The PNG data is generated once and written to the file.
            string base64 = hexColor switch
            {
                "#FF0000" => "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=",
                "#00FF00" => "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQADgwF/6VYV2QAAAABJRU5ErkJggg==",
                _ => throw new ArgumentException("Unsupported color")
            };
            byte[] imageBytes = Convert.FromBase64String(base64);
            File.WriteAllBytes(filePath, imageBytes);
        }

        // Builds the template document containing a title placeholder and an image placeholder inside a textbox.
        private static void BuildTemplate(string templatePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Title placeholder.
            builder.Writeln("Report Title: <<[model.Title]>>");

            // Insert a textbox to host the image tag.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("<<image [model.ImagePath] -fitSize>>");

            // Save the template.
            doc.Save(templatePath);
        }
    }
}
