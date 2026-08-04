using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Newtonsoft.Json;
using System.Text;

namespace LinqReportingPdfExample
{
    // Root data model for the report.
    public class ReportModel
    {
        public string Title { get; set; } = "Sample LINQ Reporting";
        public string HtmlSnippet { get; set; } = "<b>Bold HTML snippet</b>";
        public string ImagePath { get; set; } = "sample.png";
        public List<Item> Items { get; set; } = new();
    }

    // Simple item used in a foreach loop.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        private const string TemplateFileName = "template.docx";
        private const string OutputPdfFileName = "report.pdf";

        public static void Main()
        {
            // Ensure the required image exists.
            CreateSampleImage("sample.png");

            // Build the DOCX template with LINQ Reporting tags.
            CreateTemplate();

            // Load the template document.
            Document doc = new Document(TemplateFileName);

            // Prepare sample data.
            ReportModel model = new ReportModel();
            model.Items.Add(new Item { Index = 1, Name = "Alpha" });
            model.Items.Add(new Item { Index = 2, Name = "Beta" });
            model.Items.Add(new Item { Index = 3, Name = "Gamma" });

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options required.
            engine.BuildReport(doc, model, "model");

            // Save the populated document as PDF.
            doc.Save(OutputPdfFileName, SaveFormat.Pdf);
        }

        // Creates a simple DOCX template containing various LINQ tags.
        private static void CreateTemplate()
        {
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Title tag.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln();

            // HTML snippet tag with -html switch.
            builder.Writeln("<<[model.HtmlSnippet] -html>>");
            builder.Writeln();

            // Foreach loop over Items collection.
            builder.Writeln("<<foreach [item in model.Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln();

            // Image inside a textbox.
            builder.Writeln("Image:");
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("<<image [model.ImagePath] -fitSize>>");

            // Save the template to disk.
            template.Save(TemplateFileName);
        }

        // Generates a 1x1 pixel PNG file from a Base64 string.
        private static void CreateSampleImage(string filePath)
        {
            // Base64 for a transparent 1x1 PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YV6ZV8AAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(filePath, imageBytes);
        }
    }
}
