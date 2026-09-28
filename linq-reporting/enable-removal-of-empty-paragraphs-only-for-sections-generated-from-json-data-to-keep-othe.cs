using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

namespace LinqReportingRemoveEmptyParagraphs
{
    public class Section
    {
        public string Title { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
    }

    public class ReportModel
    {
        public List<Section> Sections { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare template document
            string templatePath = "Template.docx";
            CreateTemplate(templatePath);

            // Prepare JSON data and deserialize to model
            string json = @"
            {
                ""Sections"": [
                    { ""Title"": ""Section 1"", ""Content"": ""This is the first section content."" },
                    { ""Title"": ""Section 2"", ""Content"": """" }
                ]
            }";
            ReportModel model = JsonConvert.DeserializeObject<ReportModel>(json)!;

            // Load template
            Document doc = new Document(templatePath);

            // Configure reporting engine to remove empty paragraphs generated from data
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build report
            engine.BuildReport(doc, model, "model");

            // Save output
            string outputPath = "Output.docx";
            doc.Save(outputPath);
        }

        private static void CreateTemplate(string path)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Static header
            builder.Writeln("=== Report Header ===");
            // Static empty paragraph that should remain untouched
            builder.Writeln("");

            // Begin LINQ Reporting section generated from JSON
            builder.Writeln("<<foreach [sec in Sections]>>");
            builder.Writeln("<<[sec.Title]>>");
            builder.Writeln("<<[sec.Content]>>");
            builder.Writeln("<</foreach>>");

            // Static footer
            builder.Writeln("=== Report Footer ===");

            doc.Save(path);
        }
    }
}
