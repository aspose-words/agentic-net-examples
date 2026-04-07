using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSortingExample
{
    public class Item
    {
        public string Name { get; set; } = "";
        public int Value { get; set; }
    }

    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Sample data (unsorted)
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new() { Name = "Banana", Value = 3 },
                    new() { Name = "Apple", Value = 5 },
                    new() { Name = "Cherry", Value = 2 }
                }
            };

            // Create the template document
            const string templatePath = "Template.docx";
            CreateTemplate(templatePath);

            // Load the template
            var doc = new Document(templatePath);

            // Build the report – the template sorts the collection inline
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }

        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Items sorted by Name:");
            builder.Writeln("<<foreach [item in Items.OrderBy(item => item.Name)]>>");
            builder.Writeln(" - <<[item.Name]>> : <<[item.Value]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(filePath);
        }
    }
}
