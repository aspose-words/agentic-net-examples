using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingColorExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Create sample data
            var model = new ReportModel
            {
                Items = new()
                {
                    new Item { Name = "Task 1", Status = "Open" },
                    new Item { Name = "Task 2", Status = "Closed" },
                    new Item { Name = "Task 3", Status = "InProgress" }
                }
            };

            // Create the template document programmatically
            var templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Task Status Report");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Task: <<[item.Name]>> - ");
            builder.Writeln("<<textColor [item.Status == \"Open\" ? \"Green\" : (item.Status == \"Closed\" ? \"Red\" : \"Orange\")]>>");
            builder.Writeln("<<[item.Status]>>");
            builder.Writeln("<</textColor>>");
            builder.Writeln("<</foreach>>");

            doc.Save(templatePath);

            // Load the template and build the report
            var templateDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(templateDoc, model, "model");

            // Save the generated report
            var outputPath = "ReportOutput.docx";
            templateDoc.Save(outputPath);
        }
    }
}
