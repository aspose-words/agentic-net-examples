using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingConditionalDefault
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
            // Register code page provider (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Item 1", Status = "Active" },
                    new Item { Name = "Item 2", Status = "Pending" },
                    new Item { Name = "Item 3", Status = "Other" }
                }
            };

            // Create template document programmatically
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("LINQ Reporting - Conditional Default Example");
            builder.Writeln();

            // Begin foreach loop over Items
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Name: <<[item.Name]>> - Status: ");

            // Conditional blocks for known statuses
            builder.Writeln("<<if [item.Status == \"Active\"]>>Active<</if>>");
            builder.Writeln("<<if [item.Status == \"Pending\"]>>Pending<</if>>");

            // Default block when no previous condition matches
            builder.Writeln("<<if [item.Status != \"Active\" && item.Status != \"Pending\"]>>Unknown<</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            doc.Save(templatePath);

            // Load the template for report generation
            var templateDoc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine
            var engine = new ReportingEngine();
            engine.BuildReport(templateDoc, model, "model");

            // Save the generated report
            var reportPath = "Report.docx";
            templateDoc.Save(reportPath);
        }
    }
}
