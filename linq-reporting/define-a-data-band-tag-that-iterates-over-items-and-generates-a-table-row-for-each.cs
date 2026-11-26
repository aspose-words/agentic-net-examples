using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingExample
{
    // Root data model
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Item model used in the foreach loop
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Sample data
            var model = new ReportModel
            {
                Items = new()
                {
                    new Item { Index = 1, Name = "Apple" },
                    new Item { Index = 2, Name = "Banana" },
                    new Item { Index = 3, Name = "Cherry" }
                }
            };

            // -----------------------------------------------------------------
            // Create the template document programmatically
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Begin the foreach data band
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table that will be repeated for each item
            Table table = builder.StartTable();

            // Header row (appears for each item in this simple example)
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row – values are taken from the current item
            builder.InsertCell();
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.EndRow();

            // End the table and the foreach block
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and build the report
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
