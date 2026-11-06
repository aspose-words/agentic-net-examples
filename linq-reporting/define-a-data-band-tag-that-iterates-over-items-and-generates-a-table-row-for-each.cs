using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingExample
{
    // Sample data model
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a template document programmatically
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Write the foreach tag that iterates over Items
            builder.Writeln("<<foreach [item in Items]>>");

            // Start a table inside the foreach block
            Table table = builder.StartTable();

            // Header row (optional, will appear for each iteration if placed inside foreach)
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row bound to the current item
            builder.InsertCell();
            // Display item.Index
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            // Display item.Name
            builder.Writeln("<<[item.Name]>>");
            builder.EndRow();

            // End the table and foreach block
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template for report generation
            Document doc = new Document(templatePath);

            // Prepare sample data
            ReportModel model = new()
            {
                Items = new()
                {
                    new Item { Index = 1, Name = "Alpha" },
                    new Item { Index = 2, Name = "Beta" },
                    new Item { Index = 3, Name = "Gamma" }
                }
            };

            // Build the report using LINQ Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
