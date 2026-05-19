using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingMultiSection
{
    // Header data source
    public class Header
    {
        public string Title { get; set; } = string.Empty;
        public string Date { get; set; } = string.Empty;
    }

    // Body data source
    public class Body
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }

    // Footer data source
    public class Footer
    {
        public string Note { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document with three sections (header, body, footer)
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // ----- Header section -----
            builder.Writeln("=== Header Section ===");
            builder.Writeln("Title: <<[header.Title]>>");
            builder.Writeln("Date: <<[header.Date]>>");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // ----- Body section -----
            builder.Writeln("=== Body Section ===");
            builder.Writeln("<<foreach [item in body.Items]>>");

            // Table header (appears once)
            var table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Item");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.EndRow();

            // Table row for each item
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Quantity]>>");
            builder.EndRow();
            builder.EndTable();

            builder.Writeln("<</foreach>>");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // ----- Footer section -----
            builder.Writeln("=== Footer Section ===");
            builder.Writeln("<<[footer.Note]>>");

            // Save the template to disk
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare separate data sources for header, body, and footer
            // -----------------------------------------------------------------
            var header = new Header
            {
                Title = "Quarterly Sales Report",
                Date = DateTime.Now.ToString("yyyy-MM-dd")
            };

            var body = new Body
            {
                Items = new List<Item>
                {
                    new Item { Name = "Product A", Quantity = 120 },
                    new Item { Name = "Product B", Quantity = 85 },
                    new Item { Name = "Product C", Quantity = 47 }
                }
            };

            var footer = new Footer
            {
                Note = "Confidential – For internal use only"
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine with three data sources
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;

            // BuildReport overload that accepts multiple data sources
            bool success = engine.BuildReport(
                doc,
                new object[] { header, body, footer },
                new[] { "header", "body", "footer" });

            // Save the generated report
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Simple console output to indicate completion
            Console.WriteLine(success
                ? $"Report generated successfully: {outputPath}"
                : "Report generation failed.");
        }
    }
}
