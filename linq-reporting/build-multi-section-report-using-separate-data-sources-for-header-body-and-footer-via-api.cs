using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingMultiSection
{
    // Header data model
    public class HeaderModel
    {
        public string Title { get; set; } = "Quarterly Sales Report";
        public string Date { get; set; } = DateTime.Now.ToString("MMMM dd, yyyy");
    }

    // Body item model
    public class BodyItem
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    // Body data model containing a collection of items
    public class BodyModel
    {
        public List<BodyItem> Items { get; set; } = new();
    }

    // Footer data model
    public class FooterModel
    {
        public int Total { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the final report
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Section 1 – Header
            builder.Writeln("<<[header.Title]>>");
            builder.Writeln("Date: <<[header.Date]>>");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Section 2 – Body (repeating items)
            builder.Writeln("Products:");
            builder.Writeln("<<foreach [item in body.Items]>>");
            builder.Writeln(" - <<[item.Name]>> : <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Section 3 – Footer
            builder.Writeln("Total items: <<[footer.Total]>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare data sources.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // Header data
            var header = new HeaderModel();

            // Body data with sample items
            var body = new BodyModel
            {
                Items =
                {
                    new BodyItem { Name = "Apples", Quantity = 120 },
                    new BodyItem { Name = "Bananas", Quantity = 85 },
                    new BodyItem { Name = "Cherries", Quantity = 60 }
                }
            };

            // Footer data (calculate total quantity)
            var footer = new FooterModel
            {
                Total = 0
            };
            foreach (var item in body.Items)
                footer.Total += item.Quantity;

            // -----------------------------------------------------------------
            // 3. Build the report using multiple data sources.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine
            {
                // Example option – remove empty paragraphs after processing
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // BuildReport overload that accepts multiple data sources and their names
            bool success = engine.BuildReport(
                doc,
                new object[] { header, body, footer },
                new[] { "header", "body", "footer" });

            // Optional: check success flag (relevant when InlineErrorMessages option is used)
            if (!success)
                Console.WriteLine("Report generation encountered errors.");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
