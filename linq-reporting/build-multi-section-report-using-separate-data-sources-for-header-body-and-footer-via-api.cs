using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace MultiSectionReport
{
    // Data model for the header section.
    public class HeaderModel
    {
        public string Title { get; set; } = "Monthly Sales Report";
    }

    // Data model for a single item in the body section.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    // Data model for the body section.
    public class BodyModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Data model for the footer section.
    public class FooterModel
    {
        public int PageNumber { get; set; } = 1;
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            string templatePath = "MultiSectionTemplate.docx";

            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Header section.
            builder.Writeln("Header:");
            builder.Writeln("<<[header.Title]>>");
            builder.Writeln();

            // Body section with a foreach loop over body.Items.
            builder.Writeln("Body:");
            builder.Writeln("<<foreach [item in body.Items]>>");
            builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln();

            // Footer section.
            builder.Writeln("Footer:");
            builder.Writeln("Page <<[footer.PageNumber]>>");
            builder.Writeln();

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare realistic sample data for each section.
            // -----------------------------------------------------------------
            HeaderModel header = new HeaderModel
            {
                Title = "Quarterly Revenue Summary"
            };

            BodyModel body = new BodyModel();
            body.Items.Add(new Item { Index = 1, Name = "North America" });
            body.Items.Add(new Item { Index = 2, Name = "Europe" });
            body.Items.Add(new Item { Index = 3, Name = "Asia-Pacific" });

            FooterModel footer = new FooterModel
            {
                PageNumber = 5
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report using multiple data sources.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.Options = ReportBuildOptions.None;

            // Pass the three data sources together with their names.
            object[] dataSources = { header, body, footer };
            string[] dataSourceNames = { "header", "body", "footer" };

            engine.BuildReport(reportDoc, dataSources, dataSourceNames);

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = "MultiSectionReport.docx";
            reportDoc.Save(outputPath);

            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
        }
    }
}
