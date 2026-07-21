using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Header data model
    public class HeaderInfo
    {
        public string Title { get; set; } = "Monthly Sales Report";
        public string Date { get; set; } = DateTime.Now.ToString("MMMM yyyy");
    }

    // Body data model
    public class BodyInfo
    {
        public List<Item> Items { get; set; } = new()
        {
            new Item { Name = "Product A", Quantity = 120 },
            new Item { Name = "Product B", Quantity = 85 },
            new Item { Name = "Product C", Quantity = 47 }
        };
    }

    // Footer data model
    public class FooterInfo
    {
        public string Note { get; set; } = "Confidential – For internal use only";
    }

    // Item used in the body collection
    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    // Wrapper model that contains separate sections
    public class ReportModel
    {
        public HeaderInfo Header { get; set; } = new();
        public BodyInfo Body { get; set; } = new();
        public FooterInfo Footer { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Header section
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("<<[model.Header.Title]>>");
            builder.Writeln("Date: <<[model.Header.Date]>>");

            // Body section (main document)
            builder.MoveToSection(0);
            builder.Writeln("Report Body");
            builder.Writeln("<<foreach [item in model.Body.Items]>>");
            builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");

            // Footer section
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln("Footer: <<[model.Footer.Note]>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Prepare the data source
            ReportModel model = new ReportModel();

            // Configure the reporting engine
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;

            // Build the report using the wrapper object and root name "model"
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 3. Save the generated report
            // -------------------------------------------------
            reportDoc.Save(outputPath);
        }
    }
}
