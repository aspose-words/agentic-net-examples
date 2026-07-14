using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingConditionalDefault
{
    // Data model classes
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public string Status { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report
            string templatePath = "template.docx";
            string reportPath = "report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Report:");
            // Start a foreach loop over Items
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item: <<[item.Name]>> - Status: ");

            // Specific conditions
            builder.Writeln("<<if [item.Status == \"New\"]>>New<</if>>");
            builder.Writeln("<<if [item.Status == \"InProgress\"]>>In Progress<</if>>");
            builder.Writeln("<<if [item.Status == \"Completed\"]>>Completed<</if>>");

            // Default block when none of the above conditions are true
            builder.Writeln("<<if [item.Status != \"New\" && item.Status != \"InProgress\" && item.Status != \"Completed\"]>>Other<</if>>");

            // End of foreach
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for report generation
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare sample data
            // -------------------------------------------------
            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Task 1", Status = "New" },
                    new Item { Name = "Task 2", Status = "InProgress" },
                    new Item { Name = "Task 3", Status = "Completed" },
                    new Item { Name = "Task 4", Status = "OnHold" } // Will trigger the default block
                }
            };

            // -------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 5. Save the generated report
            // -------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
