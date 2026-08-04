using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data item with a Name property.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
    }

    // Root model that will be passed to the reporting engine.
    public class ReportModel
    {
        // Collection of items; initialized to avoid nullable warnings.
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a Word template that contains a LINQ Reporting tag.
            //    The tag uses ElementAt to fetch the fifth element (index 4).
            // -----------------------------------------------------------------
            const string templateFile = "Template.docx";

            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // The expression will be evaluated by the ReportingEngine.
            builder.Writeln("Fifth item: <<[model.Items.ElementAt(4).Name]>>");

            // Save the template to disk.
            templateDoc.Save(templateFile);

            // -----------------------------------------------------------------
            // 2. Load the template back (required by the workflow rules).
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templateFile);

            // -----------------------------------------------------------------
            // 3. Prepare sample data with at least five items.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new() { Name = "Item 1" },
                    new() { Name = "Item 2" },
                    new() { Name = "Item 3" },
                    new() { Name = "Item 4" },
                    new() { Name = "Item 5" }, // This is the fifth item.
                    new() { Name = "Item 6" }
                }
            };

            // -----------------------------------------------------------------
            // 4. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputFile = "Report.docx";
            reportDoc.Save(outputFile);
        }
    }
}
