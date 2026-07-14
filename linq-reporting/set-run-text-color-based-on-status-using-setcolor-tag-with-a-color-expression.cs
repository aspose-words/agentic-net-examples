using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of items to be displayed in the report.
        public List<Item> Items { get; set; } = new();
    }

    // Individual item with a status field.
    public class Item
    {
        public string Status { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank document and insert the LINQ Reporting template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");

            // Use the textColor tag. The color expression selects a color based on the item's status.
            // The expression returns a known color name string.
            builder.Writeln(
                "<<textColor [item.Status == \"Completed\" ? \"Green\" : " +
                "item.Status == \"Pending\" ? \"Orange\" : \"Red\"]>>" +
                "<<[item.Status]>>" +
                "<</textColor>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // 2. Prepare sample data.
            ReportModel model = new()
            {
                Items = new List<Item>
                {
                    new() { Status = "Completed" },
                    new() { Status = "Pending" },
                    new() { Status = "Failed" },
                    new() { Status = "Completed" }
                }
            };

            // 3. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 4. Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
