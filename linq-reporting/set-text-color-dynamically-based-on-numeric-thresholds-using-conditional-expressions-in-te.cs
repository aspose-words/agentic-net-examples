using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model for each item.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Score { get; set; }
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document and a builder to construct the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");

            // Use a textColor tag with a conditional expression to set the color based on the Score.
            // Red for Score < 50, Green for Score >= 80, otherwise Black.
            builder.Writeln(
                "<<textColor [item.Score < 50 ? \"Red\" : item.Score >= 80 ? \"Green\" : \"Black\"]>>" +
                "Item: <<[item.Name]>>, Score: <<[item.Score]>>" +
                " <</textColor>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Prepare sample data.
            ReportModel model = new ReportModel();
            model.Items.Add(new Item { Name = "Alice", Score = 45 });
            model.Items.Add(new Item { Name = "Bob", Score = 73 });
            model.Items.Add(new Item { Name = "Charlie", Score = 88 });

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
