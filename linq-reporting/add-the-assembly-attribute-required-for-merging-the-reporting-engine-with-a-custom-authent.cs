using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model for the report.
    public class ReportModel
    {
        public string Title { get; set; } = "Sample Report";
        public List<Item> Items { get; set; } = new();
    }

    // Item class used in the collection.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank template document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Prepare sample data.
            ReportModel model = new ReportModel();
            model.Items.Add(new Item { Index = 1, Name = "Apple" });
            model.Items.Add(new Item { Index = 2, Name = "Banana" });
            model.Items.Add(new Item { Index = 3, Name = "Cherry" });

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
