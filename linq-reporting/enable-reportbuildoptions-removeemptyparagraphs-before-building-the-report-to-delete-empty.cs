using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data. The Description property is empty, which will result in an empty paragraph.
        var model = new ReportModel
        {
            Title = "Sample Report",
            Description = "",
            Items = new List<Item>
            {
                new Item { Name = "Item 1", Value = "Value 1" },
                new Item { Name = "Item 2", Value = "" } // Empty value to demonstrate removal.
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title paragraph.
        builder.Writeln("<<[model.Title]>>");

        // Description paragraph – will become empty after processing.
        builder.Writeln("<<[model.Description]>>");

        // Simple foreach loop to list items.
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("<<[item.Name]>>: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine to remove empty paragraphs.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("ReportOutput.docx");
    }

    // Data model used by the template.
    public class ReportModel
    {
        public string Title { get; set; } = "";
        public string Description { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public string Value { get; set; } = "";
    }
}
