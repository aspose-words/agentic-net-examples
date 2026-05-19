using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a data band that iterates over the Items collection.
        // Inside the band we refer to members without the object reference (contextual access).
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Index: <<[Index]>>, Name: <<[Name]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" }
            }
        };

        // Build the report using the model as the root data source named "model".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item displayed inside the data band.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
