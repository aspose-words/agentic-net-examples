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

        // Template: iterate over Items, print Index, then force move to next item,
        // so the Name field will be skipped for every record.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Index: <<[item.Index]>>");
        // The true condition always evaluates to true, causing the <<next>> tag to execute.
        builder.Writeln("<<if [true]>><<next>><</if>>");
        builder.Writeln("Name: <<[item.Name]>>");
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

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root data model referenced by the template (named "model").
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class used inside the data band.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
