using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; } = 0;
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" }
            }
        };

        // Create a blank document and insert LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Data band that iterates over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Output the current item's name.
        builder.Writeln("Item: <<[item.Name]>>");
        // Force movement to the next item using a true condition.
        builder.Writeln("<<if [true]>> <<next>> <</if>>");
        // End of the data band.
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
