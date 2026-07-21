using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to construct the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin the data band – the whole table will be repeated for each item in Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Build the table inside the foreach block.
        Table table = builder.StartTable();

        // Header row – two cells that will be merged horizontally.
        builder.InsertCell();
        builder.Write("<<cellMerge>>Group A");
        builder.InsertCell();
        builder.Write("<<cellMerge>>Group A");
        builder.EndRow();

        // Data row – will be generated for each item.
        builder.InsertCell();
        builder.Write("<<[item.Col1]>>");
        builder.InsertCell();
        builder.Write("<<[item.Col2]>>");
        builder.EndRow();

        // Finish the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Col1 = "A1", Col2 = "B1" },
                new Item { Col1 = "A2", Col2 = "B2" },
                new Item { Col1 = "A3", Col2 = "B3" }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("MergedTableReport.docx");
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class representing a row in the table.
public class Item
{
    public string Col1 { get; set; } = string.Empty;
    public string Col2 { get; set; } = string.Empty;
}
