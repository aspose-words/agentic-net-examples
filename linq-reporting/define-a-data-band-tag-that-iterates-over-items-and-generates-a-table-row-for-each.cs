using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Needed for the Table class

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;

    public Item(int index, string name)
    {
        Index = index;
        Name = name;
    }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin the data band (foreach) that will iterate over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table inside the foreach block.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row – each iteration will fill these cells.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        // Close the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel();
        model.Items.Add(new Item(1, "Apple"));
        model.Items.Add(new Item(2, "Banana"));
        model.Items.Add(new Item(3, "Cherry"));

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the final report.
        loadedTemplate.Save(reportPath);
    }
}
