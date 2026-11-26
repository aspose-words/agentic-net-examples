using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with a header row.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Status");
        builder.InsertCell();
        builder.Writeln("Description");
        builder.EndRow();

        // Data row with colored status text.
        builder.InsertCell();
        builder.Writeln(
            "<<textColor [item.Status == \"Completed\" ? \"Green\" : item.Status == \"Pending\" ? \"Orange\" : \"Red\"]>><<[item.Status]>> <</textColor>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Description]>>");
        builder.EndRow();

        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new List<Item>
            {
                new Item { Status = "Completed", Description = "Task finished successfully." },
                new Item { Status = "Pending", Description = "Task is awaiting execution." },
                new Item { Status = "Failed", Description = "Task encountered an error." }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model used in the template.
public class Item
{
    public string Status { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}
