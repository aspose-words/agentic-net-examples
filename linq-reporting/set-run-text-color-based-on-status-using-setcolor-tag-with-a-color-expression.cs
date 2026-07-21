using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create template document
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Report: Items with Status Colors");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln(
            "<<textColor [item.Status == \"Success\" ? \"Green\" : item.Status == \"Warning\" ? \"Orange\" : \"Red\"]>>" +
            "Item: <<[item.Name]>> | Status: <<[item.Status]>>" +
            "<</textColor>>");
        builder.Writeln("<</foreach>>");

        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // Prepare data model
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Task A", Status = "Success" },
                new Item { Name = "Task B", Status = "Warning" },
                new Item { Name = "Task C", Status = "Error" },
                new Item { Name = "Task D", Status = "Success" }
            }
        };

        // Load template and build report
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);

        Console.WriteLine($"Report generated at: {reportPath}");
    }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
}
