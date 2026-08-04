using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Text = "Approved", Color = "Green" },
                new Item { Text = "Pending", Color = "Orange" },
                new Item { Text = "Rejected", Color = "Red" }
            }
        };

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [item in Items]>>");
        // Apply dynamic text color using the textColor tag.
        builder.Writeln("<<textColor [item.Color]>>[item.Text]<</textColor>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template and build the report.
        var loadedTemplate = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save("Report.docx");
    }
}

// Data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item containing text and a color expression.
public class Item
{
    public string Text { get; set; } = string.Empty;
    public string Color { get; set; } = string.Empty;
}
