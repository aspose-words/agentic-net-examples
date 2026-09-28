using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a foreach loop over Items and display the value and its square root.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Value: <<[item.Value]>>   Sqrt: <<[item.Sqrt]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item(4),
                new Item(9),
                new Item(16),
                new Item(25)
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model with a numeric field and a computed square‑root property.
public class Item
{
    public double Value { get; set; }

    // Computed property used in the template.
    public double Sqrt => Math.Sqrt(Value);

    public Item(double value)
    {
        Value = value;
    }
}
