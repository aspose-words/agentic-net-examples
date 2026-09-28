using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a data band that iterates over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Output the item name.
        builder.Writeln("Item: <<[item.Name]>>");
        // Force movement to the next item using a true conditional expression.
        builder.Writeln("<<if [true]>>");
        builder.Writeln("<<next>>");
        builder.Writeln("<</if>>");
        // This line will be skipped because of the <<next>> tag.
        builder.Writeln("This line will be skipped.");
        // End of the data band.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template for report generation.
        var templateDoc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alpha", Value = 1 },
                new Item { Name = "Beta", Value = 2 },
                new Item { Name = "Gamma", Value = 3 }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        templateDoc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item class used in the data band.
public class Item
{
    public string Name { get; set; } = "";
    public int Value { get; set; }
}
