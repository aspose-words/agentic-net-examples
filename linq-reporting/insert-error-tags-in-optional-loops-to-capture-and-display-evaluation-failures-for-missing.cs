using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("Age: <<[item.Age]>>");

        // Optional loop for item details.
        builder.Writeln("<<if [item.Details != null]>>");
        builder.Writeln("Details:");
        builder.Writeln("<<foreach [detail in item.Details]>>");
        builder.Writeln(" - Description: <<[detail.Description]>>");
        // This property does not exist and will cause an evaluation error.
        builder.Writeln(" - Missing: <<[detail.Missing]>>");
        // Capture any errors that occurred in this block.
        builder.Writeln("<<error>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</if>>");

        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new()
            {
                new Item
                {
                    Name = "Alice",
                    Age = 30,
                    Details = new()
                    {
                        new Detail { Description = "Detail A1" },
                        new Detail { Description = "Detail A2" }
                    }
                },
                new Item
                {
                    Name = "Bob",
                    Age = null, // Age missing.
                    Details = null // No details.
                }
            }
        };

        // Build the report with inline error messages.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated. Success: {success}");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int? Age { get; set; }
    public List<Detail>? Details { get; set; }
}

public class Detail
{
    public string Description { get; set; } = string.Empty;
}
