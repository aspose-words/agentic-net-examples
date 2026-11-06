using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Square root report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Value: <<[item.Value]>>   Sqrt: <<[item.Sqrt]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new List<Item>
            {
                new() { Value = 4 },
                new() { Value = 9 },
                new() { Value = 16 },
                new() { Value = 25 }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public double Value { get; set; }

    // Helper property to expose the square root of Value.
    public double Sqrt => Math.Sqrt(Value);
}
