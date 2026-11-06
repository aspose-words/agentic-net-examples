using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Value { get; set; }
    public string Description { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // Create the template document with LINQ Reporting tags
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Dynamic Text Color Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<textColor [item.Value > 100 ? \"Red\" : \"Green\"]>>Item <<[item.Description]>>: <<[item.Value]>> <</textColor>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template (simulating a separate load step)
        var doc = new Document(templatePath);

        // Prepare sample data
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Value = 45, Description = "Low" },
                new Item { Value = 120, Description = "High" },
                new Item { Value = 100, Description = "Medium" }
            }
        };

        // Build the report using the LINQ Reporting engine
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(outputPath);
    }
}
