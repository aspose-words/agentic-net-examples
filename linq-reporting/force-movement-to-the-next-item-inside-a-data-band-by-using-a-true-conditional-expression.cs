using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Id = 1, Name = "Alpha" },
                new Item { Id = 2, Name = "Beta" },
                new Item { Id = 3, Name = "Gamma" }
            }
        };

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Report Start");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Id]>> - <<[item.Name]>>");
        // Force movement to the next item using a true conditional expression.
        builder.Writeln("<<if [true]>> <<next>> <</if>>");
        // This line will be skipped for each item because of the <<next>> above.
        builder.Writeln("This line should NOT appear in the output.");
        builder.Writeln("<</foreach>>");
        builder.Writeln("Report End");
        builder.Document.Save(templatePath);

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "report.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
}
