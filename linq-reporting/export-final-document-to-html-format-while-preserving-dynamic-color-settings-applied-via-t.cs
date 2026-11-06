using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Item
{
    public string Name { get; set; } = "";
    public string ColorName { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create the LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report of Items with Dynamic Colors:");
        builder.Writeln("<<foreach [item in Items]>>");
        // Apply dynamic text color using the textColor tag.
        builder.Writeln("<<textColor [item.ColorName]>><<[item.Name]>> <</textColor>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document doc = new Document(templatePath);

        // Sample data.
        ReportModel model = new()
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", ColorName = "Red" },
                new Item { Name = "Banana", ColorName = "Gold" },
                new Item { Name = "Grapes", ColorName = "Purple" }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Export to HTML preserving colors.
        string htmlPath = Path.Combine(outputDir, "report.html");
        doc.Save(htmlPath, SaveFormat.Html);
    }
}
