using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        var model = new ReportModel
        {
            Title = "Sample Report",
            Items = new()
            {
                new Item { Name = "Item A", Quantity = 3 },
                new Item { Name = "Item B", Quantity = 7 },
                new Item { Name = "Item C", Quantity = 1 }
            }
        };

        // Serialize the model to JSON.
        string jsonString = JsonConvert.SerializeObject(model, Formatting.Indented);

        // Create a template document with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<[root.Title]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in root.Items]>>");
        builder.Writeln(" - <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a memory stream.
        using var templateStream = new MemoryStream();
        templateDoc.Save(templateStream, SaveFormat.Docx);
        templateStream.Position = 0;

        // Load the template from the stream for reporting.
        var reportDoc = new Document(templateStream);

        // Create a JsonDataSource from the JSON string via a memory stream.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonString));
        var jsonDataSource = new JsonDataSource(jsonStream);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "root");

        // Save the generated report as RTF.
        reportDoc.Save("ReportOutput.rtf", SaveFormat.Rtf);
    }
}

// Data model classes.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
