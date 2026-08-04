using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string json = @"{
            ""Items"": [
                { ""Index"": 1, ""Name"": ""Apple"",  ""Quantity"": 5 },
                { ""Index"": 2, ""Name"": ""Banana"", ""Quantity"": 3 },
                { ""Index"": 3, ""Name"": ""Cherry"", ""Quantity"": 12 }
            ]
        }";

        // Deserialize JSON into model.
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(json) ?? new ReportModel();

        // Create template document.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Items Report");
        builder.Writeln();

        // Header table (static).
        Table headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln(); // space between tables.

        // Data rows table (repeated via foreach).
        builder.Writeln("<<foreach [item in Items]>>");
        Table dataTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report using the model as the root named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}
