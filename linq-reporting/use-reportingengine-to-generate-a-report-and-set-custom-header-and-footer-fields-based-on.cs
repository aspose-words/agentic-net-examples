using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class ReportModel
{
    public string Title { get; set; } = "Sample Report";
    public string HeaderText { get; set; } = "Report Header";
    public string FooterText { get; set; } = "Report Footer";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public string Value { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "data.json";
        var sampleData = new ReportModel
        {
            Title = "Quarterly Sales Report",
            HeaderText = "Company XYZ - Confidential",
            FooterText = "Generated on " + DateTime.Now.ToString("yyyy-MM-dd"),
            Items = new List<Item>
            {
                new Item { Name = "Product A", Value = "$1,200" },
                new Item { Name = "Product B", Value = "$3,450" },
                new Item { Name = "Product C", Value = "$2,780" }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load JSON into model.
        string json = File.ReadAllText(jsonPath);
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(json)!;

        // Create template document.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header with LINQ Reporting tag.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("<<[model.HeaderText]>>");

        // Footer with LINQ Reporting tag.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("<<[model.FooterText]>>");

        // Return to main body.
        builder.MoveToDocumentEnd();

        // Title.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Table of items using foreach.
        builder.Writeln("<<foreach [item in Items]>>");
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Value]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine("output", "report.docx");
        Directory.CreateDirectory("output");
        reportDoc.Save(outputPath);
    }
}
