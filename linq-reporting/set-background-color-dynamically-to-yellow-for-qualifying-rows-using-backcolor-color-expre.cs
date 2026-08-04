using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words
        EncodingProvider provider = CodePagesEncodingProvider.Instance;
        Encoding.RegisterProvider(provider);

        // Create template document
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin foreach over Items
        builder.Writeln("<<foreach [item in Items]>>");

        // Start table
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.IsImportant]>>" +
            "<<backColor [\"Yellow\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.IsImportant == false]>>" +
            "<<[item.Name]>> <</if>>");

        builder.InsertCell();
        builder.Writeln(
            "<<if [item.IsImportant]>>" +
            "<<backColor [\"Yellow\"]>><<[item.Value]>> <</backColor>><</if>>" +
            "<<if [item.IsImportant == false]>>" +
            "<<[item.Value]>> <</if>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach
        builder.Writeln("<</foreach>>");

        // Save template
        templateDoc.Save(templatePath);

        // Load template for reporting
        Document doc = new Document(templatePath);

        // Prepare data model
        ReportModel model = new()
        {
            Items = new List<Item>
            {
                new Item { Name = "Alpha", Value = 10, IsImportant = true },
                new Item { Name = "Beta", Value = 20, IsImportant = false },
                new Item { Name = "Gamma", Value = 30, IsImportant = true },
                new Item { Name = "Delta", Value = 40, IsImportant = false }
            }
        };

        // Build report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save output
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public int Value { get; set; }
    public bool IsImportant { get; set; }
}
