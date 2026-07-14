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
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a reusable template.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Simulate two separate user requests with isolated ReportingEngine instances.
        var user1 = new ReportModel
        {
            UserName = "Alice",
            Items = new()
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" }
            }
        };

        var user2 = new ReportModel
        {
            UserName = "Bob",
            Items = new()
            {
                new Item { Index = 1, Name = "Carrot" },
                new Item { Index = 2, Name = "Date" },
                new Item { Index = 3, Name = "Eggplant" }
            }
        };

        GenerateReportForUser(templatePath, user1);
        GenerateReportForUser(templatePath, user2);
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Header with user name.
        builder.Writeln("Report for user: <<[model.UserName]>>");
        builder.Writeln();

        // Items table using foreach.
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in Items]>>");

        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    private static void GenerateReportForUser(string templatePath, ReportModel model)
    {
        // Load the template for this request.
        var doc = new Document(templatePath);

        // Create a new ReportingEngine instance for isolation.
        var engine = new ReportingEngine();

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = $"Report_{model.UserName}.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report for {model.UserName} generated: {(success ? "Success" : "Failed")} -> {outputPath}");
    }
}

public class ReportModel
{
    public string UserName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}
