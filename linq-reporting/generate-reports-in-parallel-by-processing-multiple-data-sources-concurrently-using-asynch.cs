using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static async Task Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory("Output");

        // Create a simple LINQ Reporting template programmatically.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Prepare two distinct data sources.
        ReportModel model1 = new()
        {
            Title = "First Report",
            Items = new()
            {
                new Item { Name = "Alpha", Value = 100 },
                new Item { Name = "Beta", Value = 200 }
            }
        };

        ReportModel model2 = new()
        {
            Title = "Second Report",
            Items = new()
            {
                new Item { Name = "Gamma", Value = 300 },
                new Item { Name = "Delta", Value = 400 }
            }
        };

        // Load separate document instances for each report.
        Document doc1 = new(templatePath);
        Document doc2 = new(templatePath);

        // Configure the reporting engine.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;

        // Run report generation in parallel.
        Task<bool> task1 = Task.Run(() => engine.BuildReport(doc1, model1, "model"));
        Task<bool> task2 = Task.Run(() => engine.BuildReport(doc2, model2, "model"));

        await Task.WhenAll(task1, task2);

        // Save the generated reports.
        doc1.Save(Path.Combine("Output", "Report1.docx"));
        doc2.Save(Path.Combine("Output", "Report2.docx"));
    }

    private static void CreateTemplate(string path)
    {
        Document template = new();
        DocumentBuilder builder = new(template);

        // Title placeholder.
        builder.Writeln("Report Title: <<[model.Title]>>");

        // Loop over items.
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("- <<[item.Name]>> : <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        template.Save(path);
    }
}

public class ReportModel
{
    public string Title { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public int Value { get; set; }
}
