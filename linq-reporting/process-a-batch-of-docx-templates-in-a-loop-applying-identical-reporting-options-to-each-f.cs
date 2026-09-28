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
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders for templates and output reports.
        string baseDir = Directory.GetCurrentDirectory();
        string templateDir = Path.Combine(baseDir, "Templates");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(templateDir);
        Directory.CreateDirectory(outputDir);

        // Create three identical template files programmatically.
        for (int i = 1; i <= 3; i++)
        {
            string templatePath = Path.Combine(templateDir, $"Template{i}.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Report for <<[order.CustomerName]>>");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);
        }

        // Sample data model.
        Order order = new Order
        {
            CustomerName = "Acme Corp",
            Items = new List<Item>
            {
                new() { Index = 1, Name = "Widget" },
                new() { Index = 2, Name = "Gadget" },
                new() { Index = 3, Name = "Doohickey" }
            }
        };

        // Configure the reporting engine.
        ReportingEngine.UseReflectionOptimization = true;
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Process each template file, generate a report, and save it.
        foreach (string templateFile in Directory.GetFiles(templateDir, "*.docx"))
        {
            Document reportDoc = new Document(templateFile);
            bool success = engine.BuildReport(reportDoc, order, "order");
            // success indicates whether the report was built without errors when InlineErrorMessages is set.

            string outputPath = Path.Combine(outputDir, $"Report_{Path.GetFileName(templateFile)}");
            reportDoc.Save(outputPath);
        }
    }
}

// Data model classes.
public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}
