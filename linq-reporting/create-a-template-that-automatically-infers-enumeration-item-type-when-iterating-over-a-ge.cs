using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel();
        model.Items.Add(new Item { Index = 1, Name = "Alpha" });
        model.Items.Add(new Item { Index = 2, Name = "Beta" });
        model.Items.Add(new Item { Index = 3, Name = "Gamma" });

        // Create the template document with LINQ Reporting tags
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Item List:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Index: <<[item.Index]>>, Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for report generation
        var reportDoc = new Document(templatePath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
