using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "SampleItem";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel();
        model.Items.Add(new Item { Name = "First" });
        model.Items.Add(new Item { Name = "Second" });

        // Create a template document with LINQ Reporting tags.
        var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Item count: <<[model.Items.Count]>>");

        // Attempt to access a modifying method (Add). This will be blocked by RestrictedTypes.
        // The expression is syntactically valid but refers to a restricted member, so it will be replaced
        // by the MissingMemberMessage ("Restricted").
        builder.Writeln("Attempt to add: <<[model.Items.Add]>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Restrict access to List<Item> members (e.g., Add, Remove) to prevent modifications.
        ReportingEngine.SetRestrictedTypes(typeof(List<Item>));

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "Restricted";

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        reportDoc.Save(outputPath);

        // Output the resulting text to the console.
        Console.WriteLine(reportDoc.GetText());
    }
}
