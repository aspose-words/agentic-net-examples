using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReflectionOptimizationExample
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Large Data Set Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>>  Value: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template (demonstrates load step).
        Document doc = new Document(templatePath);

        // Prepare a large data set.
        ReportModel model = new ReportModel();
        for (int i = 1; i <= 1000; i++)
        {
            model.Items.Add(new Item
            {
                Name = $"Item{i}",
                Value = i * 10
            });
        }

        // Enable reflection call optimization (caching) for the reporting engine.
        ReportingEngine.UseReflectionOptimization = true;

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "ReportOutput.docx";
        doc.Save(outputPath);
    }
}

// Root data model referenced by the template.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class used in the collection.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}
