using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Value { get; set; }
}

public class Model
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data with at least five items.
        var model = new Model();
        model.Items.AddRange(new[]
        {
            new Item { Name = "Item1", Value = 1 },
            new Item { Name = "Item2", Value = 2 },
            new Item { Name = "Item3", Value = 3 },
            new Item { Name = "Item4", Value = 4 },
            new Item { Name = "Item5", Value = 5 },
            new Item { Name = "Item6", Value = 6 }
        });

        // Create the template document containing the ElementAt expression.
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Fifth item name: <<[model.Items.ElementAt(4).Name]>>");
        builder.Writeln("Fifth item value: <<[model.Items.ElementAt(4).Value]>>");
        templateDoc.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
