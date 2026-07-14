using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        // Output the original value and its square root using Math.Sqrt.
        builder.Writeln("Value: <<[item.Value]>>, Sqrt: <<[Math.Sqrt(item.Value)]>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk as required by the lifecycle rules.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Load the template back for report generation.
        var reportDoc = new Document(templatePath);

        // Step 3: Prepare the data model.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Value = 4.0 },
                new Item { Value = 9.0 },
                new Item { Value = 16.0 }
            }
        };

        // Step 4: Configure the ReportingEngine.
        var engine = new ReportingEngine();
        // Register the Math class so its static members can be used in expressions.
        engine.KnownTypes.Add(typeof(Math));

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Step 5: Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model classes must be public with public properties.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public double Value { get; set; }
}
