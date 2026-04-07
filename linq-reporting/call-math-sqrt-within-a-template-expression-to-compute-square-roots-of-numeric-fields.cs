using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel();
        model.Numbers.Add(new Item { Value = 4 });
        model.Numbers.Add(new Item { Value = 9 });
        model.Numbers.Add(new Item { Value = 16 });
        model.Numbers.Add(new Item { Value = 2.5 });

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Square Roots Report");
        builder.Writeln("<<foreach [item in Numbers]>>");
        // Use Math.Sqrt via KnownTypes registration.
        builder.Writeln("Original: <<[item.Value]>>   Sqrt: <<[Math.Sqrt(item.Value)]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        // Register System.Math so its static members can be used in expressions.
        engine.KnownTypes.Add(typeof(Math));
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Wrapper class that serves as the root data source for the template.
public class ReportModel
{
    // Collection of numeric items.
    public List<Item> Numbers { get; set; } = new();
}

// Simple data entity containing a numeric value.
public class Item
{
    public double Value { get; set; }
}
