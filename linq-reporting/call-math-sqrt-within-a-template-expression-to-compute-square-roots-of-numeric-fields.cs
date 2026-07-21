using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = string.Empty;
    public double Number { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create the template document programmatically.
        const string templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("<<foreach [item in Items]>>");
        // Use Math.Sqrt within the expression to compute the square root of the numeric field.
        builder.Writeln("Name: <<[item.Name]>>, Number: <<[item.Number]>>, Sqrt: <<[Math.Sqrt(item.Number)]>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // 2. Load the template document.
        var doc = new Document(templatePath);

        // 3. Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alpha", Number = 4.0 },
                new Item { Name = "Beta",  Number = 9.0 },
                new Item { Name = "Gamma", Number = 16.0 }
            }
        };

        // 4. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        // Register System.Math so its static members can be used in template expressions.
        engine.KnownTypes.Add(typeof(Math));
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        doc.Save("Report.docx");
    }
}
