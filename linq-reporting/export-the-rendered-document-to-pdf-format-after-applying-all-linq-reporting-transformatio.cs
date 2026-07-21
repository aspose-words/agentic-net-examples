using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags into the template.
        builder.Writeln("Report for <<[model.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Build a sample data model.
        ReportModel model = new()
        {
            CustomerName = "Acme Corporation",
            Items = new()
            {
                new() { Name = "Widget", Quantity = 12 },
                new() { Name = "Gadget", Quantity = 7 },
                new() { Name = "Doohickey", Quantity = 3 }
            }
        };

        // (Optional) Serialize the model to JSON to demonstrate the Newtonsoft.Json dependency.
        string json = JsonConvert.SerializeObject(model);
        Console.WriteLine("Sample data model serialized to JSON:");
        Console.WriteLine(json);

        // Create the reporting engine and build the report.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // If the build succeeded, save the rendered document as PDF.
        if (success)
        {
            doc.Save("Report.pdf", SaveFormat.Pdf);
            Console.WriteLine("Report generated and saved as Report.pdf");
        }
        else
        {
            Console.WriteLine("Report generation failed.");
        }
    }
}

// Data model classes used by the LINQ Reporting engine.
public class ReportModel
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}
