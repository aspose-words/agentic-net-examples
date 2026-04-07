using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // 2. Insert LINQ Reporting tags.
        // The foreach tag iterates over the Items collection of the root data source.
        // No generic type parameters are needed – the engine determines the item type at runtime.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Index: <<[item.Index]>>, Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // 3. Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" }
            }
        };

        // 4. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple example.
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(template, model, "model");

        // 5. Save the generated report.
        template.Save("Report.docx");
    }
}

// Root data model exposed to the template.
public class ReportModel
{
    // The collection that the foreach tag will iterate over.
    public List<Item> Items { get; set; } = new();
}

// Items in the collection. The engine will infer this type automatically.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
