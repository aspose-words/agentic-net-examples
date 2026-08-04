using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Path for the template document.
        const string templatePath = "template.docx";

        // Create the template with LINQ Reporting tags.
        CreateTemplate(templatePath);

        // First data set.
        var model1 = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Price = 1.20 },
                new Item { Name = "Banana", Price = 0.80 }
            }
        };

        // Second data set.
        var model2 = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Carrot", Price = 0.50 },
                new Item { Name = "Date", Price = 2.00 }
            }
        };

        // Single ReportingEngine instance enables caching of the compiled template.
        var engine = new ReportingEngine();

        // Build first report.
        var doc1 = new Document(templatePath);
        engine.BuildReport(doc1, model1, "model");
        doc1.Save("Report1.docx");

        // Build second report using the same engine (cached template is reused).
        var doc2 = new Document(templatePath);
        engine.BuildReport(doc2, model2, "model");
        doc2.Save("Report2.docx");
    }

    // Creates a simple Word template containing a foreach loop.
    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Items Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    // Root data model referenced in the template as <<[model]>>.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Simple item class used in the collection.
    public class Item
    {
        public string Name { get; set; } = "";
        public double Price { get; set; }
    }
}
