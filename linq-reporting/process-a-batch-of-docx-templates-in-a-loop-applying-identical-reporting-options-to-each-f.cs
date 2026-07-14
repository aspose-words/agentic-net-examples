using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings used by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model that will be used for all reports.
        Order sampleOrder = new Order
        {
            CustomerName = "Acme Corp",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Widget" },
                new Item { Index = 2, Name = "Gadget" },
                new Item { Index = 3, Name = "Doohickey" }
            }
        };

        // Create folders for templates and generated reports.
        string templatesDir = Path.Combine(Environment.CurrentDirectory, "Templates");
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Outputs");
        Directory.CreateDirectory(templatesDir);
        Directory.CreateDirectory(outputDir);

        // Generate a few identical template files programmatically.
        for (int i = 1; i <= 3; i++)
        {
            string templatePath = Path.Combine(templatesDir, $"Template{i}.docx");
            CreateTemplate(templatePath);
        }

        // Process each template file, applying the same reporting options and data.
        foreach (string templateFile in Directory.GetFiles(templatesDir, "*.docx"))
        {
            // Load the template document.
            Document doc = new Document(templateFile);

            // Configure the reporting engine with desired options.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report using the sample data model. The root name in the template is "order".
            engine.BuildReport(doc, sampleOrder, "order");

            // Save the generated report next to the template with a distinct name.
            string outputPath = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(templateFile) + "_Report.docx");
            doc.Save(outputPath);
        }
    }

    // Creates a simple LINQ Reporting template containing a foreach loop.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Header with a placeholder for the customer name.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        // Each iteration writes the item index and name.
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to the specified path.
        doc.Save(filePath);
    }
}

// Root data model for the report.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Simple item model used inside the Order.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
