using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model for the report.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header with the customer's name.
        builder.Writeln("Order for <<[order.CustomerName]>>");
        builder.Writeln();

        // Loop over the collection of items.
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (simulating a real-world scenario).
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data.
        // -----------------------------------------------------------------
        Order sampleOrder = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // -----------------------------------------------------------------
        // 4. Enable reflection optimization and build the report.
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = true; // Enable runtime proxy generation.
        ReportingEngine engine = new ReportingEngine();

        // The root object name in the template is "order".
        engine.BuildReport(doc, sampleOrder, "order");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }
}
