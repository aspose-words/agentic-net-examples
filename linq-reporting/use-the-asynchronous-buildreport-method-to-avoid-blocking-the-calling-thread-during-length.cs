using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Entry point – asynchronous to avoid blocking while the report is being built.
    public static async Task Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare file paths.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string reportPath   = Path.Combine(Environment.CurrentDirectory, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple template that lists order items.
        builder.Writeln("Order for <<[order.CustomerName]>>:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (simulating a real‑world scenario).
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
        // 4. Build the report asynchronously.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Wrap the synchronous BuildReport call in Task.Run to avoid blocking.
        bool success = await Task.Run(() => engine.BuildReport(doc, sampleOrder, "order"));

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        if (success)
        {
            doc.Save(reportPath);
            Console.WriteLine($"Report generated successfully: {reportPath}");
        }
        else
        {
            Console.WriteLine("Report generation failed.");
        }
    }
}

// ---------------------------------------------------------------------
// Data model – must be public with public properties for LINQ Reporting.
// ---------------------------------------------------------------------
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
