using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "John Doe";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var order = new Order
        {
            CustomerName = "Acme Corp",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Widget A" },
                new Item { Index = 2, Name = "Widget B" },
                new Item { Index = 3, Name = "Widget C" }
            }
        };

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for report generation.
        var templateDoc = new Document(templatePath);

        // Configure the reporting engine with RemoveEmptyParagraphs option.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Measure memory before report generation.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long memoryBefore = GC.GetTotalMemory(true);
        var process = Process.GetCurrentProcess();
        long privateBefore = process.PrivateMemorySize64;

        // Build the report.
        engine.BuildReport(templateDoc, order, "order");

        // Measure memory after report generation.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long memoryAfter = GC.GetTotalMemory(true);
        process.Refresh();
        long privateAfter = process.PrivateMemorySize64;

        // Save the generated report.
        var outputPath = "output.docx";
        templateDoc.Save(outputPath);

        // Output memory consumption details.
        Console.WriteLine($"GC Total Memory Increase: {memoryAfter - memoryBefore:N0} bytes");
        Console.WriteLine($"Private Memory Increase: {privateAfter - privateBefore:N0} bytes");
    }
}
