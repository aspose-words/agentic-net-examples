using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Xml.Serialization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Generate large data set
        List<Order> orders = new();
        for (int i = 0; i < 2000; i++)
        {
            var order = new Order
            {
                CustomerName = $"Customer {i}",
                Items = new()
                {
                    new Item { Name = "Item A", Quantity = i % 5 + 1 },
                    new Item { Name = "Item B", Quantity = (i + 2) % 5 + 1 }
                }
            };
            orders.Add(order);
        }

        // Optional: serialize to XML (demonstrates XML handling)
        string xmlPath = Path.Combine(outputDir, "orders.xml");
        var serializer = new XmlSerializer(typeof(List<Order>), new XmlRootAttribute("Orders"));
        using (var stream = new FileStream(xmlPath, FileMode.Create, FileAccess.Write))
        {
            serializer.Serialize(stream, orders);
        }

        // Create LINQ Reporting template
        string templatePath = Path.Combine(outputDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report of Orders");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order: <<[order.CustomerName]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Benchmark without reflection optimization
        ReportingEngine.UseReflectionOptimization = false;
        var docWithoutOpt = new Document(templatePath);
        var engineWithout = new ReportingEngine();
        var swWithout = Stopwatch.StartNew();
        bool successWithout = engineWithout.BuildReport(docWithoutOpt, new ReportModel { Orders = orders }, "model");
        swWithout.Stop();
        string reportWithoutPath = Path.Combine(outputDir, "report_without_optimization.docx");
        docWithoutOpt.Save(reportWithoutPath);

        // Benchmark with reflection optimization
        ReportingEngine.UseReflectionOptimization = true;
        var docWithOpt = new Document(templatePath);
        var engineWith = new ReportingEngine();
        var swWith = Stopwatch.StartNew();
        bool successWith = engineWith.BuildReport(docWithOpt, new ReportModel { Orders = orders }, "model");
        swWith.Stop();
        string reportWithPath = Path.Combine(outputDir, "report_with_optimization.docx");
        docWithOpt.Save(reportWithPath);

        // Output results
        Console.WriteLine($"Report without optimization: {(successWithout ? "Success" : "Failed")} in {swWithout.ElapsedMilliseconds} ms");
        Console.WriteLine($"Report with optimization:    {(successWith ? "Success" : "Failed")} in {swWith.ElapsedMilliseconds} ms");
        Console.WriteLine($"Outputs saved to: {outputDir}");
    }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}
