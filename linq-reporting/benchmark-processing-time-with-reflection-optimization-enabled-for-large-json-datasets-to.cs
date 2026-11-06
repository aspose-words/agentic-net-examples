using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template, data and output
        const string templatePath = "template.docx";
        const string jsonPath = "data.json";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Use a property on the root object to display the current date/time.
        builder.Writeln("Report generated at: <<[Now]>>");
        builder.Writeln();
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>   Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>   Price: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Generate a large JSON dataset and save it
        // -----------------------------------------------------------------
        const int orderCount = 1000;   // number of orders
        const int itemsPerOrder = 100; // items per order

        var root = new RootModel();

        for (int o = 1; o <= orderCount; o++)
        {
            var order = new Order
            {
                Id = o,
                CustomerName = $"Customer {o}"
            };

            for (int i = 1; i <= itemsPerOrder; i++)
            {
                order.Items.Add(new Item
                {
                    Index = i,
                    Name = $"Product {i}",
                    Price = Math.Round((decimal)(i * 1.23), 2)
                });
            }

            root.Orders.Add(order);
        }

        // Serialize to JSON and write to file
        string json = JsonConvert.SerializeObject(root, Formatting.Indented);
        File.WriteAllText(jsonPath, json, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 3. Load the template and JSON data, then benchmark report generation
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Deserialize JSON back to object model (simulating a large JSON load)
        string jsonFromFile = File.ReadAllText(jsonPath, Encoding.UTF8);
        var dataRoot = JsonConvert.DeserializeObject<RootModel>(jsonFromFile)!;

        // Enable reflection optimization
        ReportingEngine.UseReflectionOptimization = true;

        ReportingEngine engine = new ReportingEngine();

        Stopwatch sw = Stopwatch.StartNew();
        bool success = engine.BuildReport(doc, dataRoot, "root");
        sw.Stop();

        // Save the generated report
        doc.Save(outputPath);

        // Output benchmark result
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Elapsed time with reflection optimization: {sw.ElapsedMilliseconds} ms");
    }
}

// ---------------------------------------------------------------------
// Data model classes (must be public with public properties)
// ---------------------------------------------------------------------
public class RootModel
{
    // Property used in the template to display the current date/time.
    public DateTime Now => DateTime.Now;

    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}
