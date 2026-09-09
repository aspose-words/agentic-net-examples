using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Paths for temporary files
        string jsonPath = "orders.json";
        string templatePath = "template.docx";
        string outputPath = "Report.docx";

        // 1. Create sample JSON data
        var orders = new List<Order>
        {
            new Order { Price = 10.5m, Quantity = 3 },
            new Order { Price = 7.2m,  Quantity = 5 }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(orders));

        // 2. Build the template document with LINQ Reporting tags
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the JSON array
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Price: <<[order.Price]>>");
        builder.Writeln("Quantity: <<[order.Quantity]>>");
        // Calculated field: Price * Quantity
        builder.Writeln("Total: <<[order.Price * order.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // 3. Load the template document (as required before BuildReport)
        var loadedTemplate = new Document(templatePath);

        // 4. Create a JSON data source
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report
        ReportingEngine engine = new ReportingEngine();
        // The root name "orders" must match the name used in the template tags
        engine.BuildReport(loadedTemplate, jsonDataSource, "orders");

        // 6. Save the generated report
        loadedTemplate.Save(outputPath);
    }

    // Simple POCO matching the JSON structure (used only for creating sample JSON)
    public class Order
    {
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }
}
