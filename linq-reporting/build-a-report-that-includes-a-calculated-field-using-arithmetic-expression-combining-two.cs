using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

// Register code page provider for Aspose.Words if needed.
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public class ReportGenerator
{
    public static void Main()
    {
        // Paths for the template and output documents.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";
        const string jsonPath = "Data.json";

        // 1. Create sample JSON data.
        var sampleData = new DataRoot
        {
            Orders = new()
            {
                new Order { Id = 1, Price = 19.99m, Quantity = 2, Tax = 3.20m },
                new Order { Id = 2, Price = 5.50m, Quantity = 5, Tax = 1.10m },
                new Order { Id = 3, Price = 12.00m, Quantity = 1, Tax = 0.96m }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // 2. Build the LINQ Reporting template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln("==============");
        builder.Writeln();

        // Table header.
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.InsertCell();
        builder.Writeln("Tax");
        builder.InsertCell();
        builder.Writeln("Total (Price*Qty + Tax)");
        builder.EndRow();

        // Data rows using foreach tag.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Price]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Quantity]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Tax]>>");
        builder.InsertCell();
        // Calculated field: Price * Quantity + Tax
        builder.Writeln("<<[order.Price * order.Quantity + order.Tax]>>");
        builder.EndRow();
        builder.Writeln("<</foreach>>");

        builder.EndTable();

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template document.
        var doc = new Document(templatePath);

        // 4. Load JSON data into the model.
        var jsonContent = File.ReadAllText(jsonPath);
        var model = JsonConvert.DeserializeObject<DataRoot>(jsonContent)!;

        // 5. Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 6. Save the generated report.
        doc.Save(outputPath);
    }
}

// Root object that matches the JSON structure.
public class DataRoot
{
    public List<Order> Orders { get; set; } = new();
}

// Order model with numeric properties.
public class Order
{
    public int Id { get; set; }
    public decimal Price { get; set; }
    public int Quantity { get; set; }
    public decimal Tax { get; set; }
}
