using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // Contains ReportingEngine, JsonDataSource, XmlDataSource

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for temporary files
        string templatePath = "Template.docx";
        string xmlPath = "Orders.xml";
        string jsonPath = "Products.json";
        string outputPath = "ReportOutput.docx";

        // -----------------------------------------------------------------
        // 1. Create sample XML data source file
        // -----------------------------------------------------------------
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <OrderId>1</OrderId>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <OrderId>2</OrderId>
    </Order>
</Orders>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Create sample JSON data source file
        // -----------------------------------------------------------------
        string jsonContent = @"{
    ""products"": [
        { ""Name"": ""Apple"",  ""Price"": 1.20 },
        { ""Name"": ""Banana"", ""Price"": 0.80 },
        { ""Name"": ""Orange"", ""Price"": 1.50 }
    ]
}";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 3. Build the template document with LINQ Reporting tags
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Multi‑Section Report ===");
        builder.Writeln();

        // Outer foreach over XML orders
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Products:");

        // Inner foreach over JSON products
        builder.Writeln("<<foreach [product in products]>>");
        builder.Writeln("- <<[product.Name]>> : $<<[product.Price]>>");
        builder.Writeln("<</foreach>>"); // end inner foreach

        builder.Writeln("<</foreach>>"); // end outer foreach

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load data sources
        // -----------------------------------------------------------------
        var xmlDataSource = new XmlDataSource(xmlPath);
        var jsonDataSource = new JsonDataSource(jsonPath);

        // -----------------------------------------------------------------
        // 5. Load the template and build the report
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // BuildReport with multiple data sources: orders (XML) and products (JSON)
        engine.BuildReport(reportDoc,
            new object[] { xmlDataSource, jsonDataSource },
            new[] { "orders", "products" });

        // -----------------------------------------------------------------
        // 6. Save the final report
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
