using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data with a root object named "order".
        string jsonContent = @"
{
  ""order"": {
    ""Items"": [
      { ""Name"": ""Apple"",  ""Price"": 0.5, ""Quantity"": 10 },
      { ""Name"": ""Banana"", ""Price"": 0.3, ""Quantity"": 15 },
      { ""Name"": ""Cherry"", ""Price"": 1.2, ""Quantity"": 5 }
    ]
  }
}";
        File.WriteAllText("data.json", jsonContent);

        // Create a Word template programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Order Report");
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.InsertCell();
        builder.Writeln("Total");
        builder.EndRow();

        // Table rows (repeated for each item).
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Price]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        // Calculated field: Price * Quantity.
        builder.InsertCell();
        builder.Writeln("<<[item.Price * item.Quantity]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save("template.docx");

        // Load the template for reporting.
        Document doc = new Document("template.docx");

        // Load JSON data source from the file path (not the raw JSON string).
        JsonDataSource dataSource = new JsonDataSource("data.json");

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "order");

        // Save the generated report.
        doc.Save("report.docx");

        Console.WriteLine("Report generated: report.docx");
    }
}
