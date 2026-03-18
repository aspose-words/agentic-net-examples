using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class JsonReportExample
{
    static void Main()
    {
        // 1. Prepare a JSON file that contains an array of orders.
        // Each order has Item (string), Quantity (int) and Price (double).
        string jsonPath = "orders.json";
        string jsonContent = @"
        [
            { ""Item"": ""Pen"",   ""Quantity"": 2, ""Price"": 1.5 },
            { ""Item"": ""Notebook"", ""Quantity"": 1, ""Price"": 3.75 },
            { ""Item"": ""Eraser"",   ""Quantity"": 5, ""Price"": 0.5 }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Create a Word template document in memory.
        // The template uses Reporting Engine syntax:
        //   <<foreach [orders]>> … <</foreach>> iterates over the JSON array.
        //   Inline arithmetic expressions such as <<[Quantity] * [Price]>> calculate per‑row amount.
        //   The aggregate function sum() calculates the total of all amounts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("Item\tQty\tPrice\tAmount");
        builder.Writeln("<<foreach [orders]>>");
        builder.Writeln("<<[Item]>>\t<<[Quantity]>>\t<<[Price]>>\t<<[Quantity] * [Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("-------------------------------------------------");
        builder.Writeln("Total:\t\t\t<<sum([orders].[Quantity] * [orders].[Price])>>");

        // 3. Load the JSON data as a data source.
        JsonDataSource dataSource = new JsonDataSource(jsonPath);

        // 4. Build the report. The second argument is the data source object,
        //    the third argument is the name used inside the template (optional here).
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "orders");

        // 5. Save the populated document.
        doc.Save("OrderReport.docx");
    }
}
