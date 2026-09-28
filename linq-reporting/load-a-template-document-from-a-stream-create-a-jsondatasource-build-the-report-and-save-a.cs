using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document in memory.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a simple LINQ Reporting foreach loop over "orders".
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a memory stream.
        using var templateStream = new MemoryStream();
        templateDoc.Save(templateStream, SaveFormat.Docx);
        templateStream.Position = 0; // Reset before loading.

        // Load the template document from the stream.
        var reportDoc = new Document(templateStream);

        // Prepare sample JSON data with a root array named "orders".
        string json = @"{
            ""orders"": [
                { ""CustomerName"": ""Alice"", ""Amount"": 123.45 },
                { ""CustomerName"": ""Bob"",   ""Amount"": 67.89 }
            ]
        }";

        // Create a JsonDataSource from the JSON string.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        var jsonDataSource = new JsonDataSource(jsonStream);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "orders");

        // Save the generated report as RTF.
        reportDoc.Save("Report.rtf");
    }
}
