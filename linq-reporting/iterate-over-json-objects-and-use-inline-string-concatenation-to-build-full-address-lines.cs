using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data representing a list of addresses
        string json = @"
[
    {
        ""Street"": ""123 Main St"",
        ""City"": ""Springfield"",
        ""State"": ""IL"",
        ""Zip"": ""62704""
    },
    {
        ""Street"": ""456 Oak Ave"",
        ""City"": ""Metropolis"",
        ""State"": ""NY"",
        ""Zip"": ""10001""
    },
    {
        ""Street"": ""789 Pine Rd"",
        ""City"": ""Gotham"",
        ""State"": ""NJ"",
        ""Zip"": ""07097""
    }
]";
        // Write JSON to a local file
        string jsonPath = "addresses.json";
        File.WriteAllText(jsonPath, json);

        // Create a new blank document that will serve as the template
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Address Report");
        builder.Writeln();

        // LINQ Reporting tags:
        // Iterate over the JSON array (named "addresses") and concatenate address parts inline.
        builder.Writeln("<<foreach [addr in addresses]>>");
        builder.Writeln("<<[addr.Street + \", \" + addr.City + \", \" + addr.State + \" \" + addr.Zip]>>");
        builder.Writeln("<</foreach>>");

        // Create a JsonDataSource from the JSON file
        JsonDataSource dataSource = new JsonDataSource(jsonPath);

        // Build the report using the data source. The root name in the template is "addresses".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "addresses");

        // Save the generated report
        doc.Save("Report.docx");
    }
}
