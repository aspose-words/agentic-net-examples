using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting; // Ensure reporting namespace is available

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data
        string jsonPath = "addresses.json";
        string jsonContent = @"{
            ""Addresses"": [
                { ""Street"": ""123 Main St"", ""City"": ""Springfield"", ""State"": ""IL"", ""Zip"": ""62704"" },
                { ""Street"": ""456 Oak Ave"", ""City"": ""Metropolis"", ""State"": ""NY"", ""Zip"": ""10001"" },
                { ""Street"": ""789 Pine Rd"", ""City"": ""Gotham"", ""State"": ""NJ"", ""Zip"": ""07097"" }
            ]
        }";
        File.WriteAllText(jsonPath, jsonContent);

        // Create a template document programmatically
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Address List:");
        // Loop over the Addresses collection in the JSON root object named 'data'
        builder.Writeln("<<foreach [addr in data.Addresses]>>");
        // Build full address line using inline string concatenation
        builder.Writeln("<<[addr.Street + \", \" + addr.City + \", \" + addr.State + \" \" + addr.Zip]>>");
        builder.Writeln("<</foreach>>");

        // Configure JSON data source with options to always generate a root object
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, loadOptions);

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "data");

        // Save the generated report
        doc.Save("Report.docx");
    }
}
