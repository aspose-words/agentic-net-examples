using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Paths for temporary files.
        const string jsonFilePath = "data.json";
        const string reportFilePath = "Report.docx";

        // 1. Create sample JSON data.
        CreateSampleJsonFile(jsonFilePath);

        // 2. Build a Word template in memory with LINQ Reporting tags.
        Document template = CreateTemplateDocument();

        // 3. Load JSON data from a file stream into a JsonDataSource.
        using (FileStream jsonStream = File.OpenRead(jsonFilePath))
        {
            JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, jsonDataSource, "persons");
        }

        // 5. Save the generated report to a local file.
        template.Save(reportFilePath, SaveFormat.Docx);

        // 6. Simulate uploading the report to Azure Blob Storage.
        // In a real scenario you would use Azure.Storage.Blobs SDK here.
        SimulateUploadToAzureBlob(reportFilePath);
    }

    // Creates a simple JSON file containing an array of person objects.
    private static void CreateSampleJsonFile(string path)
    {
        string jsonContent = @"[
    { ""Name"": ""Alice"", ""Age"": 30 },
    { ""Name"": ""Bob"",   ""Age"": 25 },
    { ""Name"": ""Charlie"", ""Age"": 28 }
]";
        File.WriteAllText(path, jsonContent, Encoding.UTF8);
    }

    // Generates a Word document with LINQ Reporting tags that iterate over the JSON data.
    private static Document CreateTemplateDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Person Report");
        builder.ParagraphFormat.ClearFormatting();

        // Begin a foreach loop over the "persons" collection.
        builder.Writeln("<<foreach [person in persons]>>");
        // Output each person's name and age.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age:  <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        return doc;
    }

    // Simulates uploading a file to Azure Blob Storage.
    private static void SimulateUploadToAzureBlob(string filePath)
    {
        // This method only demonstrates the intended action.
        // Replace its body with actual Azure Blob upload code using Azure.Storage.Blobs SDK.
        Console.WriteLine($"[Simulation] Uploaded '{filePath}' to Azure Blob Storage.");
    }
}
