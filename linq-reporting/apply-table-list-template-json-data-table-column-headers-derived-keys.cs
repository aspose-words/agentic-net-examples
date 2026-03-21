using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Prepare JSON data in a temporary file.
        string json = @"{ ""Name"": ""John Doe"" }";
        string jsonPath = Path.Combine(Path.GetTempPath(), "data.json");
        File.WriteAllText(jsonPath, json);

        // Create a simple Word template in memory with a reporting tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Hello <<[data.Name]>>!");

        // Load the JSON data source from the temporary file.
        JsonDataSource jsonSource = new JsonDataSource(jsonPath);

        // Build the report using the data source name "data".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, jsonSource, "data");

        // Save the populated document.
        template.Save("Result.docx");
    }
}
