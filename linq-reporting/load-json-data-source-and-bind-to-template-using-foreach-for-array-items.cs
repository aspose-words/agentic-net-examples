using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a sample JSON file containing an array of person objects.
        string jsonPath = "people.json";
        string jsonContent = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"", ""Age"": 25 },
            { ""Name"": ""Charlie"", ""Age"": 28 }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // Build the template document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags that will iterate over the JSON array.
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Load the JSON data source.
        JsonDataSource dataSource = new JsonDataSource(jsonPath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
