using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class JsonWhereExample
{
    static void Main()
    {
        // ---------- Prepare temporary files ----------
        string tempDir = Path.GetTempPath();

        // JSON data source
        string jsonPath = Path.Combine(tempDir, "people.json");
        string jsonContent = @"[
            { ""Name"": ""John"",  ""Age"": 45, ""City"": ""London"" },
            { ""Name"": ""Anna"",  ""Age"": 28, ""City"": ""Paris"" },
            { ""Name"": ""Mike"",  ""Age"": 34, ""City"": ""London"" },
            { ""Name"": ""Laura"", ""Age"": 22, ""City"": ""Berlin"" }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // Word template containing the Reporting Engine expression
        string templatePath = Path.Combine(tempDir, "ReportTemplate.docx");
        Document templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        // The expression filters persons where Age > 30 and City == "London",
        // then joins their names with commas.
        builder.Writeln("{{persons.Where(p => p.Age > 30 && p.City == \"London\").Select(p => p.Name).Join(\", \")}}");
        templateDoc.Save(templatePath);

        // ---------- Load the template ----------
        Document doc = new Document(templatePath);

        // ---------- Configure JSON parsing ----------
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            SimpleValueParseMode = JsonSimpleValueParseMode.Loose,
            PreserveSpaces = true
        };

        // ---------- Create the JSON data source ----------
        JsonDataSource jsonData = new JsonDataSource(jsonPath, loadOptions);

        // ---------- Build the report ----------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, jsonData, "persons");

        // ---------- Save the generated document ----------
        string outputPath = Path.Combine(tempDir, "FilteredReport.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
