using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template, JSON data and the generated report.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string jsonPath = Path.Combine(outputDir, "People.json");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // 1. Create sample JSON data.
        var people = new List<Person>
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "Bob", Age = 25 },
            new Person { Name = "Charlie", Age = 35 }
        };
        string jsonContent = System.Text.Json.JsonSerializer.Serialize(people);
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People List:");
        // foreach tag iterating over the JSON array named 'persons'.
        builder.Writeln("<<foreach [in persons]>>");
        builder.Writeln("- <<[Name]>> is <<[Age]>> years old.");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template document for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Create a JsonDataSource from the JSON file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The root name 'persons' must match the name used in the foreach tag.
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // 6. Save the generated report.
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated successfully at: {reportPath}");
    }

    // Simple data model used only for JSON serialization.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }
}
