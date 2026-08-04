using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // 1. Create sample JSON data.
        string jsonPath = Path.Combine(outputDir, "people.json");
        string jsonContent = @"
[
  {
    ""FirstName"": ""John"",
    ""LastName"": ""Doe"",
    ""Street"": ""123 Main St"",
    ""City"": ""Springfield"",
    ""State"": ""IL"",
    ""Zip"": ""62704""
  },
  {
    ""FirstName"": ""Jane"",
    ""LastName"": ""Smith"",
    ""Street"": ""456 Oak Ave"",
    ""City"": ""Metropolis"",
    ""State"": ""NY"",
    ""Zip"": ""10001""
  }
]";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build the template document programmatically.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("Address Report");
        builder.Writeln();

        // Begin a foreach loop over the JSON collection named 'persons'.
        builder.Writeln("<<foreach [person in persons]>>");

        // Write a line with the full address using inline string concatenation.
        builder.Writeln(
            "<<[person.FirstName + \" \" + person.LastName + \": \" + person.Street + \", \" + person.City + \", \" + person.State + \" \" + person.Zip]>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template document (simulating a separate load step).
        Document reportDoc = new Document(templatePath);

        // 4. Create a JsonDataSource from the JSON file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // 6. Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated successfully at: {reportPath}");
    }
}
