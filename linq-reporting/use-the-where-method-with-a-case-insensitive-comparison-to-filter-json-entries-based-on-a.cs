using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data (array of objects with Name and Age)
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "people.json");
        string jsonContent = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""alice"", ""Age"": 28 },
            { ""Name"": ""Charlie"", ""Age"": 35 }
        ]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // Create a template document programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Write a heading
        builder.Writeln("People whose name is 'alice' (case‑insensitive):");

        // LINQ Reporting tag: iterate over all persons and filter with an IF tag
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("<<if [person.Name.ToLower() == \"alice\"]>>- <<[person.Name]>> (Age: <<[person.Age]>>)<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Create a JsonDataSource from the JSON file
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report using the data source; the root name is "persons"
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // Save the generated report
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
