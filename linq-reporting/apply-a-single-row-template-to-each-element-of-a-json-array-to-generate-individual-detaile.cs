using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportGenerator
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths for temporary data and documents
        const string jsonPath = "Data.json";
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // 1. Create sample JSON data (array of objects) and write it to a file
        string jsonContent = @"[
            { ""Name"": ""John Doe"", ""Age"": 30, ""Address"": ""123 Main St"" },
            { ""Name"": ""Jane Smith"", ""Age"": 25, ""Address"": ""456 Oak Ave"" },
            { ""Name"": ""Bob Johnson"", ""Age"": 40, ""Address"": ""789 Pine Rd"" }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build the template document programmatically
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("----------------");

        // Insert LINQ Reporting tags
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name   : <<[person.Name]>>");
        builder.Writeln("Age    : <<[person.Age]>>");
        builder.Writeln("Address: <<[person.Address]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // 4. Create a JsonDataSource from the JSON file
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report using the ReportingEngine
        ReportingEngine engine = new ReportingEngine();
        // No special options required for this simple scenario
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // 6. Save the generated report
        reportDoc.Save(outputPath);
    }
}
