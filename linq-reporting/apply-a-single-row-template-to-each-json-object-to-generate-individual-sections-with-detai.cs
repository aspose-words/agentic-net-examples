using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths.
        string templatePath = "template.docx";
        string jsonPath = "data.json";
        string outputPath = "report.docx";

        // Create sample JSON data (array of person objects wrapped in a root object).
        string jsonContent = @"{
            ""Persons"": [
                {
                    ""Name"": ""Alice Johnson"",
                    ""Age"": 30,
                    ""Email"": ""alice.johnson@example.com"",
                    ""Address"": ""123 Maple Street, Springfield""
                },
                {
                    ""Name"": ""Bob Smith"",
                    ""Age"": 45,
                    ""Email"": ""bob.smith@example.com"",
                    ""Address"": ""456 Oak Avenue, Metropolis""
                },
                {
                    ""Name"": ""Carol Davis"",
                    ""Age"": 28,
                    ""Email"": ""carol.davis@example.com"",
                    ""Address"": ""789 Pine Road, Gotham""
                }
            ]
        }";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the JSON array named "Persons".
        builder.Writeln("<<foreach [person in Persons]>>");

        // Insert a heading for each person.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Email: <<[person.Email]>>");
        builder.Writeln("Address: <<[person.Address]>>");

        // Insert a page break after each section (optional, improves readability).
        builder.InsertBreak(BreakType.PageBreak);

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Load JSON data into a JsonDataSource using a memory stream.
        string jsonString = File.ReadAllText(jsonPath, Encoding.UTF8);
        using MemoryStream jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonString));
        JsonDataSource jsonData = new JsonDataSource(jsonStream);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonData, "Persons");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
