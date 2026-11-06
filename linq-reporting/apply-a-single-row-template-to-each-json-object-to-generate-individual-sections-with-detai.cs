using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths used in the example
        string templatePath = "template.docx";
        string jsonPath = "data.json";
        string outputPath = "report.docx";

        // -------------------------------------------------------------
        // 1. Create a Word template with LINQ Reporting tags.
        // -------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Persons Report");
        builder.Writeln();

        // Begin foreach over the JSON array named "Persons"
        builder.Writeln("<<foreach [person in Persons]>>");

        // Each person gets its own section (simple paragraph list)
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Email: <<[person.Email]>>");
        builder.Writeln("Address: <<[person.Address]>>");
        builder.Writeln(); // blank line between entries

        // End foreach
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------------------
        // 2. Create sample JSON data matching the template.
        // -------------------------------------------------------------
        string jsonContent = @"
{
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
        File.WriteAllText(jsonPath, jsonContent.Trim());

        // -------------------------------------------------------------
        // 3. Load the template and build the report using the JSON data source.
        // -------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Create a JSON data source pointing to the file
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report; the root name used in the template is "Persons"
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "Persons");

        // -------------------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------------------
        reportDoc.Save(outputPath, SaveFormat.Docx);
    }
}
