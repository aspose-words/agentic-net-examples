using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths.
        string jsonPath = "people.json";
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // 1. Create sample JSON data (an array of person objects).
        string jsonContent = @"
[
  { ""Name"": ""Alice"", ""Age"": 30, ""Address"": ""123 Main St"" },
  { ""Name"": ""Bob"",   ""Age"": 25, ""Address"": ""456 Oak Ave"" },
  { ""Name"": ""Carol"", ""Age"": 28, ""Address"": ""789 Pine Rd"" }
]";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("People Report");
        builder.Writeln();

        // Header table (static header row).
        Table headerTable = builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("Address");
        builder.EndRow();

        builder.EndTable();

        // Begin the foreach block that iterates over the JSON array.
        builder.Writeln("<<foreach [person in persons]>>");

        // Table that will be repeated for each person (single data row).
        Table dataTable = builder.StartTable();

        builder.InsertCell();
        builder.Writeln("<<[person.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Address]>>");
        builder.EndRow();

        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template document for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Create a JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name used in the template tags is "persons".
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // 6. Save the generated report.
        reportDoc.Save(outputPath);
    }
}
