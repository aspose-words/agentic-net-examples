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
        // Register code page provider for encodings such as Windows-1252.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder and file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string jsonPath = Path.Combine(outputDir, "Data.json");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // Sample JSON array of person objects.
        string json = @"[
  { ""Name"": ""Alice"", ""Age"": 30, ""Address"": ""123 Main St, Springfield"" },
  { ""Name"": ""Bob"", ""Age"": 45, ""Address"": ""456 Oak Ave, Shelbyville"" },
  { ""Name"": ""Charlie"", ""Age"": 28, ""Address"": ""789 Pine Rd, Capital City"" }
]";
        File.WriteAllText(jsonPath, json);

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin foreach block over the JSON array named "persons".
        builder.Writeln("<<foreach [person in persons]>>");

        // Section heading for each person.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln();

        // Table with two columns: Age and Address.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("Address");
        builder.EndRow();

        // Data row bound to the current person.
        builder.InsertCell();
        builder.Writeln("<<[person.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Address]>>");
        builder.EndRow();

        builder.EndTable();

        // Blank line between entries.
        builder.Writeln();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the report using the JSON data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        using FileStream jsonStream = File.OpenRead(jsonPath);
        JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}
