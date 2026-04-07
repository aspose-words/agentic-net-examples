using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class JsonLinqReportingExample
{
    public static void Main()
    {
        // Register code page provider for any encoding needs (required by Aspose.Words on some platforms).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string jsonPath = "people.json";
        string templatePath = "template.docx";
        string outputPath = "report.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample JSON file containing an array of person objects.
        // -----------------------------------------------------------------
        string jsonContent = @"[
  { ""Name"": ""John Doe"", ""Age"": 30 },
  { ""Name"": ""Jane Smith"", ""Age"": 25 },
  { ""Name"": ""Bob Johnson"", ""Age"": 40 }
]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // ---------------------------------------------------------------
        // 2. Build a Word template programmatically and insert LINQ tags.
        // ---------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template back (demonstrates the required load step).
        // ---------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // ---------------------------------------------------------------
        // 4. Create a JsonDataSource from the JSON file.
        // ---------------------------------------------------------------
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // ---------------------------------------------------------------
        // 5. Build the report using ReportingEngine.
        // ---------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // The root name "persons" matches the name used in the template tags.
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // ---------------------------------------------------------------
        // 6. Save the generated report.
        // ---------------------------------------------------------------
        reportDoc.Save(outputPath);

        // Inform the user (optional, no interactive input required).
        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}
