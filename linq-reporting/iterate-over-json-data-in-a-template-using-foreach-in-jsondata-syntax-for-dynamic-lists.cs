using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directories
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string jsonPath = Path.Combine(workDir, "data.json");
        string outputPath = Path.Combine(workDir, "output.docx");

        // ---------- Create JSON data ----------
        string jsonContent = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""Charlie"", ""Age"": 28 }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // ---------- Create template document ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People List:");
        builder.Writeln("<<foreach [person in jsonData]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // ---------- Load template for reporting ----------
        Document reportDoc = new Document(templatePath);

        // Load JSON data source
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonData, "jsonData");

        // Save the generated report
        reportDoc.Save(outputPath);
    }
}
