using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare file paths
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string jsonPath = Path.Combine(workDir, "Data.json");
        string reportPath = Path.Combine(workDir, "Report.docx");

        // Create sample JSON data (array of objects)
        string jsonContent = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""Charlie"", ""Age"": 28 }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // Build the template document with LINQ Reporting tags
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [p in data]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        template.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Load JSON data source
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "data");

        // Save the generated report
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated: {reportPath}");
    }
}
