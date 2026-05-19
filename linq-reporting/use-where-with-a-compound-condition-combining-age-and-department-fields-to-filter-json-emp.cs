using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "employees.json";
        string jsonContent = @"[
            { ""Name"": ""Alice"",   ""Age"": 28, ""Department"": ""HR"" },
            { ""Name"": ""Bob"",     ""Age"": 35, ""Department"": ""Sales"" },
            { ""Name"": ""Charlie"", ""Age"": 42, ""Department"": ""Sales"" },
            { ""Name"": ""Diana"",   ""Age"": 31, ""Department"": ""IT"" }
        ]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Filtered Employees (Age > 30 && Department == \"Sales\"):");
        builder.Writeln("<<foreach [emp in employees]>>");
        builder.Writeln("<<if [emp.Age > 30 && emp.Department == \"Sales\"]>><<[emp.Name]>> - <<[emp.Age]>> - <<[emp.Department]>> <</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template back (demonstrates load step).
        Document doc = new Document(templatePath);

        // Create a JSON data source.
        JsonDataSource dataSource = new JsonDataSource(jsonPath);

        // Build the report using the data source. The root name "employees" matches the tag in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "employees");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
